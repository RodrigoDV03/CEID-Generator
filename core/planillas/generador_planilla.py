import logging
import os
import datetime
from dataclasses import dataclass, field
from typing import Callable, List, Optional

import pandas as pd
from openpyxl import load_workbook
from core.planillas.csv_processing import (
    cargar_archivo,
    filtrar_docentes_validos,
    limpiar_docentes,
    normalizar_columnas_planilla,
    procesar_csv_nuevo_formato,
)
from core.planillas import columnas
from core.planillas.calculo_montos import (
    agregar_servicio_coordinacion,
    agregar_examen_clasificacion,
    agrupar_y_calcular,
    preparar_coordinacion_agrupada,
)
from core.planillas.expansion_tablas import construir_tabla_planilla_generador_resumida
from core.planillas.planillas_anteriores import ResumenComparacion, excluir_cursos_emitidos
from core.planillas.formato_datos import formatear_numero, traducir_dias
from core.planillas import cache as cache_module
from core.planillas.transformations import aplicar_transformaciones_base
from core.planillas import table_builders
from core.planillas.excel_styles import *
from core.fases.utils import TextUtils

logger = logging.getLogger(__name__)


# Etapas visibles en la barra de progreso. La generación de planilla no
# itera fila por fila (es procesamiento vectorizado con pandas), así que el
# progreso se reporta por etapa del pipeline en vez de por docente.
ETAPAS_PLANILLA = [
    "Cargando archivo de cursos",
    "Cargando archivo de docentes y comparando con planillas anteriores",
    "Agrupando cursos por docente y calculando montos",
    "Agregando examen de clasificación",
    "Agregando servicio de coordinación",
    "Construyendo tabla de planilla",
    "Guardando hojas en el archivo Excel",
    "Aplicando formato a las hojas",
    "Ordenando hojas y guardando archivo final",
]


@dataclass
class ResultadoPlanilla:
    """
    Resultado explícito de generar la planilla, en vez del patrón anterior de
    devolver un `str` que empezaba con "❌" cuando algo fallaba (y que
    obligaba a quien llamaba a detectar el error con `.startswith("❌")`).
    """
    ok: bool
    mensaje: str
    ruta_archivo: Optional[str] = None
    error: Optional[str] = None
    sin_coincidencia: List[str] = field(default_factory=list)
    comparacion: Optional[ResumenComparacion] = None


def generar_planilla(
    data_path,
    excel_docentes,
    excel_exa_clasif,
    excel_coordinacion,
    month,
    monto_bono: float = 0,
    carpeta_destino: str | None = None,
    planillas_anteriores: Optional[List[str]] = None,
    on_progreso: Optional[Callable[[int, int, str], None]] = None,
) -> ResultadoPlanilla:
    """
    `planillas_anteriores`: CSV de las planillas ya emitidas este mes. Si se
    pasan, esta es una planilla complementaria (2ª, 3ª...) y solo incluye los
    cursos nuevos o que cambiaron de docente (ver `planillas_anteriores.py`).
    """
    año_actual = datetime.datetime.now().year
    es_enero = month.lower() == 'enero'
    total_etapas = len(ETAPAS_PLANILLA)

    def avanzar(indice_etapa: int) -> None:
        if on_progreso:
            on_progreso(indice_etapa, total_etapas, ETAPAS_PLANILLA[indice_etapa - 1])

    # Docentes del archivo de cursos que no alcanzaron el umbral de
    # coincidencia difusa contra la base de docentes. Antes quedaban con
    # Docente=None y desaparecían del cálculo sin que nadie se enterara.
    sin_coincidencia: List[str] = []
    planillas_anteriores = list(planillas_anteriores or [])
    comparacion: Optional[ResumenComparacion] = None

    def registrar_sin_coincidencia(nombres) -> None:
        for nombre in nombres:
            if nombre not in sin_coincidencia:
                sin_coincidencia.append(nombre)

    print(f"🚀 Iniciando generación de planilla para {month} {año_actual}")
    print(f"📁 Archivo coordinación: {excel_coordinacion}")

    try:
        cache_module.limpiar_cache_excel()
        cache_module.limpiar_cache_procesamiento()

        avanzar(1)
        print("📂 Cargando archivo de datos...")
        datos = cargar_archivo(data_path)
        datos = normalizar_columnas_planilla(datos)

        # Procesar CSV con nuevo formato de columnas
        if 'Detalle Curso' in datos.columns or 'Horario Completo' in datos.columns:
            datos = procesar_csv_nuevo_formato(datos)
            print("✅ Columnas del CSV procesadas correctamente")

        avanzar(2)
        # Leer Excel de docentes forzando columna de contrato como texto
        # para preservar ceros adelante (ej: 0057, 0130)
        dtype_contrato = {alias: str for alias in columnas.ALIAS_COLUMNA_CONTRATO}
        datos_docentes = pd.read_excel(excel_docentes, sheet_name="list", dtype=dtype_contrato)
        print(f"✅ Docentes cargados: {len(datos_docentes)} registros")

        if columnas.N_RUC not in datos_docentes.columns:
            for columna_ruc in columnas.ALIAS_COLUMNA_RUC[1:]:
                if columna_ruc in datos_docentes.columns:
                    datos_docentes[columnas.N_RUC] = datos_docentes[columna_ruc]
                    break

        datos_docentes[columnas.CATEGORIA_MONTO] = pd.to_numeric(datos_docentes[columnas.CATEGORIA_MONTO], errors='coerce').fillna(0)

        if 'N.° Contrato' in datos_docentes.columns:
            datos_docentes['N.° Contrato'] = datos_docentes['N.° Contrato'].apply(
                lambda valor: formatear_numero(valor, 4)
            )
        else:
            datos_docentes['N.° Contrato'] = ''

        datos = limpiar_docentes(datos, columnas.DOCENTE_ORIGINAL)
        datos = filtrar_docentes_validos(datos, columnas.DOCENTE_ORIGINAL)
        datos_docentes = limpiar_docentes(datos_docentes, columnas.DOCENTE)

        # Planilla complementaria (2ª, 3ª... del mes): dejar solo los cursos
        # nuevos o los que cambiaron de docente respecto a las anteriores.
        # Va antes de las transformaciones para que todas las hojas (y los
        # documentos de fase, que leen Planilla_Generador) vean lo mismo.
        if planillas_anteriores:
            datos, comparacion = excluir_cursos_emitidos(datos, planillas_anteriores, data_path)
            print(
                f"🔎 Comparación con {comparacion.planillas_anteriores} planilla(s) anterior(es): "
                f"{comparacion.excluidos} curso(s) ya emitido(s) excluido(s), "
                f"{comparacion.cambio_docente} con cambio de docente, {comparacion.nuevos} nuevo(s)."
            )
            if datos.empty:
                return ResultadoPlanilla(
                    ok=False,
                    mensaje=(
                        "No se generó la planilla: todos los cursos del CSV ya salieron en las "
                        "planillas anteriores (no hay cursos nuevos ni cambios de docente)."
                    ),
                    comparacion=comparacion,
                )

        datos_csv_original = datos.copy()

        datos_csv_original_procesados = aplicar_transformaciones_base(datos_csv_original)

        datos = aplicar_transformaciones_base(datos)
        datos['detalles_curso'] = datos['Curso']

        avanzar(3)
        agrupar = cache_module.agrupar_y_calcular_con_cache(
            datos, datos_docentes, 'detalles_curso', agrupar_y_calcular,
            on_sin_coincidencia=registrar_sin_coincidencia,
        )
        print("✅ Agrupación completada")

        avanzar(4)
        agrupar = agregar_examen_clasificacion(agrupar, excel_exa_clasif, TextUtils.normalizar_texto, datos_docentes)
        print("✅ Examen de clasificación agregado")

        avanzar(5)
        print("🔄 Procesando servicio de coordinación...")
        agrupar = agregar_servicio_coordinacion(agrupar, excel_coordinacion, TextUtils.normalizar_texto, datos_docentes)
        print("✅ Servicio de coordinación agregado")

        avanzar(6)
        TABLA = cache_module.construir_tabla_planilla_con_cache(agrupar, es_enero, monto_bono, table_builders.construir_tabla_planilla)

        # Usar carpeta destino seleccionada o crear carpeta por defecto
        if carpeta_destino and os.path.isdir(carpeta_destino):
            carpeta_salida = carpeta_destino
        else:
            carpeta_salida = f"{month} {año_actual}"
            os.makedirs(carpeta_salida, exist_ok=True)

        # La 2ª planilla del mes no debe sobrescribir a la 1ª.
        numero_planilla = len(planillas_anteriores) + 1
        sufijo = f" ({numero_planilla})" if numero_planilla > 1 else ""
        nombre_salida = f"Planilla {month} {año_actual}{sufijo}.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)

        avanzar(7)
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            TABLA.rename(columns={columnas.N_RUC: columnas.N_RUC_DISPLAY}).to_excel(writer, sheet_name=f"Planilla {month}", index=False)
            columnas_extra_disponibles = [c for c in columnas.COLUMNAS_EXTRA_DOCENTES if c in datos_docentes.columns]

            if os.path.exists(excel_exa_clasif):
                # Usar cache para evitar lecturas múltiples
                clasif_df = cache_module.cargar_excel_con_cache(excel_exa_clasif, sheet_name=0, header=1)
                clasif_df.to_excel(writer, sheet_name="Examen de clasificación", index=False)

            # Crear hoja de coordinación si existe el archivo
            if excel_coordinacion and os.path.exists(excel_coordinacion):
                tabla_coordinacion = table_builders.construir_tabla_coordinacion(
                    excel_coordinacion, TextUtils.normalizar_texto, datos_docentes,
                    cache_module.cargar_excel_con_cache, preparar_coordinacion_agrupada,
                )
                tabla_coordinacion.to_excel(writer, sheet_name="Servicio actualización", index=False)

            # Construir hoja Planilla_Generador usando datos ya procesados
            agrupar_gen = cache_module.agrupar_y_calcular_con_cache(
                datos_csv_original_procesados, datos_docentes, 'Curso', agrupar_y_calcular,
                on_sin_coincidencia=registrar_sin_coincidencia,
            )
            agrupar_gen = agregar_examen_clasificacion(agrupar_gen, excel_exa_clasif, TextUtils.normalizar_texto, datos_docentes)
            agrupar_gen = agregar_servicio_coordinacion(agrupar_gen, excel_coordinacion, TextUtils.normalizar_texto, datos_docentes)

            TABLA_GENERADOR = cache_module.construir_tabla_planilla_con_cache(agrupar_gen, es_enero, monto_bono, table_builders.construir_tabla_planilla)
            TABLA_GENERADOR = TABLA_GENERADOR.merge(
                datos_docentes[[columnas.DOCENTE] + columnas_extra_disponibles], on=columnas.DOCENTE, how='left'
            ).rename(columns=columnas.RENOMBRES_TABLA_GENERADOR)

            TABLA_GENERADOR_RESUMIDA = construir_tabla_planilla_generador_resumida(agrupar_gen, TABLA_GENERADOR, datos_csv_original_procesados)
            TABLA_GENERADOR_RESUMIDA.to_excel(writer, sheet_name="Planilla_Generador", index=False)

            df_carga = table_builders.construir_tabla_carga_academica(datos_csv_original_procesados, 'Planilla', traducir_dias)
            df_carga.to_excel(writer, sheet_name="Carga académica", index=False)

            if comparacion is not None:
                comparacion.detalle.to_excel(writer, sheet_name="Comparación", index=False)

        wb = load_workbook(ruta_salida)

        titulo_hojas = [
            (f"Planilla {month}",
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nPLANILLA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            ("Carga académica",
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nCARGA ACADÉMICA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            ("Examen de clasificación",
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nEXAMEN DE CLASIFICACIÓN - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL"),

            ("Servicio actualización",
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nSERVICIO DE ACTUALIZACIÓN DE MATERIALES - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL"),

            ("Comparación",
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nCOMPARACIÓN CON PLANILLAS ANTERIORES - PERIODO {month.upper()} {año_actual}\nPLANILLA N.° {numero_planilla}"),
        ]

        avanzar(8)
        # Procesar formato de todas las hojas de manera optimizada
        procesar_formato_multiple_hojas(wb, titulo_hojas)

        avanzar(9)
        hojas_ordenadas = [
            f"Planilla {month}",
            "Carga académica",
            "Examen de clasificación",
            "Servicio actualización",
            "Planilla_Generador",
            "Comparación",
        ]
        wb = ordenar_hojas_excel(wb, hojas_ordenadas)
        wb.save(ruta_salida)

        if sin_coincidencia:
            print(
                f"⚠️  {len(sin_coincidencia)} docente(s) del archivo de cursos no se pudieron emparejar "
                f"con la base de docentes y quedaron fuera del cálculo: {', '.join(sin_coincidencia)}"
            )

        return ResultadoPlanilla(
            ok=True,
            mensaje=f"{nombre_salida} generado correctamente.",
            ruta_archivo=ruta_salida,
            sin_coincidencia=sin_coincidencia,
            comparacion=comparacion,
        )

    except Exception as e:
        logger.exception("Error generando la planilla de %s %s", month, año_actual)
        return ResultadoPlanilla(
            ok=False,
            mensaje=f"No se pudo generar la planilla: {e}",
            error=str(e),
            sin_coincidencia=sin_coincidencia,
            comparacion=comparacion,
        )
    finally:
        cache_module.limpiar_cache_planilla()
        cache_module.limpiar_cache_procesamiento()
