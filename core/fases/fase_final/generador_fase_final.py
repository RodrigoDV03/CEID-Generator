import datetime
import pandas as pd
from typing import Callable, Optional, Dict
from core.fases.models import (
    DocenteData,
    PaymentData,
    DocumentConfig,
    CursoDetalle,
    ResultadoDocente,
    ResultadoProcesoFase,
)
from core.fases.services import ExcelReaderService
from core.fases.base_generator import BaseFaseGenerator
from .builders import ConformidadBuilder, ControlAvanceBuilder


class FaseFinalGenerator(BaseFaseGenerator):

    def __init__(self, config: DocumentConfig):
        super().__init__(config)

        # Builders
        self.conformidad_builder = ConformidadBuilder(config)
        self.control_builder = ControlAvanceBuilder(config)

    def generar_conformidad(
        self,
        docente: DocenteData,
        payment: PaymentData,
        planilla_path: str
    ) -> str:
        # Crear carpeta del docente
        carpeta_docente = self.docente_service.crear_carpeta_docente(
            self.config.carpeta_destino,
            self.config,
            docente
        )

        # Generar descripción completa
        descripcion_completa = self.generar_descripcion_completa(docente, payment, planilla_path)

        # Generar documento
        ruta = self.conformidad_builder.generar(
            docente,
            payment,
            descripcion_completa,
            carpeta_docente
        )

        return ruta

    def generar_control_avance(
        self,
        docente: DocenteData,
        payment: PaymentData
    ) -> Optional[str]:
        # Solo para docentes con contrato
        if not docente.es_contrato:
            return None

        # Solo si se especificó número de armada
        if self.config.numero_armada == 'sin armada':
            return None

        # Crear carpeta del docente
        carpeta_docente = self.docente_service.crear_carpeta_docente(
            self.config.carpeta_destino,
            self.config,
            docente
        )

        # Generar documento
        ruta = self.control_builder.generar(
            docente,
            payment,
            carpeta_docente
        )

        return ruta

    def procesar_docente(
        self,
        docente: DocenteData,
        payment: PaymentData,
        planilla_path: str
    ) -> ResultadoDocente:
        # Validar docente
        es_valido, mensaje = self.docente_service.validar_docente(docente)
        if not es_valido:
            print(f"❌ Error - {mensaje}")
            return ResultadoDocente(nombre=docente.nombre or "N/A", ok=False, error=mensaje)

        try:
            # Generar conformidad (para todos)
            self.generar_conformidad(docente, payment, planilla_path)
            print(f"✅ {docente.nombre} - Documento de conformidad generado correctamente.")
            return ResultadoDocente(nombre=docente.nombre, ok=True)

        except Exception as e:
            print(f"❌ Error generando conformidad para {docente.nombre}: {e}")
            return ResultadoDocente(nombre=docente.nombre, ok=False, error=str(e))

    def procesar_control_pagos(
        self,
        df_control: 'pd.DataFrame',
        resultado_proceso: ResultadoProcesoFase,
        idiomas_por_docente: Optional[Dict[str, str]] = None,
        contratos_por_docente: Optional[Dict[str, str]] = None,
        on_progreso: Optional[Callable[[int, int, str], None]] = None,
    ) -> None:
        idiomas_por_docente = idiomas_por_docente or {}
        contratos_por_docente = contratos_por_docente or {}

        total_filas = len(df_control)
        for indice, (_, fila) in enumerate(df_control.iterrows(), start=1):
            nombre = "N/A"
            try:
                # Extraer nombre del docente
                nombre = self.excel_service.extraer_docente_nombre_control(fila)
                if nombre == "N/A":
                    continue

                nombre_key = nombre.strip().upper()
                idioma_docente = idiomas_por_docente.get(nombre_key, "")
                numero_contrato = contratos_por_docente.get(nombre_key, "")

                # Crear objeto DocenteData básico
                docente = DocenteData(
                    nombre=nombre,
                    dni="",
                    ruc="",
                    idioma=idioma_docente,
                    especialidad=str(getattr(fila, "Especialidad", "")),
                    numero_contrato=numero_contrato,
                    estado_docente="CONTRATO"
                )

                # Extraer datos de pago del control
                payment = self.excel_service.extraer_payment_data_control(fila)

                # Generar control de avance
                ruta = self.generar_control_avance(docente, payment)
                if ruta:
                    print(f"✅ {nombre} - Control de pagos generado correctamente.")
                    resultado_proceso.registrar(ResultadoDocente(nombre=nombre, ok=True))

            except Exception as e:
                print(f"❌ Error procesando control para {nombre}: {e}")
                resultado_proceso.registrar(ResultadoDocente(nombre=nombre, ok=False, error=str(e)))
                continue
            finally:
                if on_progreso:
                    on_progreso(indice, total_filas, "controles de avance")

    def procesar_control_pagos_administrativos(
        self,
        df_planilla: 'pd.DataFrame',
        resultado_proceso: ResultadoProcesoFase,
        on_progreso: Optional[Callable[[int, int, str], None]] = None,
    ) -> None:
        total_filas = len(df_planilla)
        for indice, (_, fila) in enumerate(df_planilla.iterrows(), start=1):
            nombre = "N/A"
            try:
                nombre = str(getattr(fila, "Docente", "N/A")).strip()
                if not nombre or nombre == "N/A":
                    continue

                estado = str(getattr(fila, "Estado_docente", "")).strip().upper()
                if estado != "CONTRATO":
                    continue

                numero_contrato = ExcelReaderService._extraer_numero_contrato(
                    getattr(fila, "Nro_Contrato", getattr(fila, "Nro_contrato", ""))
                )

                docente = DocenteData(
                    nombre=nombre,
                    dni=str(getattr(fila, "Numero_dni", "")).strip(),
                    ruc=str(getattr(fila, "N_Ruc", "")).strip(),
                    curso=str(getattr(fila, "Curso", "")).strip(),
                    especialidad=str(getattr(fila, "Actividades_admin", "")).strip(),
                    numero_contrato=numero_contrato,
                    estado_docente="CONTRATO",
                    idioma=str(getattr(fila, "Curso", "")).strip()
                )

                payment = self.excel_service.extraer_payment_data_control(fila)
                ruta = self.generar_control_avance(docente, payment)
                if ruta:
                    print(f"✅ {nombre} - Control de pagos administrativo generado correctamente.")
                    resultado_proceso.registrar(ResultadoDocente(nombre=nombre, ok=True))

            except Exception as e:
                print(f"❌ Error procesando control administrativo para {nombre}: {e}")
                resultado_proceso.registrar(ResultadoDocente(nombre=nombre, ok=False, error=str(e)))
                continue
            finally:
                if on_progreso:
                    on_progreso(indice, total_filas, "controles de avance")


def procesar_planilla_fase_final(
    planilla_path: str,
    excel_control_pagos: Optional[str],
    hoja: str,
    carpeta_salida: str,
    mes: str,
    numero_armada: str,
    tipo_fase_final: str,
    on_progreso: Optional[Callable[[int, int, str], None]] = None,
) -> ResultadoProcesoFase:
    # Crear configuración
    config = DocumentConfig(
        mes=mes,
        anio=datetime.datetime.now().year,
        numero_armada=numero_armada,
        tipo_fase="final",
        tipo_docente=tipo_fase_final,
        carpeta_destino=carpeta_salida
    )

    # Crear generador
    generador = FaseFinalGenerator(config)

    # Leer planilla principal
    excel_service = ExcelReaderService()
    df = excel_service.leer_planilla(planilla_path, hoja)

    # Mapa de idioma por docente desde la planilla principal
    idiomas_por_docente: Dict[str, str] = {}
    contratos_por_docente: Dict[str, str] = {}

    resultado_proceso = ResultadoProcesoFase()
    total_filas = len(df)

    # Procesar cada docente (conformidad)
    for indice, (_, fila) in enumerate(df.iterrows(), start=1):
        try:
            # Extraer datos
            docente = excel_service.extraer_docente_data(fila)
            payment = excel_service.extraer_payment_data(fila)

            idiomas_por_docente[docente.nombre.strip().upper()] = docente.idioma.strip()
            contratos_por_docente[docente.nombre.strip().upper()] = docente.numero_contrato.strip()

            # Procesar docente (pasar planilla_path para leer cursos detallados)
            resultado_docente = generador.procesar_docente(docente, payment, planilla_path)
            resultado_proceso.registrar(resultado_docente)

        except ValueError as e:
            print(f"Error: {e}")
            resultado_proceso.registrar(ResultadoDocente(nombre=f"Fila {indice}", ok=False, error=str(e)))
            continue
        except Exception as e:
            print(f"Error inesperado: {e}")
            resultado_proceso.registrar(ResultadoDocente(nombre=f"Fila {indice}", ok=False, error=str(e)))
            continue
        finally:
            if on_progreso:
                on_progreso(indice, total_filas, "conformidades")

    # Procesar control de pagos si aplica
    if tipo_fase_final == "planilla docente (con contrato)" and excel_control_pagos:
        try:
            df_control = excel_service.leer_control_pagos(excel_control_pagos)
            generador.procesar_control_pagos(
                df_control, resultado_proceso, idiomas_por_docente, contratos_por_docente, on_progreso=on_progreso
            )
        except Exception as e:
            print(f"Error procesando control de pagos: {e}")
    elif tipo_fase_final == "administrativo":
        try:
            generador.procesar_control_pagos_administrativos(df, resultado_proceso, on_progreso=on_progreso)
        except Exception as e:
            print(f"Error procesando control de pagos administrativos: {e}")

    return resultado_proceso
