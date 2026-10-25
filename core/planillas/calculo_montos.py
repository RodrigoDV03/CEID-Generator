"""Cálculo de montos a pagar: agrupación por docente, examen de clasificación y servicio de coordinación."""

import os
import pandas as pd

from core.planillas import cache as _cache
from core.planillas import columnas as col
from core.planillas.docente_matching import (
    normalizar_columna_ruc,
    crear_mapeo_fuzzy,
    agregar_docentes_faltantes_al_merge,
)
from core.planillas.formato_datos import extraer_numero_horas
from core.reglas_negocio import (
    TARIFA_COORDINACION_SIN_CATEGORIA,
    UMBRAL_COINCIDENCIA_DOCENTE_PLANILLA,
    regla_para_curso,
)


def preparar_coordinacion_agrupada(coordinacion_df, normalizar_texto, mostrar_columna_docente=False, mostrar_columnas_disponibles=False):
    # Detectar automáticamente las columnas importantes.
    columna_docente = None
    columna_horas = None

    for c in coordinacion_df.columns:
        col_lower = str(c).lower().strip()
        if any(patron in col_lower for patron in ['docente', 'profesor', 'teacher', 'instructor']):
            columna_docente = c
            if mostrar_columna_docente:
                print(f"✅ Columna de docente encontrada: '{c}'")
            break

    if columna_docente is None:
        print("⚠️ No se encontró columna de docente en archivo de coordinación")
        return None

    prioridad_horas = [
        'horas totales', 'horas_totales', 'total', 'totales',
        'hora total', 'total horas', 'total_horas',
        'horas', 'hora', 'hour', 'tiempo', 'time',
        'horas semanales', 'horas_semanales', 'semanal'
    ]

    for patron in prioridad_horas:
        for c in coordinacion_df.columns:
            col_lower = str(c).lower().strip()
            if patron in col_lower:
                columna_horas = c
                break
        if columna_horas:
            break

    if columna_horas is None:
        print("⚠️ No se encontró columna de horas en archivo de coordinación")
        if mostrar_columnas_disponibles:
            print(f"Columnas disponibles: {list(coordinacion_df.columns)}")
        return None

    # Limpiar y procesar datos.
    coordinacion_df = coordinacion_df.dropna(subset=[columna_docente, columna_horas])
    coordinacion_df = coordinacion_df[coordinacion_df[columna_docente].astype(str).str.strip() != '']
    coordinacion_df = coordinacion_df[coordinacion_df[columna_docente].astype(str).str.lower() != 'nan']
    coordinacion_df[columna_docente] = coordinacion_df[columna_docente].astype(str).str.strip()
    coordinacion_df[columna_horas] = coordinacion_df[columna_horas].apply(extraer_numero_horas)

    # Agrupar por docente (sumar horas si aparece múltiples veces).
    coordinacion_agrupado = coordinacion_df[[columna_docente, columna_horas]].groupby(columna_docente, as_index=False)[columna_horas].sum()
    coordinacion_agrupado = coordinacion_agrupado.rename(columns={columna_docente: 'Docente_Original', columna_horas: col.RESUMIDA_HORAS_TOTAL})
    coordinacion_agrupado['docente_norm'] = coordinacion_agrupado['Docente_Original'].apply(normalizar_texto)

    return coordinacion_agrupado


# ------------- CÁLCULO DE MONTOS -------------

def agrupar_y_calcular(df, datos_docentes, col_curso, on_sin_coincidencia=None):
    """
    Agrupa las filas de cursos por docente y calcula sus montos.

    `on_sin_coincidencia`, si se pasa, se llama con la lista de nombres de
    docente (tal como aparecen en el archivo de cursos) que NO alcanzaron el
    umbral de coincidencia difusa contra la base de docentes. Antes esos
    nombres simplemente quedaban con `Docente = None` y desaparecían del
    cálculo sin que nadie se enterara.
    """
    df_trabajo = df.copy()

    # Cada curso superintensivo genera un curso adicional (siguiente ciclo) que
    # también se paga ("superintensivo" contiene "intensivo").
    df_trabajo['_intensivo_adicional'] = (
        df_trabajo[col_curso]
        .fillna('')
        .astype(str)
        .str.contains('intensivo', case=False, na=False)
        .astype(int)
    )

    # Horas y montos por curso según su programa (regular, Junior, Suficiencia).
    # Se acumulan por separado las horas que se pagan con la categoría del
    # docente y los montos a tarifa fija, porque la categoría recién se
    # conoce después de emparejar con la base de docentes.
    texto_programa = df_trabajo['programa_educativo'].fillna('').astype(str) if 'programa_educativo' in df_trabajo.columns else ''
    reglas = (texto_programa + ' ' + df_trabajo[col_curso].fillna('').astype(str)).apply(regla_para_curso)
    cursos_por_fila = 1 + df_trabajo['_intensivo_adicional']
    df_trabajo['_horas_dictado_categoria'] = reglas.apply(lambda r: 0 if r.tarifa_hora_fija else r.horas_dictado) * cursos_por_fila
    df_trabajo['_monto_dictado_fijo'] = reglas.apply(lambda r: r.horas_dictado * (r.tarifa_hora_fija or 0)) * cursos_por_fila
    df_trabajo['_horas_diseno'] = reglas.apply(lambda r: r.horas_diseno_examenes) * cursos_por_fila

    # Agrupar por docente: una fila por docente con todos sus cursos concatenados
    agrupado = (df_trabajo.groupby(col.DOCENTE_ORIGINAL).agg(
        curso=(col_curso, lambda x: ' / '.join(str(v).strip() for v in x if str(v).strip() and str(v).strip().lower() != 'nan')),
        cantidad_cursos_base=(col_curso, 'count'),
        cursos_intensivos_adicionales=('_intensivo_adicional', 'sum'),
        horas_dictado_categoria=('_horas_dictado_categoria', 'sum'),
        monto_dictado_fijo=('_monto_dictado_fijo', 'sum'),
        horas_diseno=('_horas_diseno', 'sum'),
    ).reset_index())

    agrupado[col.CANTIDAD_CURSOS] = agrupado['cantidad_cursos_base'] + agrupado['cursos_intensivos_adicionales']
    agrupado.drop(columns=['cantidad_cursos_base', 'cursos_intensivos_adicionales'], inplace=True)

    nombres_base = datos_docentes[col.DOCENTE].tolist()

    # Crear mapeo fuzzy una sola vez (solo sobre docentes únicos)
    mapeo_docentes = crear_mapeo_fuzzy(
        agrupado[col.DOCENTE_ORIGINAL].tolist(), nombres_base, umbral=UMBRAL_COINCIDENCIA_DOCENTE_PLANILLA
    )

    if on_sin_coincidencia:
        sin_coincidencia = sorted({nombre for nombre, match in mapeo_docentes.items() if match is None})
        if sin_coincidencia:
            on_sin_coincidencia(sin_coincidencia)

    # Aplicar mapeo usando el diccionario
    agrupado[col.DOCENTE] = agrupado[col.DOCENTE_ORIGINAL].map(mapeo_docentes)

    datos_docentes = normalizar_columna_ruc(datos_docentes)
    agrupado = agrupado.merge(
        datos_docentes[[col.DOCENTE, col.SEDE, col.CATEGORIA_LETRA, col.CATEGORIA_MONTO, col.N_RUC, col.ESTADO]],
        on=col.DOCENTE, how='left',
    )
    agrupado[col.CURSO_DICTADO] = (
        agrupado[col.CATEGORIA_MONTO] * agrupado['horas_dictado_categoria'] + agrupado['monto_dictado_fijo']
    )
    agrupado[col.DISENO_EXAMENES] = agrupado[col.CATEGORIA_MONTO] * agrupado['horas_diseno']
    agrupado.drop(columns=['horas_dictado_categoria', 'monto_dictado_fijo', 'horas_diseno'], inplace=True)
    return agrupado


def _procesar_archivo_docentes_con_monto(
    df,
    ruta_archivo,
    normalizar_texto_fn,
    datos_docentes,
    patrones_columna_monto,
    nombre_columna_resultado,
    procesar_resultado_fn=None,
):

    if not os.path.exists(ruta_archivo):
        df[nombre_columna_resultado] = 0
        return df

    try:
        # Detectar la fila real de encabezados para tolerar archivos con un título arriba.
        datos_preview = _cache.cargar_excel_con_cache(ruta_archivo, sheet_name=0, header=None)
        header_idx = 0

        max_scan = min(len(datos_preview), 10)
        for i in range(max_scan):
            fila_normalizada = {
                normalizar_texto_fn(valor)
                for valor in datos_preview.iloc[i].tolist()
            }
            if 'DOCENTE' in fila_normalizada and any(
                normalizar_texto_fn(patron) in fila_normalizada
                for patron in patrones_columna_monto
            ):
                header_idx = i
                break

        # Leer el archivo Excel usando la fila de encabezados detectada
        datos_df = _cache.cargar_excel_con_cache(ruta_archivo, sheet_name=0, header=header_idx)

        # Verificar que existe la columna Docente (con fallback a búsqueda de similares)
        if col.DOCENTE not in datos_df.columns:
            posibles_docentes = [c for c in datos_df.columns if 'docente' in str(c).lower()]
            if posibles_docentes:
                datos_df = datos_df.rename(columns={posibles_docentes[0]: col.DOCENTE})
            else:
                df[nombre_columna_resultado] = 0
                return df

        # Limpiar valores NaN/VACÍOS en la columna Docente
        datos_df = datos_df.dropna(subset=[col.DOCENTE])
        datos_df = datos_df[datos_df[col.DOCENTE].astype(str).str.strip() != '']
        datos_df = datos_df[datos_df[col.DOCENTE].astype(str).str.lower() != 'nan']

        # Buscar la columna de Monto (con patrones prioritarios)
        columna_monto = None
        for patron in patrones_columna_monto:
            if patron in datos_df.columns:
                columna_monto = patron
                break
            # Búsqueda case-insensitive si es string
            if isinstance(patron, str):
                for c in datos_df.columns:
                    if str(c).lower().strip() == patron.lower().strip():
                        columna_monto = c
                        break
                if columna_monto:
                    break

        if columna_monto is None:
            print(f"⚠️ No se encontró columna de monto en {ruta_archivo}")
            df[nombre_columna_resultado] = 0
            return df

        # Procesar y normalizar datos
        datos_df[col.DOCENTE] = datos_df[col.DOCENTE].astype(str).str.strip()
        datos_df['docente_norm'] = datos_df[col.DOCENTE].apply(normalizar_texto_fn)

        # Asegurar que df tiene docente_norm
        if 'docente_norm' not in df.columns:
            df['docente_norm'] = df[col.DOCENTE].apply(normalizar_texto_fn)

        # Hacer el merge (left para mantener todos los de df)
        merge_result = df.merge(datos_df[['docente_norm', columna_monto]], on='docente_norm', how='left')

        # Agregar docentes que no están en df pero sí en datos_df
        docentes_no_en_df = datos_df[~datos_df['docente_norm'].isin(df['docente_norm'])]

        if not docentes_no_en_df.empty and datos_docentes is not None:
            print(f"📋 Agregando {len(docentes_no_en_df)} docentes del archivo {ruta_archivo}...")

            merge_result = agregar_docentes_faltantes_al_merge(
                merge_result,
                docentes_no_en_df,
                datos_docentes,
                normalizar_texto_fn,
                campo_nombre_validacion=col.DOCENTE,
                extra_builder=lambda row_faltante: {
                    columna_monto: row_faltante[columna_monto],
                },
            )

        # Llenar la columna de resultado
        df = merge_result
        if procesar_resultado_fn:
            # Si se proporciona función custom de post-procesamiento
            df[nombre_columna_resultado] = procesar_resultado_fn(df, columna_monto)
        else:
            # Default: llenar directamente con el monto
            df[nombre_columna_resultado] = df[columna_monto].fillna(0)

        # Limpiar columna temporal
        if columna_monto in df.columns:
            df.drop(columns=[columna_monto], inplace=True)
        if 'docente_norm' in df.columns:
            df.drop(columns=['docente_norm'], inplace=True)

    except Exception as e:
        print(f"⚠️ Error al procesar {ruta_archivo}: {e}")
        df[nombre_columna_resultado] = 0

    return df


def agregar_servicio_coordinacion(df, ruta_coordinacion, normalizar_texto, datos_docentes=None):
    if not os.path.exists(ruta_coordinacion):
        df[col.SERVICIO_ACTUALIZACION] = 0
        df[col.RESUMIDA_HORAS_TOTAL] = 0
        return df

    try:
        # Leer el archivo Excel
        coordinacion_df = _cache.cargar_excel_con_cache(ruta_coordinacion, sheet_name=0, header=0)

        coordinacion_agrupado = preparar_coordinacion_agrupada(
            coordinacion_df,
            normalizar_texto,
            mostrar_columna_docente=True,
            mostrar_columnas_disponibles=True
        )
        if coordinacion_agrupado is None:
            df[col.SERVICIO_ACTUALIZACION] = 0
            df[col.RESUMIDA_HORAS_TOTAL] = 0
            return df

        # Si datos_docentes está disponible, calcular montos por categoría
        if datos_docentes is not None:
            # Normalizar nombres en datos_docentes también
            datos_docentes_temp = datos_docentes.copy()
            datos_docentes_temp['docente_norm'] = datos_docentes_temp[col.DOCENTE].apply(normalizar_texto)

            # Hacer merge con datos_docentes para obtener categoría
            coordinacion_con_categoria = coordinacion_agrupado.merge(
                datos_docentes_temp[['docente_norm', col.DOCENTE, col.CATEGORIA_MONTO]], on='docente_norm', how='left'
            )

            # Calcular monto: Horas * Categoria (Monto) (tarifa según categoría del docente)
            coordinacion_con_categoria['Monto_Coordinacion'] = (
                coordinacion_con_categoria[col.RESUMIDA_HORAS_TOTAL] * coordinacion_con_categoria[col.CATEGORIA_MONTO].fillna(0)
            )

            # Usar el nombre corregido del merge
            coordinacion_con_categoria['Docente_Correcto'] = coordinacion_con_categoria[col.DOCENTE].fillna(coordinacion_con_categoria['Docente_Original'])

        else:
            # Si no hay datos_docentes, usar la tarifa fija de excepción (sin categoría asignada)
            coordinacion_con_categoria = coordinacion_agrupado.copy()
            coordinacion_con_categoria['Monto_Coordinacion'] = coordinacion_agrupado[col.RESUMIDA_HORAS_TOTAL] * TARIFA_COORDINACION_SIN_CATEGORIA
            coordinacion_con_categoria['Docente_Correcto'] = coordinacion_agrupado['Docente_Original']

        # Preparar datos para merge con df principal
        coordinacion_final = coordinacion_con_categoria[['docente_norm', 'Monto_Coordinacion', col.RESUMIDA_HORAS_TOTAL, 'Docente_Correcto']].copy()

        # Si df no tiene docente_norm, crearlo
        if 'docente_norm' not in df.columns:
            df['docente_norm'] = df[col.DOCENTE].apply(normalizar_texto)

        # Hacer merge con df principal
        merge_result = df.merge(coordinacion_final, on='docente_norm', how='left')

        # AGREGAR DOCENTES DE COORDINACIÓN QUE NO ESTÁN EN DF (similar a examen clasificación)
        docentes_coord_no_en_df = coordinacion_final[~coordinacion_final['docente_norm'].isin(df['docente_norm'])]

        merge_result = agregar_docentes_faltantes_al_merge(
            merge_result,
            docentes_coord_no_en_df,
            datos_docentes,
            normalizar_texto,
            campo_nombre_validacion='Docente_Correcto',
            extra_builder=lambda row_coord: {
                'Monto_Coordinacion': row_coord['Monto_Coordinacion'],
                col.RESUMIDA_HORAS_TOTAL: row_coord[col.RESUMIDA_HORAS_TOTAL],
                'Docente_Correcto': row_coord['Docente_Correcto'],
            },
        )

        # Finalizar el procesamiento
        df = merge_result
        df[col.SERVICIO_ACTUALIZACION] = df['Monto_Coordinacion'].fillna(0)
        # Preservar Horas_Total con valor por defecto 0
        if col.RESUMIDA_HORAS_TOTAL in df.columns:
            df[col.RESUMIDA_HORAS_TOTAL] = df[col.RESUMIDA_HORAS_TOTAL].fillna(0)
        else:
            df[col.RESUMIDA_HORAS_TOTAL] = 0

        # Limpiar columnas temporales (MANTENER Horas_Total para la hoja expandida)
        columnas_a_eliminar = ['docente_norm', 'Monto_Coordinacion', 'Docente_Correcto']
        for c in columnas_a_eliminar:
            if c in df.columns:
                df.drop(columns=[c], inplace=True)

        print(f"✅ Servicio de coordinación procesado correctamente")

    except Exception as e:
        print(f"⚠️ Error al procesar archivo de coordinación: {e}")
        df[col.SERVICIO_ACTUALIZACION] = 0
        df[col.RESUMIDA_HORAS_TOTAL] = 0

    return df


def agregar_examen_clasificacion(df, ruta_clasificacion, normalizar_texto, datos_docentes=None):
    return _procesar_archivo_docentes_con_monto(
        df,
        ruta_clasificacion,
        normalizar_texto,
        datos_docentes,
        patrones_columna_monto=['Monto Total', 'monto'],
        nombre_columna_resultado=col.EXAMEN_CLASIFICACION,
        procesar_resultado_fn=None  # Default: llenar directo con monto
    )
