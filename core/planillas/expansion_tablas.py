"""Expansión de la tabla agrupada: una fila por servicio, y la tabla resumida por docente."""

import pandas as pd

from core.fases.utils import TextUtils
from core.planillas import columnas as col
from core.planillas.docente_matching import normalizar_columna_ruc
from core.planillas.formato_datos import formatear_numero
from core.planillas.transformations import generar_siguiente_curso_intensivo_desde_fila
from core.reglas_negocio import regla_para_curso


# ================================= EXPANSIÓN DE FILAS POR CURSO CON MODALIDAD ===================================================

def expandir_filas_por_curso(agrupar_df, datos_csv_procesados):
    filas_expandidas = []

    for _, row_docente in agrupar_df.iterrows():
        # Usar el nombre ORIGINAL del CSV (columna 'docente' en minúscula)
        # en lugar del nombre oficial (columna 'Docente' con mayúscula)
        docente_nombre_csv = row_docente.get(col.DOCENTE_ORIGINAL, row_docente.get(col.DOCENTE, None))

        # Validar que el nombre del docente no sea None o vacío
        if pd.isna(docente_nombre_csv) or not str(docente_nombre_csv).strip():
            print(f"⚠️ Saltando fila con nombre de docente vacío o None")
            continue

        # Convertir a string y limpiar
        docente_nombre_limpio = str(docente_nombre_csv).strip().upper()

        # Obtener todos los cursos de este docente del CSV original
        cursos_docente = datos_csv_procesados[
            datos_csv_procesados[col.DOCENTE_ORIGINAL].fillna('').astype(str).str.strip().str.upper() == docente_nombre_limpio
        ]

        # Debug: mostrar si no se encuentran cursos
        if cursos_docente.empty:
            print(f"⚠️ No se encontraron cursos para: {docente_nombre_csv} (oficial: {row_docente.get(col.DOCENTE, 'N/A')})")

        # Para cada curso académico del docente
        for _, curso_row in cursos_docente.iterrows():
            fila_curso = row_docente.copy()
            fila_curso['Curso_Individual'] = curso_row['Curso']
            fila_curso['Modalidad_Curso'] = curso_row['modalidad']
            fila_curso['Tipo_Servicio'] = col.TIPO_SERVICIO_CURSO_DICTADO
            regla = regla_para_curso(curso_row['Curso'])
            fila_curso['Horas_Servicio'] = regla.horas_dictado
            # Validar categoria_monto antes de multiplicar
            categoria_monto = row_docente[col.CATEGORIA_MONTO] if not pd.isna(row_docente[col.CATEGORIA_MONTO]) else 0
            fila_curso['Monto_Individual'] = (regla.tarifa_hora_fija or categoria_monto) * regla.horas_dictado
            filas_expandidas.append(fila_curso)

        # Agregar fila para examen de clasificación si aplica
        examen_clasif = row_docente[col.EXAMEN_CLASIFICACION] if not pd.isna(row_docente[col.EXAMEN_CLASIFICACION]) else 0
        if examen_clasif > 0:
            fila_examen = row_docente.copy()
            fila_examen['Curso_Individual'] = 'Examen de clasificación'
            fila_examen['Modalidad_Curso'] = 'VIRTUAL'
            fila_examen['Tipo_Servicio'] = col.TIPO_SERVICIO_EXAMEN_CLASIF
            categoria_monto = row_docente[col.CATEGORIA_MONTO] if not pd.isna(row_docente[col.CATEGORIA_MONTO]) else 1
            horas_examen = int(examen_clasif / categoria_monto) if categoria_monto > 0 else 0
            fila_examen['Horas_Servicio'] = horas_examen
            fila_examen['Monto_Individual'] = examen_clasif
            filas_expandidas.append(fila_examen)

        # Agregar fila para diseño de exámenes si aplica
        diseno_examenes = row_docente[col.DISENO_EXAMENES] if not pd.isna(row_docente[col.DISENO_EXAMENES]) else 0
        if diseno_examenes > 0:
            fila_diseno = row_docente.copy()
            fila_diseno['Curso_Individual'] = 'Diseño de exámenes'
            fila_diseno['Modalidad_Curso'] = 'N/A'
            fila_diseno['Tipo_Servicio'] = col.TIPO_SERVICIO_DISENO_EXAMENES
            categoria_monto = row_docente[col.CATEGORIA_MONTO] if not pd.isna(row_docente[col.CATEGORIA_MONTO]) else 1
            horas_diseno = int(diseno_examenes / categoria_monto) if categoria_monto > 0 else 0
            fila_diseno['Horas_Servicio'] = horas_diseno
            fila_diseno['Monto_Individual'] = diseno_examenes
            filas_expandidas.append(fila_diseno)

        # Agregar fila para servicio de actualización si aplica
        servicio_act = row_docente[col.SERVICIO_ACTUALIZACION] if not pd.isna(row_docente[col.SERVICIO_ACTUALIZACION]) else 0
        if servicio_act > 0:
            fila_servicio = row_docente.copy()
            fila_servicio['Curso_Individual'] = 'Servicio de actualización de materiales de enseñanza'
            fila_servicio['Modalidad_Curso'] = 'VIRTUAL'
            fila_servicio['Tipo_Servicio'] = col.TIPO_SERVICIO_ACTUALIZACION
            # Usar las horas reales del archivo de coordinación, NO calcularlas
            horas_total = row_docente.get(col.RESUMIDA_HORAS_TOTAL, 0)
            try:
                horas_servicio = int(horas_total) if not pd.isna(horas_total) and horas_total != '' else 0
            except (ValueError, TypeError):
                horas_servicio = 0
            fila_servicio['Horas_Servicio'] = horas_servicio
            fila_servicio['Monto_Individual'] = servicio_act
            filas_expandidas.append(fila_servicio)

    # Convertir lista de filas en DataFrame
    if not filas_expandidas:
        print("⚠️ No se generaron filas expandidas")
        # Retornar DataFrame vacío con las columnas esperadas
        columnas_nuevas = ['Curso_Individual', 'Modalidad_Curso', 'Tipo_Servicio', 'Horas_Servicio', 'Monto_Individual']
        columnas_resto = [c for c in agrupar_df.columns]
        return pd.DataFrame(columns=columnas_nuevas + columnas_resto)

    df_expandido = pd.DataFrame(filas_expandidas)

    # ============= MEJORAS DE ORGANIZACIÓN =============

    # 1. Definir orden lógico para tipos de servicio
    df_expandido['_orden_servicio'] = df_expandido['Tipo_Servicio'].map(col.ORDEN_TIPO_SERVICIO).fillna(99)

    # Mapeo a nombres descriptivos en español
    df_expandido['Tipo_Servicio_Desc'] = df_expandido['Tipo_Servicio'].map(col.NOMBRES_TIPO_SERVICIO)

    # 2. Ordenar por: Docente (alfabético) -> Tipo de Servicio (orden lógico) -> Curso Individual
    df_expandido = df_expandido.sort_values(
        by=[col.DOCENTE, '_orden_servicio', 'Curso_Individual'],
        ascending=[True, True, True]
    ).reset_index(drop=True)

    # 3. Agregar numeración global y por docente
    df_expandido.insert(0, 'N°', range(1, len(df_expandido) + 1))

    # Agregar contador de servicio por docente (1, 2, 3, etc. para cada docente)
    df_expandido['Servicio_Nro'] = df_expandido.groupby(col.DOCENTE).cumcount() + 1

    # 4. Reorganizar columnas de forma más lógica
    # Primero: Identificación y servicio, luego detalles, al final datos administrativos repetidos
    columnas_ordenadas = col.COLUMNAS_EXPANDIDA_PRINCIPALES + col.COLUMNAS_EXPANDIDA_RESUMEN + col.COLUMNAS_EXPANDIDA_ADMINISTRATIVAS

    # Agregar cualquier columna que falte al final
    columnas_restantes = [c for c in df_expandido.columns
                            if c not in columnas_ordenadas and c not in ['_orden_servicio', col.DOCENTE_ORIGINAL, 'curso', 'Tipo_Servicio']]

    columnas_finales = [c for c in columnas_ordenadas if c in df_expandido.columns] + columnas_restantes

    # Eliminar columnas temporales y redundantes
    df_expandido = df_expandido.drop(columns=['_orden_servicio', 'Tipo_Servicio'], errors='ignore')

    # Aplicar orden final
    df_expandido = df_expandido[columnas_finales]

    return df_expandido


def construir_tabla_planilla_generador_resumida(agrupar_df, tabla_generador, datos_csv_procesados):
    tabla_base = normalizar_columna_ruc(tabla_generador)

    if col.CANTIDAD_CURSOS not in tabla_base.columns and 'Cantidad Cursos' in tabla_base.columns:
        tabla_base[col.CANTIDAD_CURSOS] = tabla_base['Cantidad Cursos']

    if tabla_base.empty:
        return pd.DataFrame(columns=col.COLUMNAS_PLANILLA_GENERADOR_RESUMIDA)

    for columna in col.COLUMNAS_RESUMIDA_TEXTO:
        if columna not in tabla_base.columns:
            tabla_base[columna] = ''

    for columna in col.COLUMNAS_RESUMIDA_NUMERICAS:
        if columna not in tabla_base.columns:
            tabla_base[columna] = 0
        tabla_base[columna] = pd.to_numeric(tabla_base[columna], errors='coerce').fillna(0)

    # Evitar que filas sin nombre de docente se agrupen en una sola fila residual.
    tabla_base = tabla_base[tabla_base[col.DOCENTE].fillna('').astype(str).str.strip() != ''].copy()

    # Agrupar por Docente: una fila por docente con montos sumados
    _primer_valor = lambda x: next((v for v in x if pd.notna(v) and str(v).strip() and str(v).strip().lower() != 'nan'), '')
    _unir_unicos = lambda x: ' / '.join(dict.fromkeys(v for v in x if pd.notna(v) and str(v).strip() and str(v).strip().lower() != 'nan'))

    agrupado_resumen = tabla_base.groupby(col.DOCENTE, as_index=False).agg({
        col.N_RUC: _primer_valor,
        col.RESUMIDA_CATEGORIA_LETRA: _primer_valor,
        col.RESUMIDA_CATEGORIA_MONTO: 'max',
        col.SEDE: _unir_unicos,
        col.CANTIDAD_CURSOS: 'sum',
        col.CURSO_DICTADO: 'sum',
        col.RESUMIDA_DISENO_EXAMENES: 'sum',
        col.RESUMIDA_EXAMEN_CLASIF: 'sum',
        col.RESUMIDA_HORAS_TOTAL: 'sum',
        col.RESUMIDA_SERVICIO_ACTUALIZACION: 'sum',
        col.RESUMIDA_TOTAL_PAGO: 'sum',
        col.RESUMIDA_ESTADO_DOCENTE: _primer_valor,
        col.RESUMIDA_DOCENTE_IDIOMA: _unir_unicos,
        col.RESUMIDA_TIPO_DOCUMENTO: _primer_valor,
        col.RESUMIDA_NUMERO_DNI: _primer_valor,
        col.RESUMIDA_NUMERO_CELULAR: _primer_valor,
        col.RESUMIDA_DOMICILIO_DOCENTE: _primer_valor,
        col.RESUMIDA_CORREO_PERSONAL: _primer_valor,
        col.RESUMIDA_NRO_CONTRATO: _primer_valor,
    })

    mapeo_docentes = (
        agrupar_df[[col.DOCENTE_ORIGINAL, col.DOCENTE]]
        .dropna(subset=[col.DOCENTE])
        .assign(docente_key=lambda df: df[col.DOCENTE_ORIGINAL].apply(TextUtils.normalizar_texto))
        .drop_duplicates(subset=['docente_key'], keep='last')
    )

    cursos_df = datos_csv_procesados.copy()
    cursos_df['docente_key'] = cursos_df[col.DOCENTE_ORIGINAL].apply(TextUtils.normalizar_texto)
    cursos_df = cursos_df.merge(
        mapeo_docentes[['docente_key', col.DOCENTE]],
        on='docente_key',
        how='left'
    )

    columnas_docente = [c for c in ['Docente_y', col.DOCENTE, 'Docente_x'] if c in cursos_df.columns]
    if columnas_docente:
        cursos_df[col.DOCENTE] = (
            cursos_df[columnas_docente]
            .replace(r'^\s*$', pd.NA, regex=True)
            .bfill(axis=1)
            .iloc[:, 0]
        )
    elif col.DOCENTE_ORIGINAL in cursos_df.columns:
        cursos_df[col.DOCENTE] = cursos_df[col.DOCENTE_ORIGINAL]
    else:
        cursos_df[col.DOCENTE] = pd.NA

    if 'Curso' not in cursos_df.columns:
        cursos_df['Curso'] = (
            cursos_df['idioma'].astype(str) + ' ' +
            cursos_df['nivel'].astype(str) + ' ' +
            cursos_df['ciclo'].astype(str)
        )

    cursos_df['modalidad'] = cursos_df.get('modalidad', '').fillna('').astype(str)
    cursos_df = cursos_df.dropna(subset=[col.DOCENTE])
    cursos_df = cursos_df[cursos_df[col.DOCENTE].fillna('').astype(str).str.strip() != ''].copy()

    # Para cursos intensivos, agregar también el siguiente curso (ej: Básico 1 -> Básico 2).
    cursos_df['Curso_Siguiente_Intensivo'] = cursos_df.apply(generar_siguiente_curso_intensivo_desde_fila, axis=1)
    cursos_intensivos_adicionales = cursos_df[cursos_df['Curso_Siguiente_Intensivo'] != ''].copy()
    if not cursos_intensivos_adicionales.empty:
        cursos_intensivos_adicionales['Curso'] = cursos_intensivos_adicionales['Curso_Siguiente_Intensivo']
        cursos_df = pd.concat([cursos_df, cursos_intensivos_adicionales], ignore_index=True)

    es_virtual = cursos_df['modalidad'].str.upper().str.contains('VIRTUAL', na=False)
    cursos_df['Curso_Virtual_tmp'] = cursos_df['Curso'].where(es_virtual, '')
    cursos_df['Curso_Presencial_tmp'] = cursos_df['Curso'].where(~es_virtual, '')

    _unir_todos = lambda x: ' / '.join(v for v in x if pd.notna(v) and str(v).strip() and str(v).strip().lower() != 'nan')

    cursos_por_docente = cursos_df.groupby(col.DOCENTE, as_index=False).agg({
        'Curso_Virtual_tmp': _unir_todos,
        'Curso_Presencial_tmp': _unir_todos
    }).rename(columns={
        'Curso_Virtual_tmp': 'Curso_Virtual',
        'Curso_Presencial_tmp': 'Curso_Presencial'
    })

    agrupado_resumen = agrupado_resumen.merge(cursos_por_docente, on=col.DOCENTE, how='left')
    agrupado_resumen['Curso_Virtual'] = agrupado_resumen['Curso_Virtual'].fillna('')
    agrupado_resumen['Curso_Presencial'] = agrupado_resumen['Curso_Presencial'].fillna('')
    agrupado_resumen[col.RESUMIDA_NUMERO_DNI] = agrupado_resumen[col.RESUMIDA_NUMERO_DNI].apply(lambda valor: formatear_numero(valor, 8))
    agrupado_resumen[col.RESUMIDA_NRO_CONTRATO] = agrupado_resumen[col.RESUMIDA_NRO_CONTRATO].apply(lambda valor: formatear_numero(valor, 4))

    agrupado_resumen = agrupado_resumen.sort_values(col.DOCENTE).reset_index(drop=True)
    agrupado_resumen = agrupado_resumen.drop(columns=['N°'], errors='ignore')
    agrupado_resumen.insert(0, 'N°', range(1, len(agrupado_resumen) + 1))

    for columna in col.COLUMNAS_PLANILLA_GENERADOR_RESUMIDA:
        if columna not in agrupado_resumen.columns:
            agrupado_resumen[columna] = '' if columna not in col.COLUMNAS_RESUMIDA_NUMERICAS else 0

    return agrupado_resumen[col.COLUMNAS_PLANILLA_GENERADOR_RESUMIDA]
