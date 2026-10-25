"""Emparejamiento de nombres de docente entre el archivo de cursos y la base de docentes."""

import pandas as pd
from fuzzywuzzy import process, fuzz

from core.planillas import columnas as col
from core.reglas_negocio import UMBRAL_COINCIDENCIA_DOCENTE_PLANILLA


def normalizar_columna_ruc(df):
    """Asegura que `df` tenga la columna `N_Ruc`, tomándola de alguno de sus alias si hace falta."""
    df_normalizado = df.copy()

    if col.N_RUC in df_normalizado.columns:
        return df_normalizado

    for columna in col.ALIAS_COLUMNA_RUC[1:]:
        if columna in df_normalizado.columns:
            df_normalizado[col.N_RUC] = df_normalizado[columna]
            break

    return df_normalizado


def crear_mapeo_fuzzy(nombres_input, nombres_base, umbral=UMBRAL_COINCIDENCIA_DOCENTE_PLANILLA):
    """
    Empareja cada nombre de `nombres_input` contra el más parecido en
    `nombres_base`. Si el puntaje no alcanza `umbral`, el nombre queda
    mapeado a `None` (sin coincidencia) en vez de a un match forzado.
    """
    mapeo = {}
    for nombre in nombres_input:
        if nombre not in mapeo:  # Evitar recálculos
            resultado = process.extractOne(nombre, nombres_base, scorer=fuzz.token_sort_ratio)
            mapeo[nombre] = resultado[0] if resultado[1] >= umbral else None
    return mapeo


def agregar_docentes_faltantes_al_merge(
    merge_result,
    docentes_faltantes_df,
    datos_docentes,
    normalizar_texto,
    campo_nombre_validacion,
    extra_builder,
):
    """
    Agrega al resultado de un merge las filas de docentes que aparecen en
    `docentes_faltantes_df` (p. ej. en el archivo de coordinación o de
    examen de clasificación) pero no estaban en la tabla principal.
    """
    if docentes_faltantes_df.empty or datos_docentes is None:
        return merge_result

    datos_docentes_temp = datos_docentes.copy()
    datos_docentes_temp['docente_norm'] = datos_docentes_temp[col.DOCENTE].apply(normalizar_texto)

    filas_nuevas = []
    for _, row_faltante in docentes_faltantes_df.iterrows():
        nombre_validacion = row_faltante.get(campo_nombre_validacion)
        if pd.isna(nombre_validacion) or str(nombre_validacion).strip() == '' or str(nombre_validacion).lower() == 'nan':
            continue

        docente_info = datos_docentes_temp[datos_docentes_temp['docente_norm'] == row_faltante['docente_norm']]
        if docente_info.empty:
            continue

        docente_info = docente_info.iloc[0]
        nueva_fila = {
            col.DOCENTE: docente_info[col.DOCENTE],
            col.SEDE: docente_info[col.SEDE],
            col.CATEGORIA_LETRA: docente_info[col.CATEGORIA_LETRA],
            col.CATEGORIA_MONTO: docente_info[col.CATEGORIA_MONTO],
            col.N_RUC: docente_info[col.N_RUC],
            col.ESTADO: docente_info[col.ESTADO],
            'curso': '',
            col.CANTIDAD_CURSOS: 0,
            col.CURSO_DICTADO: 0,
            col.DISENO_EXAMENES: 0,
            col.EXAMEN_CLASIFICACION: 0,
            'docente_norm': row_faltante['docente_norm'],
        }
        nueva_fila.update(extra_builder(row_faltante))
        filas_nuevas.append(nueva_fila)

    if not filas_nuevas:
        return merge_result

    filas_nuevas_df = pd.DataFrame(filas_nuevas)
    return pd.concat([merge_result, filas_nuevas_df], ignore_index=True)
