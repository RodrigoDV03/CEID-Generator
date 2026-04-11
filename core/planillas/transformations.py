import re

import pandas as pd


def ajustar_nivel(row):
    nivel = row['nivel']
    idioma = row['idioma']
    ciclo = row['ciclo']
    if nivel == 'General':
        if idioma == 'Inglés':
            try:
                ciclo_num = int(str(ciclo).strip())
            except Exception:
                ciclo_num = None
            if ciclo_num is not None:
                if 1 <= ciclo_num <= 8:
                    return 'Posgrado Básico'
                elif 9 <= ciclo_num <= 18:
                    return 'Posgrado Intermedio'
        elif idioma == 'Portugués':
            try:
                ciclo_num = int(str(ciclo).strip())
            except Exception:
                ciclo_num = None
            if ciclo_num is not None:
                if 1 <= ciclo_num <= 4:
                    return 'Posgrado Básico'
                elif 5 <= ciclo_num <= 8:
                    return 'Posgrado Intermedio'
    return nivel


def ajustar_modalidad(row):
    dias = row['dias']
    hora_fin = row['horafin']
    modalidad = row['modalidad']

    if (hora_fin == '22:30:00' or hora_fin == '15:00:00') and (dias == '{MONDAY,TUESDAY,WEDNESDAY,THURSDAY}' or dias == '{SATURDAY,SUNDAY}'):
        modalidad = 'INTENSIVO VIRTUAL'

    return modalidad


def generar_siguiente_curso_intensivo(curso):
    if pd.isna(curso):
        return ''

    curso_str = str(curso).strip()
    if not curso_str or 'intensivo' not in curso_str.lower():
        return ''

    match = re.search(r'^(.*?)(\d+)\s*$', curso_str)
    if not match:
        return ''

    prefijo = match.group(1).rstrip()
    ciclo_actual = int(match.group(2))
    return f"{prefijo} {ciclo_actual + 1}"


def expandir_texto_cursos_intensivos(cursos_texto):
    if pd.isna(cursos_texto):
        return ''

    partes = [p.strip() for p in str(cursos_texto).split('/') if p and str(p).strip()]
    if not partes:
        return ''

    resultado = list(partes)
    cursos_ya_presentes = set(partes)

    for curso in set(partes):
        siguiente = generar_siguiente_curso_intensivo(curso)
        if siguiente and siguiente not in cursos_ya_presentes:
            resultado.append(siguiente)
            cursos_ya_presentes.add(siguiente)

    return ' / '.join(resultado)


def expandir_filas_carga_intensivos(datos):
    datos_base = datos.copy()
    if datos_base.empty:
        return datos_base

    filas_adicionales = []
    mask_intensivo = datos_base['modalidad'].astype(str).str.upper().str.contains('INTENSIVO', na=False)

    for _, row in datos_base[mask_intensivo].iterrows():
        try:
            ciclo_actual = int(str(row.get('ciclo', '')).strip())
        except (TypeError, ValueError):
            continue

        idioma = str(row.get('idioma', '')).strip()
        nivel = str(row.get('nivel', '')).strip()
        if not idioma or not nivel:
            continue

        fila_nueva = row.copy()
        ciclo_siguiente = ciclo_actual + 1
        fila_nueva['ciclo'] = str(ciclo_siguiente)
        fila_nueva['Curso'] = f"{idioma} {nivel} {ciclo_siguiente}"
        filas_adicionales.append(fila_nueva)

    if filas_adicionales:
        datos_extra = pd.DataFrame(filas_adicionales)
        datos_base = pd.concat([datos_base, datos_extra], ignore_index=True)

        columnas_dedupe = [c for c in ['docente', 'Curso', 'modalidad', 'dias', 'horainicio', 'horafin'] if c in datos_base.columns]
        if columnas_dedupe:
            datos_base = datos_base.drop_duplicates(subset=columnas_dedupe, keep='first')

    return datos_base


def generar_siguiente_curso_intensivo_desde_fila(row):
    modalidad = str(row.get('modalidad', '')).strip().upper()
    if 'INTENSIVO' not in modalidad:
        return ''

    try:
        ciclo_actual = int(str(row.get('ciclo', '')).strip())
    except (TypeError, ValueError):
        return generar_siguiente_curso_intensivo(row.get('Curso', ''))

    idioma = str(row.get('idioma', '')).strip()
    nivel = str(row.get('nivel', '')).strip()
    if not idioma or not nivel:
        return generar_siguiente_curso_intensivo(row.get('Curso', ''))

    return f"{idioma} {nivel} {ciclo_actual + 1}"


def procesar_ediciones_idioma(datos):
    datos_procesados = datos.copy()

    mask_edicion = datos_procesados['idioma'].astype(str).str.startswith('Edición', na=False)

    for idx in datos_procesados[mask_edicion].index:
        idioma_actual = str(datos_procesados.loc[idx, 'idioma']).strip()
        if not idioma_actual.endswith('Inglés'):
            datos_procesados.loc[idx, 'idioma'] = idioma_actual + ' Inglés'

    return datos_procesados


def aplicar_transformaciones_base(datos):
    datos_transformados = datos.copy()

    datos_transformados = procesar_ediciones_idioma(datos_transformados)

    datos_transformados['nivel'] = datos_transformados.apply(ajustar_nivel, axis=1)
    datos_transformados['modalidad'] = datos_transformados.apply(ajustar_modalidad, axis=1)

    mask_intensivo = datos_transformados['modalidad'] == 'INTENSIVO VIRTUAL'
    for idx in datos_transformados[mask_intensivo].index:
        nivel_original = str(datos_transformados.loc[idx, 'nivel']).strip()
        if not nivel_original.startswith('Intensivo'):
            datos_transformados.loc[idx, 'nivel'] = f'Intensivo {nivel_original}'

    datos_transformados['Curso'] = (
        datos_transformados['idioma'].astype(str) + ' ' +
        datos_transformados['nivel'].astype(str) + ' ' +
        datos_transformados['ciclo'].astype(str)
    )
    return datos_transformados


def procesar_cursos_intensivos(datos):
    datos_procesados = datos.copy()

    mask_intensivo = datos_procesados['modalidad'] == 'INTENSIVO VIRTUAL'

    if not mask_intensivo.any():
        return datos_procesados

    filas_intensivas = datos_procesados[mask_intensivo].copy()
    filas_adicionales = []

    for idx, row in filas_intensivas.iterrows():
        nivel_original = str(row['nivel']).strip()
        if not nivel_original.startswith('Intensivo'):
            datos_procesados.loc[idx, 'nivel'] = f'Intensivo {nivel_original}'

        fila_adicional = row.copy()
        idioma = str(row['idioma']).strip()

        try:
            ciclo_actual = int(str(row['ciclo']).strip())

            if idioma == 'Portugués':
                nivel_sin_intensivo = nivel_original.replace('Intensivo ', '').strip()

                if 'Básico' in nivel_sin_intensivo or 'Basico' in nivel_sin_intensivo:
                    if ciclo_actual == 5:
                        fila_adicional['nivel'] = 'Intensivo Intermedio'
                        fila_adicional['ciclo'] = '1'
                    else:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                elif 'Intermedio' in nivel_sin_intensivo:
                    if ciclo_actual == 4:
                        fila_adicional['nivel'] = 'Intensivo Avanzado'
                        fila_adicional['ciclo'] = '1'
                    else:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                elif 'Avanzado' in nivel_sin_intensivo:
                    if ciclo_actual < 3:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                    else:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                else:
                    fila_adicional['ciclo'] = str(ciclo_actual + 1)
            else:
                fila_adicional['ciclo'] = str(ciclo_actual + 1)

        except (ValueError, TypeError):
            fila_adicional['ciclo'] = str(row['ciclo']) + '+1'

        if not str(fila_adicional['nivel']).startswith('Intensivo'):
            fila_adicional['nivel'] = f"Intensivo {fila_adicional['nivel']}"

        filas_adicionales.append(fila_adicional)

    if filas_adicionales:
        filas_adicionales_df = pd.DataFrame(filas_adicionales)
        datos_procesados = pd.concat([datos_procesados, filas_adicionales_df], ignore_index=True)

    mask_todas_intensivas = datos_procesados['modalidad'] == 'INTENSIVO VIRTUAL'
    datos_procesados.loc[mask_todas_intensivas, 'Curso'] = (
        datos_procesados.loc[mask_todas_intensivas, 'idioma'].astype(str) + ' ' +
        datos_procesados.loc[mask_todas_intensivas, 'nivel'].astype(str) + ' ' +
        datos_procesados.loc[mask_todas_intensivas, 'ciclo'].astype(str)
    )

    return datos_procesados
