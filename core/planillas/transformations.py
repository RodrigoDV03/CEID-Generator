import re
import unicodedata

import pandas as pd


_DIAS_INTENSIVOS = [
    {'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY'},
    {'SATURDAY', 'SUNDAY'},
]

_MAPA_DIAS = {
    'MON': 'MONDAY',
    'MONDAY': 'MONDAY',
    'LUN': 'MONDAY',
    'LUNES': 'MONDAY',
    'TUE': 'TUESDAY',
    'TUESDAY': 'TUESDAY',
    'MAR': 'TUESDAY',
    'MARTES': 'TUESDAY',
    'WED': 'WEDNESDAY',
    'WEDNESDAY': 'WEDNESDAY',
    'MIE': 'WEDNESDAY',
    'MIERCOLES': 'WEDNESDAY',
    'THU': 'THURSDAY',
    'THURSDAY': 'THURSDAY',
    'JUE': 'THURSDAY',
    'JUEVES': 'THURSDAY',
    'FRI': 'FRIDAY',
    'FRIDAY': 'FRIDAY',
    'VIE': 'FRIDAY',
    'VIERNES': 'FRIDAY',
    'SAT': 'SATURDAY',
    'SATURDAY': 'SATURDAY',
    'SAB': 'SATURDAY',
    'SABADO': 'SATURDAY',
    'SUN': 'SUNDAY',
    'SUNDAY': 'SUNDAY',
    'DOM': 'SUNDAY',
    'DOMINGO': 'SUNDAY',
}


def _normalizar_texto_simple(texto):
    texto_limpio = unicodedata.normalize('NFKD', str(texto).strip().upper())
    return ''.join(ch for ch in texto_limpio if not unicodedata.combining(ch))


def _normalizar_dias(dias):
    if pd.isna(dias):
        return []

    texto = str(dias).strip().strip('{}')
    if not texto:
        return []

    partes = [parte.strip() for parte in re.split(r'[,/]+', texto) if parte.strip()]
    dias_normalizados = []
    for parte in partes:
        token = _normalizar_texto_simple(parte).replace('.', '')
        dias_normalizados.append(_MAPA_DIAS.get(token, token))
    return dias_normalizados


def _asegurar_prefijo_intensivo(nivel):
    nivel_limpio = str(nivel).strip()
    if not nivel_limpio:
        return nivel_limpio
    if nivel_limpio.startswith('Intensivo'):
        return nivel_limpio
    return f'Intensivo {nivel_limpio}'


def _construir_curso(idioma, nivel, ciclo):
    return f"{str(idioma).strip()} {str(nivel).strip()} {str(ciclo).strip()}"


def ajustar_nivel(row):
    nivel = row['nivel']
    idioma = row['idioma']
    ciclo = row['ciclo']
    try:
        ciclo_num = int(str(ciclo).strip())
    except Exception:
        ciclo_num = None

    if nivel == 'General' and idioma in {'Inglés', 'Portugués'} and ciclo_num is not None:
        if 1 <= ciclo_num <= 8:
            return 'Posgrado Básico'
        elif 9 <= ciclo_num <= 18:
            return 'Posgrado Intermedio'
    return nivel


def ajustar_modalidad(row):
    dias = row['dias']
    hora_fin = row['horafin']
    modalidad = row['modalidad']

    dias_normalizados = set(_normalizar_dias(dias))
    hora_fin_normalizada = _normalizar_texto_simple(hora_fin)
    if re.fullmatch(r'\d{1,2}:\d{2}$', hora_fin_normalizada):
        hora_fin_normalizada = f"{hora_fin_normalizada}:00"

    if dias_normalizados in _DIAS_INTENSIVOS and hora_fin_normalizada in {'22:30:00', '15:00:00'}:
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


def _partir_cursos_intensivos(cursos_texto):
    if pd.isna(cursos_texto):
        return []

    return [p.strip() for p in str(cursos_texto).split('/') if p and str(p).strip()]


def expandir_texto_cursos_intensivos(cursos_texto):
    partes = _partir_cursos_intensivos(cursos_texto)
    if not partes:
        return ''

    resultado = []
    cursos_ya_presentes = set()

    for curso in dict.fromkeys(partes):
        resultado.append(curso)
        cursos_ya_presentes.add(curso)
        siguiente = generar_siguiente_curso_intensivo(curso)
        if siguiente and siguiente not in cursos_ya_presentes:
            resultado.append(siguiente)
            cursos_ya_presentes.add(siguiente)

    return ' / '.join(resultado)


def _calcular_ciclo_siguiente_portugues(nivel_sin_intensivo, ciclo_actual):
    if 'Básico' in nivel_sin_intensivo or 'Basico' in nivel_sin_intensivo:
        if ciclo_actual == 5:
            return 'Intensivo Intermedio', '1'
        else:
            return None, str(ciclo_actual + 1)
    elif 'Intermedio' in nivel_sin_intensivo:
        if ciclo_actual == 4:
            return 'Intensivo Avanzado', '1'
        else:
            return None, str(ciclo_actual + 1)
    else:  # Avanzado u otro
        return None, str(ciclo_actual + 1)


def expandir_filas_carga_intensivos(datos, modalidad_filter='INTENSIVO', aplicar_logica_portugues=False):
    datos_base = datos.copy()
    if datos_base.empty:
        return datos_base

    # Si se aplica lógica de Portugués, especificar modalidad exacta
    if aplicar_logica_portugues:
        modalidad_filter = 'INTENSIVO VIRTUAL'

    filas_adicionales = []
    mask_intensivo = datos_base['modalidad'].astype(str).str.upper().str.contains(modalidad_filter, na=False)

    for _, row in datos_base[mask_intensivo].iterrows():
        nivel_original = str(row.get('nivel', '')).strip()
        if aplicar_logica_portugues:
            nivel_original = _asegurar_prefijo_intensivo(nivel_original)
            if isinstance(row.name, int):
                datos_base.at[row.name, 'nivel'] = nivel_original

        try:
            ciclo_actual = int(str(row.get('ciclo', '')).strip())
        except (TypeError, ValueError):
            continue

        idioma = str(row.get('idioma', '')).strip()
        if not idioma or not nivel_original:
            continue

        fila_nueva = row.copy()
        idioma_nueva = str(row.get('idioma', '')).strip()

        # Calcular ciclo siguiente (con lógica especial para Portugués si aplica)
        if aplicar_logica_portugues and idioma_nueva == 'Portugués':
            nivel_sin_intensivo = nivel_original.replace('Intensivo ', '').strip()
            nivel_siguiente, ciclo_siguiente = _calcular_ciclo_siguiente_portugues(nivel_sin_intensivo, ciclo_actual)
            if nivel_siguiente:
                fila_nueva['nivel'] = nivel_siguiente
            else:
                fila_nueva['nivel'] = nivel_original
            fila_nueva['ciclo'] = ciclo_siguiente
        else:
            ciclo_siguiente = ciclo_actual + 1
            fila_nueva['ciclo'] = str(ciclo_siguiente)

        fila_nueva['Curso'] = _construir_curso(idioma_nueva, fila_nueva['nivel'], fila_nueva['ciclo'])
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

    return _construir_curso(idioma, nivel, ciclo_actual + 1)


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
        datos_transformados.loc[idx, 'nivel'] = _asegurar_prefijo_intensivo(datos_transformados.loc[idx, 'nivel'])

    datos_transformados['Curso'] = datos_transformados.apply(
        lambda row: _construir_curso(row['idioma'], row['nivel'], row['ciclo']),
        axis=1,
    )
    return datos_transformados
