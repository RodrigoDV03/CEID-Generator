import pandas as pd


_excel_cache = {}
_cache_planilla_anterior = {}
_cache_agrupacion = {}
_cache_tablas_construidas = {}


def generar_key_datos(df, col_curso):
    try:
        datos_relevantes = df[['docente', col_curso]].copy()
        datos_str = datos_relevantes.to_string()
        key = f"{col_curso}_{hash(datos_str)}"
        return key
    except Exception:
        return f"{col_curso}_{df.shape[0]}_{df.shape[1]}"


def limpiar_cache_excel():
    _excel_cache.clear()


def cargar_excel_con_cache(ruta, sheet_name=0, header: int | str = 'infer'):
    cache_key = f"{ruta}_{sheet_name}_{header}"

    if cache_key not in _excel_cache:
        if header == 'infer':
            _excel_cache[cache_key] = pd.read_excel(ruta, sheet_name=sheet_name)
        else:
            _excel_cache[cache_key] = pd.read_excel(ruta, sheet_name=sheet_name, header=header)

    return _excel_cache[cache_key].copy()


def agrupar_y_calcular_con_cache(df, datos_docentes, col_curso, agrupar_fn):
    key_datos = generar_key_datos(df, col_curso)

    if key_datos in _cache_agrupacion:
        return _cache_agrupacion[key_datos].copy()

    resultado = agrupar_fn(df, datos_docentes, col_curso)
    _cache_agrupacion[key_datos] = resultado.copy()

    return resultado


def leer_planilla_anterior_con_cache(ruta_planilla, header_fn):
    if ruta_planilla in _cache_planilla_anterior:
        return _cache_planilla_anterior[ruta_planilla]

    try:
        fila_header = header_fn(ruta_planilla)
        if fila_header is None:
            fila_header = 6

        planilla_anterior = pd.read_excel(ruta_planilla, sheet_name="Primera carga académica", header=fila_header)
        _cache_planilla_anterior[ruta_planilla] = planilla_anterior

        return planilla_anterior
    except Exception as e:
        print(f"Error al leer planilla anterior: {e}")
        return pd.DataFrame()


def limpiar_cache_planilla():
    _cache_planilla_anterior.clear()


def limpiar_cache_procesamiento():
    _cache_agrupacion.clear()
    _cache_tablas_construidas.clear()


def obtener_header_planilla_con_cache(ruta_planilla):
    try:
        df_raw = pd.read_excel(ruta_planilla, sheet_name="Primera carga académica", header=None, nrows=10)

        for idx, fila in df_raw.iterrows():
            if "Docente" in fila.values and "Curso" in fila.values:
                return idx

        return None
    except Exception as e:
        print(f"Error al obtener header: {e}")
        return None


def construir_tabla_planilla_con_cache(df, es_enero, monto_bono, construir_fn):
    try:
        columnas_relevantes = ['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N_Ruc', 'curso', 'cantidad_cursos', 'Curso Dictado', 'Diseño de Examenes', 'Examen Clasif.', 'Estado']

        df_relevante = df[columnas_relevantes]
        key_tabla = hash(df_relevante.to_string() + str(es_enero) + str(monto_bono))

        if key_tabla in _cache_tablas_construidas:
            return _cache_tablas_construidas[key_tabla].copy()

        tabla = construir_fn(df, es_enero, monto_bono)
        _cache_tablas_construidas[key_tabla] = tabla.copy()

        return tabla

    except Exception:
        return construir_fn(df, es_enero, monto_bono)
