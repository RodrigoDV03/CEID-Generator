import pandas as pd
import os
from fuzzywuzzy import process, fuzz
import unicodedata

# Cache simple para archivos Excel
_excel_cache = {}

def limpiar_cache_excel():
    global _excel_cache
    _excel_cache.clear()

def cargar_excel_con_cache(ruta, sheet_name=0, header='infer'):
    cache_key = f"{ruta}_{sheet_name}_{header}"
    
    if cache_key not in _excel_cache:
        if header == 'infer':
            _excel_cache[cache_key] = pd.read_excel(ruta, sheet_name=sheet_name)
        else:
            _excel_cache[cache_key] = pd.read_excel(ruta, sheet_name=sheet_name, header=header)
    
    return _excel_cache[cache_key].copy()  # Retorna una copia para evitar modificaciones accidentales

def cargar_archivo(ruta):
    extension = os.path.splitext(ruta)[-1].lower()
    if extension == ".csv":
        try:
            return pd.read_csv(ruta, sep=',')
        except Exception:
            return pd.read_csv(ruta)
    elif extension in [".xls", ".xlsx"]:
        return pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {extension}")
    
def limpiar_docentes(df, col):
    df[col] = df[col].astype(str).str.strip()
    return df

def aplicar_transformaciones_base(datos):
    datos_transformados = datos.copy()

    mask_general = datos_transformados['nivel'] == 'General'
    mask_ingles = datos_transformados['idioma'] == 'Inglés'
    mask_portugues = datos_transformados['idioma'] == 'Portugués'
    
    # Aplicar transformaciones vectorizadas donde sea posible
    datos_transformados['nivel_optimized'] = datos_transformados['nivel']
    
    # Caso específico para Inglés General
    datos_transformados.loc[mask_general & mask_ingles, 'nivel_optimized'] = 'Posgrado Intermedio'
    
    # Caso específico para Portugués General
    mask_portugues_general = mask_general & mask_portugues
    if mask_portugues_general.any():
        # Para Portugués, necesitamos evaluar los ciclos
        for idx in datos_transformados[mask_portugues_general].index:
            try:
                ciclo_num = int(str(datos_transformados.loc[idx, 'ciclo']).strip())
                if 1 <= ciclo_num <= 4:
                    datos_transformados.loc[idx, 'nivel_optimized'] = 'Posgrado Básico'
                elif 5 <= ciclo_num <= 8:
                    datos_transformados.loc[idx, 'nivel_optimized'] = 'Posgrado Intermedio'
            except:
                pass  # Mantener el valor original si hay error
    
    # Para casos restantes, aplicar la función original
    mask_restantes = ~(mask_general & (mask_ingles | mask_portugues))
    if mask_restantes.any():
        datos_transformados.loc[mask_restantes, 'nivel_optimized'] = datos_transformados.loc[mask_restantes].apply(ajustar_nivel, axis=1)
    
    datos_transformados['nivel'] = datos_transformados['nivel_optimized']
    datos_transformados.drop('nivel_optimized', axis=1, inplace=True)
    
    # Aplicar modalidad (mantener apply por ahora ya que la lógica es compleja)
    datos_transformados['modalidad'] = datos_transformados.apply(ajustar_modalidad, axis=1)
    
    # Construir curso de forma vectorizada
    datos_transformados['Curso'] = (
        datos_transformados['idioma'].astype(str) + ' ' + 
        datos_transformados['nivel'].astype(str) + ' ' + 
        datos_transformados['ciclo'].astype(str)
    )
    
    return datos_transformados

def crear_mapeo_fuzzy(nombres_input, nombres_base, umbral=85):
    mapeo = {}
    for nombre in nombres_input:
        if nombre not in mapeo:  # Evitar recálculos
            resultado = process.extractOne(nombre, nombres_base, scorer=fuzz.token_sort_ratio)
            mapeo[nombre] = resultado[0] if resultado[1] >= umbral else None
    return mapeo

def agrupar_y_calcular(df, datos_docentes, col_curso):
    agrupado = (df.groupby('docente').agg(curso=(col_curso, lambda x: ' / '.join(x)), cantidad_cursos=(col_curso, 'count')).reset_index())
    nombres_base = datos_docentes['Docente'].tolist()
    
    # Crear mapeo fuzzy una sola vez
    mapeo_docentes = crear_mapeo_fuzzy(agrupado['docente'].tolist(), nombres_base, umbral=85)
    
    # Aplicar mapeo usando el diccionario
    agrupado['Docente'] = agrupado['docente'].map(mapeo_docentes)
    
    agrupado = agrupado.merge(datos_docentes[['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N°. Ruc', 'Estado']], on='Docente', how='left')
    agrupado['Curso Dictado'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 28
    agrupado['Diseño de Examenes'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 4
    return agrupado

def agrupar_y_calcular_con_cache(df, datos_docentes, col_curso):
    # Generar key único basado en los datos de entrada
    key_datos = generar_key_datos(df, col_curso)
    
    # Verificar si ya está en cache
    if key_datos in _cache_agrupacion:
        return _cache_agrupacion[key_datos].copy()
    
    # Si no está en cache, calcular y guardar
    resultado = agrupar_y_calcular(df, datos_docentes, col_curso)
    _cache_agrupacion[key_datos] = resultado.copy()
    
    return resultado


def agregar_clasificacion(df, ruta_clasificacion, normalizar_texto):
    if os.path.exists(ruta_clasificacion):
        try:
            # Usar cache para evitar lecturas múltiples
            clasif_df = cargar_excel_con_cache(ruta_clasificacion, sheet_name=0, header=1)
            clasif_df['Docente'] = clasif_df['Docente'].astype(str).str.strip()
            clasif_df['docente_norm'] = clasif_df['Docente'].apply(normalizar_texto)
            df['docente_norm'] = df['Docente'].apply(normalizar_texto)
            df = df.merge(clasif_df[['docente_norm', 'Monto']], on='docente_norm', how='left')
            df['Examen Clasif.'] = df['Monto'].fillna(0)
            df.drop(columns=['docente_norm', 'Monto'], inplace=True)
        except Exception as e:
            print(f"⚠️ Error al leer archivo de clasificación: {e}")
            df['Examen Clasif.'] = 0
    else:
        df['Examen Clasif.'] = 0
    return df

def construir_tabla(df):
    tabla = pd.DataFrame({
        'N°': range(1, len(df) + 1),
        'Docente': df['Docente'],
        'Sede': df['Sede'],
        'Categoria (Letra)': df['Categoria (Letra)'],
        'Categoria (Monto)': df['Categoria (Monto)'],
        'N°. Ruc': df['N°. Ruc'],
        'Curso': df['curso'],
        'Curso Dictado': df['Curso Dictado'],
        'Extra Curso': 0,
        'Cantidad Cursos': df['cantidad_cursos'],
        'Diseño de Examenes': df['Diseño de Examenes'],
        'Examen Clasif.': df['Examen Clasif.'],
        'Total Pago S/.': 0,
        'Estado': df['Estado']
    })
    tabla['Total Pago S/.'] = (tabla['Curso Dictado'] + tabla['Extra Curso'] + tabla['Diseño de Examenes'] + tabla['Examen Clasif.'])
    return tabla

def construir_tabla_con_cache(df):
    try:
        # Crear hash basado en las columnas relevantes para la tabla
        columnas_relevantes = ['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N°. Ruc', 'curso', 'cantidad_cursos', 'Curso Dictado', 'Diseño de Examenes', 'Examen Clasif.', 'Estado']
        
        df_relevante = df[columnas_relevantes]
        key_tabla = hash(df_relevante.to_string())
        
        # Verificar cache
        if key_tabla in _cache_tablas_construidas:
            return _cache_tablas_construidas[key_tabla].copy()
        
        # Si no está en cache, construir y guardar
        tabla = construir_tabla(df)
        _cache_tablas_construidas[key_tabla] = tabla.copy()
        
        return tabla
        
    except Exception as e:
        # Fallback: usar función original si hay error en cache
        return construir_tabla(df)

def ajustar_nivel(row):
        nivel = row['nivel']
        idioma = row['idioma']
        ciclo = row['ciclo']
        if nivel == 'General':
            if idioma == 'Inglés':
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

    if (hora_fin == '22:30:00'or hora_fin == '15:00:00') and (dias == '{MONDAY,TUESDAY,WEDNESDAY,THURSDAY}' or dias == '{SATURDAY,SUNDAY}'):
        modalidad = 'INTENSIVO VIRTUAL'

    return modalidad

def crear_df_carga(datos, estado_planilla):
    datos = datos.copy()

    if 'nivel' not in datos.columns:
        datos['nivel'] = datos.apply(ajustar_nivel, axis=1)
    if 'Curso' not in datos.columns:
        datos['Curso'] = datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

    df = pd.DataFrame({
        'Dias': datos['dias'].apply(traducir_dias),
        'H. Inicio': datos['horainicio'].astype(str).str[:5],
        'H. Fin': datos['horafin'].astype(str).str[:5],
        'Idioma': datos['idioma'],
        'Nivel': datos['nivel'].str.replace('Ã¡', 'á', regex=False),
        'Ciclo': datos['ciclo'],
        'Curso': datos['Curso'],
        'Sede': datos['sede'],
        'Sec.': '',
        'Matr.': datos['matriculados'],
        'Docente': datos['docente'],
        'Modalidad': datos['modalidad'],
        'Estado Planilla': estado_planilla
    })
    df = df.sort_values(by='Docente').reset_index(drop=True)
    df.insert(0, 'N°', range(1, len(df) + 1))
    return df

def ordenar_hojas_excel(wb, hojas_ordenadas):
    hojas_existentes = wb.sheetnames
    nuevas_hojas = [hoja for hoja in hojas_ordenadas if hoja in hojas_existentes]
    for idx, hoja in enumerate(nuevas_hojas):
        wb._sheets.insert(idx, wb[hoja])
    wb._sheets = wb._sheets[:len(nuevas_hojas)] + [s for s in wb._sheets if s.title not in nuevas_hojas]
    return wb

def traducir_dias(dias_raw: str) -> str:
    dias_dict = {
        'MONDAY': 'Lun', 'TUESDAY': 'Mar', 'WEDNESDAY': 'Mié',
        'THURSDAY': 'Jue', 'FRIDAY': 'Vie', 'SATURDAY': 'Sáb', 'SUNDAY': 'Dom'
    }
    dias = dias_raw.strip('{}').split(',') if isinstance(dias_raw, str) else []
    return ', '.join(dias_dict.get(d.strip().upper(), d.strip()) for d in dias)

def normalizar_texto(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize('NFKD', texto)
    return ''.join([c for c in texto if not unicodedata.combining(c)])

# Cache para evitar relecturas innecesarias de planilla anterior
_cache_planilla_anterior = {}

# Cache para resultados de procesamiento (agrupar_y_calcular, construir_tabla)
_cache_agrupacion = {}
_cache_tablas_construidas = {}

def obtener_header_planilla_con_cache(ruta_planilla):
    try:
        # Leer solo las primeras 10 filas para encontrar el header (más eficiente)
        df_raw = pd.read_excel(ruta_planilla, sheet_name="Primera carga académica", header=None, nrows=10)
        
        for idx, fila in df_raw.iterrows():
            if "Docente" in fila.values and "Curso" in fila.values:
                return idx
                
        return None
    except Exception as e:
        print(f"Error al obtener header: {e}")
        return None

def leer_planilla_anterior_con_cache(ruta_planilla):
    # Verificar si el archivo ya está en cache
    if ruta_planilla in _cache_planilla_anterior:
        return _cache_planilla_anterior[ruta_planilla]
    
    try:
        # Obtener la fila de header de manera optimizada
        fila_header = obtener_header_planilla_con_cache(ruta_planilla)
        if fila_header is None:
            fila_header = 6  # Valor por defecto si no se encuentra
            
        # Solo leer si no está en cache
        planilla_anterior = pd.read_excel(ruta_planilla, sheet_name="Primera carga académica", header=fila_header)
        
        # Guardar en cache para futuros accesos
        _cache_planilla_anterior[ruta_planilla] = planilla_anterior
        
        return planilla_anterior
    except Exception as e:
        print(f"Error al leer planilla anterior: {e}")
        return pd.DataFrame()

def limpiar_cache_planilla():
    global _cache_planilla_anterior
    _cache_planilla_anterior = {}

def limpiar_cache_procesamiento():
    global _cache_agrupacion, _cache_tablas_construidas
    _cache_agrupacion = {}
    _cache_tablas_construidas = {}

def generar_key_datos(df, col_curso):
    try:
        # Crear un hash basado en el contenido de datos relevantes
        datos_relevantes = df[['docente', col_curso]].copy()
        datos_str = datos_relevantes.to_string()
        key = f"{col_curso}_{hash(datos_str)}"
        return key
    except:
        # Fallback: usar shape y columnas como key menos preciso
        return f"{col_curso}_{df.shape[0]}_{df.shape[1]}"

def filtrar_combinaciones_optimizado(datos, combinacion_anterior):
    if not combinacion_anterior:
        return datos
    
    # Convertir el set de combinaciones a DataFrame para hacer merge eficiente
    combinaciones_df = pd.DataFrame(list(combinacion_anterior), columns=['docente', 'Curso'])

    merged = datos.merge(combinaciones_df, on=['docente', 'Curso'], how='left', indicator=True)
    
    # Filtrar solo las que NO están en la planilla anterior
    datos_filtrados = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
    
    return datos_filtrados