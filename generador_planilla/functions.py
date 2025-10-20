import pandas as pd
import os
from fuzzywuzzy import process, fuzz
import unicodedata

# ----------------------- CARGAR ARCHIVO -----------------------

def cargar_archivo(ruta):
    extension = os.path.splitext(ruta)[-1].lower()
    if extension == ".csv":
        # Intentar lectura normal primero
        try:
            df = pd.read_csv(ruta)
            # Si tiene solo una columna, probablemente está mal formateado
            if len(df.columns) == 1:
                return parsear_csv_comillas_dobles(ruta)
            return df
        except Exception:
            return parsear_csv_comillas_dobles(ruta)
    elif extension in [".xls", ".xlsx"]:
        return pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {extension}")

def parsear_csv_comillas_dobles(ruta):
    """Parsea CSV con formato de comillas dobles anidadas"""
    
    with open(ruta, 'r', encoding='utf-8') as f:
        lineas = f.readlines()
    
    datos_parseados = []
    
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue
            
        # Remover comillas externas si existen
        if linea.startswith('"') and linea.endswith('"'):
            linea = linea[1:-1]
        
        # Dividir la línea considerando el formato especial
        campos = []
        campo_actual = ""
        dentro_comillas = False
        dentro_llaves = 0
        i = 0
        
        while i < len(linea):
            char = linea[i]
            
            # Manejar comillas dobles ""
            if i < len(linea) - 1 and linea[i:i+2] == '""':
                if not dentro_comillas:
                    dentro_comillas = True
                    i += 2
                    continue
                else:
                    dentro_comillas = False
                    i += 2
                    continue
            
            # Manejar llaves para días
            elif char == '{':
                dentro_llaves += 1
                campo_actual += char
            elif char == '}':
                dentro_llaves -= 1
                campo_actual += char
            
            # Manejar comas separadoras
            elif char == ',' and not dentro_comillas and dentro_llaves == 0:
                campos.append(campo_actual.strip())
                campo_actual = ""
            else:
                campo_actual += char
            
            i += 1
        
        # Agregar el último campo
        if campo_actual:
            campos.append(campo_actual.strip())
        
        # Limitar a 13 columnas
        if len(campos) > 13:
            campos = campos[:13]
            
        datos_parseados.append(campos)
    
    if datos_parseados:
        headers = datos_parseados[0]
        filas = datos_parseados[1:]
        return pd.DataFrame(filas, columns=headers)
    
    return pd.DataFrame()

def limpiar_docentes(df, col):
    df[col] = df[col].astype(str).str.strip()
    return df

# ---------------- TRANSFORMACIONES DE DATOS ----------------

def traducir_dias(dias_raw: str) -> str:
    dias_dict = {'MONDAY': 'Lun', 'TUESDAY': 'Mar', 'WEDNESDAY': 'Mié', 'THURSDAY': 'Jue', 'FRIDAY': 'Vie', 'SATURDAY': 'Sáb', 'SUNDAY': 'Dom'}
    dias = dias_raw.strip('{}').split(',') if isinstance(dias_raw, str) else []
    return ', '.join(dias_dict.get(d.strip().upper(), d.strip()) for d in dias)

def normalizar_texto(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize('NFKD', texto)
    return ''.join([c for c in texto if not unicodedata.combining(c)])


def procesar_ediciones_idioma(datos):
    datos_procesados = datos.copy()
    
    # Identificar filas donde la columna idioma empieza con "Edición"
    mask_edicion = datos_procesados['idioma'].astype(str).str.startswith('Edición', na=False)
    
    # Para las filas que empiecen con "Edición", agregar "Inglés" al final si no lo tienen
    for idx in datos_procesados[mask_edicion].index:
        idioma_actual = str(datos_procesados.loc[idx, 'idioma']).strip()
        if not idioma_actual.endswith('Inglés'):
            datos_procesados.loc[idx, 'idioma'] = idioma_actual + ' Inglés'
    
    return datos_procesados

def aplicar_transformaciones_base(datos):
    datos_transformados = datos.copy()
    
    # NUEVO: Procesar celdas de "Edición" para agregar "Inglés"
    datos_transformados = procesar_ediciones_idioma(datos_transformados)

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
    
    # NUEVO: Procesar cursos intensivos (duplicar filas y agregar "Intensivo" al nivel)
    datos_transformados = procesar_cursos_intensivos(datos_transformados)
    
    # Construir curso de forma vectorizada (se ejecuta después del procesamiento intensivo)
    datos_transformados['Curso'] = (
        datos_transformados['idioma'].astype(str) + ' ' + 
        datos_transformados['nivel'].astype(str) + ' ' + 
        datos_transformados['ciclo'].astype(str)
    )
    return datos_transformados
    
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

def procesar_cursos_intensivos(datos):
    datos_procesados = datos.copy()
    
    # Identificar filas con modalidad INTENSIVO VIRTUAL
    mask_intensivo = datos_procesados['modalidad'] == 'INTENSIVO VIRTUAL'
    
    if not mask_intensivo.any():
        return datos_procesados
    
    filas_intensivas = datos_procesados[mask_intensivo].copy()
    filas_adicionales = []
    
    # Procesar cada fila intensiva
    for idx, row in filas_intensivas.iterrows():
        nivel_original = str(row['nivel']).strip()
        if not nivel_original.startswith('Intensivo'):
            datos_procesados.loc[idx, 'nivel'] = f'Intensivo {nivel_original}'
        
        fila_adicional = row.copy()
        try:
            ciclo_actual = int(str(row['ciclo']).strip())
            fila_adicional['ciclo'] = str(ciclo_actual + 1)
        except (ValueError, TypeError):
            fila_adicional['ciclo'] = str(row['ciclo']) + '+1'
        
        # Agregar "Intensivo" al nivel de la fila adicional
        if not nivel_original.startswith('Intensivo'):
            fila_adicional['nivel'] = f'Intensivo {nivel_original}'
        
        filas_adicionales.append(fila_adicional)
    
    # Agregar las filas adicionales al DataFrame
    if filas_adicionales:
        filas_adicionales_df = pd.DataFrame(filas_adicionales)
        datos_procesados = pd.concat([datos_procesados, filas_adicionales_df], ignore_index=True)
    
    # Reconstruir la columna Curso para todas las filas afectadas
    mask_todas_intensivas = datos_procesados['modalidad'] == 'INTENSIVO VIRTUAL'
    datos_procesados.loc[mask_todas_intensivas, 'Curso'] = (
        datos_procesados.loc[mask_todas_intensivas, 'idioma'].astype(str) + ' ' + 
        datos_procesados.loc[mask_todas_intensivas, 'nivel'].astype(str) + ' ' + 
        datos_procesados.loc[mask_todas_intensivas, 'ciclo'].astype(str)
    )
    
    return datos_procesados

def crear_mapeo_fuzzy(nombres_input, nombres_base, umbral=85):
    mapeo = {}
    for nombre in nombres_input:
        if nombre not in mapeo:  # Evitar recálculos
            resultado = process.extractOne(nombre, nombres_base, scorer=fuzz.token_sort_ratio)
            mapeo[nombre] = resultado[0] if resultado[1] >= umbral else None
    return mapeo

# ------------- CÁLCULO DE MONTOS -------------

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


def agregar_examen_clasificacion(df, ruta_clasificacion, normalizar_texto):
    if os.path.exists(ruta_clasificacion):
        try:
            # Leer archivo de clasificación
            clasif_df = pd.read_excel(ruta_clasificacion, sheet_name=0, header=1)
            # Verificar que existe la columna Docente
            if 'Docente' not in clasif_df.columns:
                print("⚠️ No se encontró columna 'Docente' en el archivo de clasificación")
                # Buscar columnas similares
                posibles_docentes = [col for col in clasif_df.columns if 'docente' in str(col).lower()]
                if posibles_docentes:
                    print(f"Posibles columnas de docente: {posibles_docentes}")
                    clasif_df = clasif_df.rename(columns={posibles_docentes[0]: 'Docente'})
                else:
                    df['Examen Clasif.'] = 0
                    return df
            
            # Buscar la columna de Monto de manera más específica
            columna_monto = None
            
            # Prioridad 1: Buscar exactamente 'Monto'
            if 'Monto' in clasif_df.columns:
                columna_monto = 'Monto'
            else:
                # Prioridad 2: Buscar variaciones de 'Monto'
                for col in clasif_df.columns:
                    if str(col).lower().strip() == 'monto':
                        columna_monto = col
                        break
            
            if columna_monto is None:
                print("⚠️ No se encontró columna de monto en el archivo de clasificación")
                df['Examen Clasif.'] = 0
                return df
            
            # Procesar los datos
            clasif_df['Docente'] = clasif_df['Docente'].astype(str).str.strip()
            clasif_df['docente_norm'] = clasif_df['Docente'].apply(normalizar_texto)
            df['docente_norm'] = df['Docente'].apply(normalizar_texto)
            
            # Hacer el merge con la columna correcta
            merge_result = df.merge(clasif_df[['docente_norm', columna_monto]], on='docente_norm', how='left')
            
            df = merge_result
            df['Examen Clasif.'] = df[columna_monto].fillna(0)
            df.drop(columns=['docente_norm', columna_monto], inplace=True)
            
        except Exception as e:
            print(f"⚠️ Error al leer archivo de clasificación: {e}")
            df['Examen Clasif.'] = 0
    else:
        df['Examen Clasif.'] = 0
    return df

# --------------------------- CONSTRUCCIÓN DE TABLAS ---------------------------

def construir_tabla_planilla(df):
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

def construir_tabla_carga_academica(datos, estado_planilla):
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

def filtrar_combinaciones_optimizado(datos, combinacion_anterior):
    if not combinacion_anterior:
        return datos
    
    # Convertir el set de combinaciones a DataFrame para hacer merge eficiente
    combinaciones_df = pd.DataFrame(list(combinacion_anterior), columns=['docente', 'Curso'])

    merged = datos.merge(combinaciones_df, on=['docente', 'Curso'], how='left', indicator=True)
    
    # Filtrar solo las que NO están en la planilla anterior
    datos_filtrados = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
    
    return datos_filtrados


# ===============================================================================================================
# NOTA: Las funciones de cache manual han sido eliminadas y reemplazadas por el patrón Repository.
# El cache ahora se maneja centralizadamente a través de repositories/cache_repository.py
# ===============================================================================================================
