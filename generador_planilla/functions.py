import pandas as pd
import os
from fuzzywuzzy import process, fuzz
import unicodedata

# ----------------------- CARGAR ARCHIVO -----------------------

def cargar_archivo(ruta):
    extension = os.path.splitext(ruta)[-1].lower()
    if extension == ".csv":
        # Intentar lectura normal primero con detección automática de delimitador
        try:
            # Primero intentar con punto y coma (;)
            df = pd.read_csv(ruta, sep=';', encoding='utf-8-sig')
            if len(df.columns) > 1:
                return df
            
            # Si solo tiene una columna, intentar con coma
            df = pd.read_csv(ruta, sep=',', encoding='utf-8-sig')
            if len(df.columns) > 1:
                return df
            
            # Si todavía no funciona, usar el parser personalizado
            return parsear_csv_comillas_dobles(ruta)
        except Exception as e:
            print(f"⚠️ Error al leer CSV con pandas: {e}")
            return parsear_csv_comillas_dobles(ruta)
    elif extension in [".xls", ".xlsx"]:
        return pd.read_excel(ruta)
    else:
        raise ValueError(f"Formato no soportado: {extension}")

def parsear_csv_comillas_dobles(ruta):
    
    with open(ruta, 'r', encoding='utf-8') as f:
        lineas = f.readlines()
    
    datos_parseados = []
    
    for idx, linea in enumerate(lineas):
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
        
        # Verificar que todas las filas tengan el mismo número de columnas que el header
        max_cols = len(headers)
        filas_ajustadas = []
        for fila in filas:
            if len(fila) < max_cols:
                # Rellenar con valores vacíos si faltan columnas
                fila = fila + [''] * (max_cols - len(fila))
            elif len(fila) > max_cols:
                # Truncar si hay más columnas
                fila = fila[:max_cols]
            filas_ajustadas.append(fila)
        
        print(f"✅ Creando DataFrame con {len(headers)} columnas y {len(filas_ajustadas)} filas")
        return pd.DataFrame(filas_ajustadas, columns=headers)
    
    return pd.DataFrame()

def limpiar_docentes(df, col):
    df[col] = df[col].astype(str).str.strip()
    return df

def procesar_csv_nuevo_formato(df):
    """
    Procesa el DataFrame con el nuevo formato de columnas del CSV y lo convierte
    al formato interno esperado por el sistema.
    
    Nuevas columnas esperadas:
    - Programa Educativo
    - Detalle Curso
    - Modalidad
    - Sede
    - Horario Completo
    - Docente
    - Fecha Inicio Clases
    - Fecha Fin Clases
    - Total Matriculados
    """
    import re
    
    df_procesado = df.copy()
    
    # Mapeo directo de columnas
    mapeo_columnas = {
        'Modalidad': 'modalidad',
        'Sede': 'sede',
        'Docente': 'docente',
        'Total Matriculados': 'matriculados'
    }
    
    # Renombrar columnas directas
    for col_origen, col_destino in mapeo_columnas.items():
        if col_origen in df_procesado.columns:
            df_procesado[col_destino] = df_procesado[col_origen]
    
    # Procesar "Detalle Curso" para extraer idioma, nivel y ciclo
    if 'Detalle Curso' in df_procesado.columns:
        def extraer_info_curso(detalle):
            if pd.isna(detalle):
                return pd.Series({'idioma': '', 'nivel': '', 'ciclo': ''})
            
            detalle_str = str(detalle).strip()
            # Patrón: "Idioma Nivel Número"
            # Ejemplos: "Alemán Básico 2", "Francés Básico 1", "Inglés Avanzado 3"
            patron = r'^([A-Za-zÀ-ÿ]+)\s+([A-Za-zÀ-ÿ]+)\s+(\d+)$'
            match = re.match(patron, detalle_str)
            
            if match:
                idioma = match.group(1)
                nivel = match.group(2)
                ciclo = match.group(3)
                return pd.Series({'idioma': idioma, 'nivel': nivel, 'ciclo': ciclo})
            else:
                # Si no coincide con el patrón, intentar una extracción más flexible
                partes = detalle_str.split()
                if len(partes) >= 3:
                    # Último elemento debe ser número (ciclo)
                    ciclo = partes[-1]
                    # Penúltimo elemento es el nivel
                    nivel = partes[-2]
                    # Todo lo demás es el idioma
                    idioma = ' '.join(partes[:-2])
                    return pd.Series({'idioma': idioma, 'nivel': nivel, 'ciclo': ciclo})
                else:
                    return pd.Series({'idioma': detalle_str, 'nivel': '', 'ciclo': ''})
        
        df_info = df_procesado['Detalle Curso'].apply(extraer_info_curso)
        df_procesado['idioma'] = df_info['idioma']
        df_procesado['nivel'] = df_info['nivel']
        df_procesado['ciclo'] = df_info['ciclo']
    
    # Procesar "Horario Completo" para extraer dias, horainicio y horafin
    if 'Horario Completo' in df_procesado.columns:
        def extraer_info_horario(horario):
            if pd.isna(horario):
                return pd.Series({'dias': '', 'horainicio': '', 'horafin': ''})
            
            horario_str = str(horario).strip()
            
            # Mapeo de días en español a inglés
            dias_map = {
                'Lun': 'MONDAY',
                'Mar': 'TUESDAY',
                'Mie': 'WEDNESDAY',
                'Mié': 'WEDNESDAY',
                'Jue': 'THURSDAY',
                'Vie': 'FRIDAY',
                'Sab': 'SATURDAY',
                'Sáb': 'SATURDAY',
                'Dom': 'SUNDAY'
            }
            
            # Patrón: "Día1,Día2,Día3 HH:MM am/pm - HH:MM am/pm"
            # Ejemplo: "Lun,Mar,Mie,Jue 07:30 pm - 09:30 pm"
            try:
                # Separar días del rango de horas
                if ' ' in horario_str:
                    partes = horario_str.split(' ', 1)
                    dias_str = partes[0]
                    horas_str = partes[1] if len(partes) > 1 else ''
                    
                    # Procesar días
                    dias_lista = [d.strip() for d in dias_str.split(',')]
                    dias_ingles = [dias_map.get(d, d.upper()) for d in dias_lista]
                    dias_formateado = '{' + ','.join(dias_ingles) + '}'
                    
                    # Procesar horas
                    if '-' in horas_str:
                        horas_partes = horas_str.split('-')
                        hora_inicio_str = horas_partes[0].strip()
                        hora_fin_str = horas_partes[1].strip() if len(horas_partes) > 1 else ''
                        
                        # Convertir de 12h a 24h
                        def convertir_12h_a_24h(hora_str):
                            # Patrón: "07:30 pm" o "08:00 am"
                            patron = r'(\d{1,2}):(\d{2})\s*(am|pm)'
                            match = re.search(patron, hora_str, re.IGNORECASE)
                            if match:
                                hora = int(match.group(1))
                                minuto = match.group(2)
                                periodo = match.group(3).lower()
                                
                                if periodo == 'pm' and hora != 12:
                                    hora += 12
                                elif periodo == 'am' and hora == 12:
                                    hora = 0
                                
                                return f"{hora:02d}:{minuto}:00"
                            return hora_str
                        
                        hora_inicio = convertir_12h_a_24h(hora_inicio_str)
                        hora_fin = convertir_12h_a_24h(hora_fin_str)
                        
                        return pd.Series({
                            'dias': dias_formateado,
                            'horainicio': hora_inicio,
                            'horafin': hora_fin
                        })
                
                return pd.Series({'dias': '', 'horainicio': '', 'horafin': ''})
            except Exception as e:
                print(f"Error al procesar horario '{horario_str}': {e}")
                return pd.Series({'dias': '', 'horainicio': '', 'horafin': ''})
        
        df_horario = df_procesado['Horario Completo'].apply(extraer_info_horario)
        df_procesado['dias'] = df_horario['dias']
        df_procesado['horainicio'] = df_horario['horainicio']
        df_procesado['horafin'] = df_horario['horafin']
    
    return df_procesado

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
    
    # Agregar "Intensivo" al nivel para cursos intensivos (sin duplicar filas)
    mask_intensivo = datos_transformados['modalidad'] == 'INTENSIVO VIRTUAL'
    for idx in datos_transformados[mask_intensivo].index:
        nivel_original = str(datos_transformados.loc[idx, 'nivel']).strip()
        if not nivel_original.startswith('Intensivo'):
            datos_transformados.loc[idx, 'nivel'] = f'Intensivo {nivel_original}'
    
    # Construir curso de forma vectorizada
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
        idioma = str(row['idioma']).strip()
        
        try:
            ciclo_actual = int(str(row['ciclo']).strip())
            
            # Lógica específica para Portugués
            if idioma == 'Portugués':
                nivel_sin_intensivo = nivel_original.replace('Intensivo ', '').strip()
                
                if 'Básico' in nivel_sin_intensivo or 'Basico' in nivel_sin_intensivo:
                    if ciclo_actual == 5:
                        # Cambiar a Intermedio 1
                        fila_adicional['nivel'] = 'Intensivo Intermedio'
                        fila_adicional['ciclo'] = '1'
                    else:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                elif 'Intermedio' in nivel_sin_intensivo:
                    if ciclo_actual == 4:
                        # Cambiar a Avanzado 1
                        fila_adicional['nivel'] = 'Intensivo Avanzado'
                        fila_adicional['ciclo'] = '1'
                    else:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                elif 'Avanzado' in nivel_sin_intensivo:
                    if ciclo_actual < 3:
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                    else:
                        # Para Avanzado 3, mantener el mismo ciclo (o manejar según reglas de negocio)
                        fila_adicional['ciclo'] = str(ciclo_actual + 1)
                else:
                    # Para otros niveles de portugués, incrementar normalmente
                    fila_adicional['ciclo'] = str(ciclo_actual + 1)
            else:
                # Para otros idiomas, incrementar ciclo normalmente
                fila_adicional['ciclo'] = str(ciclo_actual + 1)
                
        except (ValueError, TypeError):
            fila_adicional['ciclo'] = str(row['ciclo']) + '+1'
        
        # Agregar "Intensivo" al nivel de la fila adicional si no lo tiene
        if not str(fila_adicional['nivel']).startswith('Intensivo'):
            fila_adicional['nivel'] = f'Intensivo {fila_adicional["nivel"]}'
        
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
    
    agrupado = agrupado.merge(datos_docentes[['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N_Ruc', 'Estado']], on='Docente', how='left')
    agrupado['Curso Dictado'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 28
    agrupado['Diseño de Examenes'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 4
    return agrupado


def agregar_servicio_coordinacion(df, ruta_coordinacion, normalizar_texto, datos_docentes=None):
    if not os.path.exists(ruta_coordinacion):
        df['Servicio Actualización'] = 0
        df['Horas_Total'] = 0
        return df
    
    try:
        # Leer el archivo Excel
        coordinacion_df = cargar_excel_con_cache(ruta_coordinacion, sheet_name=0, header=0)
        
        # Detectar automáticamente las columnas importantes
        columna_docente = None
        columna_horas = None
        
        
        # Buscar columna de docente de manera más flexible
        for col in coordinacion_df.columns:
            col_lower = str(col).lower().strip()
            if any(patron in col_lower for patron in ['docente', 'profesor', 'teacher', 'instructor']):
                columna_docente = col
                print(f"✅ Columna de docente encontrada: '{col}'")
                break
        
        if columna_docente is None:
            print("⚠️ No se encontró columna de docente en archivo de coordinación")
            df['Servicio Actualización'] = 0
            df['Horas_Total'] = 0
            return df
        
        # Buscar columna de horas de manera más flexible - PRIORIDAD A HORAS TOTALES
        prioridad_horas = [
            'horas totales', 'horas_totales', 'total', 'totales',  # Máxima prioridad
            'hora total', 'total horas', 'total_horas',
            'horas', 'hora', 'hour', 'tiempo', 'time',  # Prioridad media
            'horas semanales', 'horas_semanales', 'semanal'  # Menor prioridad
        ]
        
        for patron in prioridad_horas:
            for col in coordinacion_df.columns:
                col_lower = str(col).lower().strip()
                if patron in col_lower:
                    columna_horas = col
                    break
            if columna_horas:  # Si encontró una columna, salir del loop externo
                break
        
        if columna_horas is None:
            print("⚠️ No se encontró columna de horas en archivo de coordinación")
            print(f"Columnas disponibles: {list(coordinacion_df.columns)}")
            df['Servicio Actualización'] = 0
            df['Horas_Total'] = 0
            return df
        
        
        # Limpiar y procesar datos - REMOVER NaN Y VALORES VACÍOS
        coordinacion_df = coordinacion_df.dropna(subset=[columna_docente, columna_horas])
        coordinacion_df = coordinacion_df[coordinacion_df[columna_docente].astype(str).str.strip() != '']
        coordinacion_df = coordinacion_df[coordinacion_df[columna_docente].astype(str).str.lower() != 'nan']
        
        coordinacion_df[columna_docente] = coordinacion_df[columna_docente].astype(str).str.strip()
        
        # EXTRAER SOLO LOS NÚMEROS DE LA COLUMNA DE HORAS (ej: "32 horas" -> 32)
        def extraer_numero_horas(texto):
            import re
            if pd.isna(texto):
                return 0
            # Buscar el primer número en el texto
            numeros = re.findall(r'\d+', str(texto))
            return int(numeros[0]) if numeros else 0
        
        coordinacion_df[columna_horas] = coordinacion_df[columna_horas].apply(extraer_numero_horas)
        
        # Agrupar por docente (sumar horas si aparece múltiples veces)
        coordinacion_agrupado = coordinacion_df[[columna_docente, columna_horas]].groupby(columna_docente, as_index=False)[columna_horas].sum()
        coordinacion_agrupado = coordinacion_agrupado.rename(columns={columna_docente: 'Docente_Original', columna_horas: 'Horas_Total'})
        
        # Aplicar normalización para matching
        coordinacion_agrupado['docente_norm'] = coordinacion_agrupado['Docente_Original'].apply(normalizar_texto)
        
        # Si datos_docentes está disponible, calcular montos por categoría
        if datos_docentes is not None:
            # Normalizar nombres en datos_docentes también
            datos_docentes_temp = datos_docentes.copy()
            datos_docentes_temp['docente_norm'] = datos_docentes_temp['Docente'].apply(normalizar_texto)
            
            # Hacer merge con datos_docentes para obtener categoría
            coordinacion_con_categoria = coordinacion_agrupado.merge(
                datos_docentes_temp[['docente_norm', 'Docente', 'Categoria (Monto)']],
                on='docente_norm', how='left'
            )
            
            # Calcular monto: Horas * Categoria (Monto) * 1 (tarifa base por hora de coordinación)
            coordinacion_con_categoria['Monto_Coordinacion'] = (
                coordinacion_con_categoria['Horas_Total'] * 
                coordinacion_con_categoria['Categoria (Monto)'].fillna(0)
            )
            
            # Usar el nombre corregido del merge
            coordinacion_con_categoria['Docente_Correcto'] = coordinacion_con_categoria['Docente'].fillna(
                coordinacion_con_categoria['Docente_Original']
            )
            
        else:
            # Si no hay datos_docentes, usar un monto fijo por hora
            coordinacion_con_categoria = coordinacion_agrupado.copy()
            coordinacion_con_categoria['Monto_Coordinacion'] = coordinacion_agrupado['Horas_Total'] * 50  # Monto fijo
            coordinacion_con_categoria['Docente_Correcto'] = coordinacion_agrupado['Docente_Original']
        
        # Preparar datos para merge con df principal
        coordinacion_final = coordinacion_con_categoria[['docente_norm', 'Monto_Coordinacion', 'Horas_Total', 'Docente_Correcto']].copy()
        
        # Si df no tiene docente_norm, crearlo
        if 'docente_norm' not in df.columns:
            df['docente_norm'] = df['Docente'].apply(normalizar_texto)
        
        # Hacer merge con df principal
        merge_result = df.merge(coordinacion_final, on='docente_norm', how='left')
        
        # AGREGAR DOCENTES DE COORDINACIÓN QUE NO ESTÁN EN DF (similar a examen clasificación)
        docentes_coord_no_en_df = coordinacion_final[~coordinacion_final['docente_norm'].isin(df['docente_norm'])]
        
        if not docentes_coord_no_en_df.empty and datos_docentes is not None:
            
            filas_nuevas = []
            for _, row_coord in docentes_coord_no_en_df.iterrows():
                # VALIDAR QUE EL DOCENTE NO SEA VACÍO O NaN
                if pd.isna(row_coord['Docente_Correcto']) or str(row_coord['Docente_Correcto']).strip() == '' or str(row_coord['Docente_Correcto']).lower() == 'nan':
                    continue
                
                # Buscar información del docente en datos_docentes
                docente_info = datos_docentes[datos_docentes['Docente'].apply(normalizar_texto) == row_coord['docente_norm']]
                
                if not docente_info.empty:
                    docente_info = docente_info.iloc[0]
                    nueva_fila = {
                        'Docente': docente_info['Docente'],
                        'Sede': docente_info['Sede'],
                        'Categoria (Letra)': docente_info['Categoria (Letra)'],
                        'Categoria (Monto)': docente_info['Categoria (Monto)'],
                        'N_Ruc': docente_info['N_Ruc'],
                        'Estado': docente_info['Estado'],
                        'curso': '',  # Sin carga académica
                        'cantidad_cursos': 0,
                        'Curso Dictado': 0,
                        'Diseño de Examenes': 0,
                        'Examen Clasif.': 0,
                        'docente_norm': row_coord['docente_norm'],
                        'Monto_Coordinacion': row_coord['Monto_Coordinacion'],
                        'Horas_Total': row_coord['Horas_Total'],
                        'Docente_Correcto': row_coord['Docente_Correcto']
                    }
                    filas_nuevas.append(nueva_fila)
                else:
                    pass
            
            # Agregar las nuevas filas al resultado
            if filas_nuevas:
                filas_nuevas_df = pd.DataFrame(filas_nuevas)
                merge_result = pd.concat([merge_result, filas_nuevas_df], ignore_index=True)
        
        # Finalizar el procesamiento
        df = merge_result
        df['Servicio Actualización'] = df['Monto_Coordinacion'].fillna(0)
        # Preservar Horas_Total con valor por defecto 0
        if 'Horas_Total' in df.columns:
            df['Horas_Total'] = df['Horas_Total'].fillna(0)
        else:
            df['Horas_Total'] = 0
        
        # Limpiar columnas temporales (MANTENER Horas_Total para la hoja expandida)
        columnas_a_eliminar = ['docente_norm', 'Monto_Coordinacion', 'Docente_Correcto']
        for col in columnas_a_eliminar:
            if col in df.columns:
                df.drop(columns=[col], inplace=True)
        
        print(f"✅ Servicio de coordinación procesado correctamente")
        
    except Exception as e:
        print(f"⚠️ Error al procesar archivo de coordinación: {e}")
        df['Servicio Actualización'] = 0
        df['Horas_Total'] = 0
    
    return df


def agregar_examen_clasificacion(df, ruta_clasificacion, normalizar_texto, datos_docentes=None):
    if os.path.exists(ruta_clasificacion):
        try:
            # Usar cache para evitar lecturas múltiples
            clasif_df = cargar_excel_con_cache(ruta_clasificacion, sheet_name=0, header=1)
            # Verificar que existe la columna Docente
            if 'Docente' not in clasif_df.columns:
                print("⚠️ No se encontró columna 'Docente' en el archivo de clasificación")
                # Buscar columnas similares
                posibles_docentes = [col for col in clasif_df.columns if 'docente' in str(col).lower()]
                if posibles_docentes:
                    clasif_df = clasif_df.rename(columns={posibles_docentes[0]: 'Docente'})
                else:
                    df['Examen Clasif.'] = 0
                    return df
            
            # LIMPIAR VALORES NaN/VACÍOS EN LA COLUMNA DOCENTE ANTES DE PROCESAR
            clasif_df = clasif_df.dropna(subset=['Docente'])
            clasif_df = clasif_df[clasif_df['Docente'].astype(str).str.strip() != '']
            clasif_df = clasif_df[clasif_df['Docente'].astype(str).str.lower() != 'nan']
            
            # Buscar la columna de Monto de manera más específica
            columna_monto = None
            
            # Prioridad 1: Buscar exactamente 'Monto Total'
            if 'Monto Total' in clasif_df.columns:
                columna_monto = 'Monto Total'
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
            
            # Si ya existe la columna docente_norm en df, la usamos; sino la creamos
            if 'docente_norm' not in df.columns:
                df['docente_norm'] = df['Docente'].apply(normalizar_texto)
            
            # Hacer el merge con la columna correcta (left para mantener todos los de df)
            merge_result = df.merge(clasif_df[['docente_norm', columna_monto]], on='docente_norm', how='left')
            
            # NUEVA FUNCIONALIDAD: Agregar docentes del examen de clasificación que NO están en df
            docentes_clasif_no_en_df = clasif_df[~clasif_df['docente_norm'].isin(df['docente_norm'])]
            
            if not docentes_clasif_no_en_df.empty and datos_docentes is not None:
                print(f"📋 Agregando {len(docentes_clasif_no_en_df)} docentes del examen de clasificación sin carga académica...")
                
                # Para cada docente del examen que no está en df, crear una fila nueva
                filas_nuevas = []
                for _, row_clasif in docentes_clasif_no_en_df.iterrows():
                    # VALIDAR QUE EL DOCENTE NO SEA VACÍO O NaN
                    if pd.isna(row_clasif['Docente']) or str(row_clasif['Docente']).strip() == '' or str(row_clasif['Docente']).lower() == 'nan':
                        continue
                    
                    # Buscar información del docente en datos_docentes
                    docente_info = datos_docentes[datos_docentes['Docente'].apply(normalizar_texto) == row_clasif['docente_norm']]
                    
                    if not docente_info.empty:
                        docente_info = docente_info.iloc[0]
                        nueva_fila = {
                            'Docente': docente_info['Docente'],
                            'Sede': docente_info['Sede'],
                            'Categoria (Letra)': docente_info['Categoria (Letra)'],
                            'Categoria (Monto)': docente_info['Categoria (Monto)'],
                            'N_Ruc': docente_info['N_Ruc'],
                            'Estado': docente_info['Estado'],
                            'curso': '',  # Sin carga académica
                            'cantidad_cursos': 0,
                            'Curso Dictado': 0,
                            'Diseño de Examenes': 0,
                            'docente_norm': row_clasif['docente_norm'],
                            columna_monto: row_clasif[columna_monto]
                        }
                        filas_nuevas.append(nueva_fila)
                    else:
                        pass
                
                # Agregar las nuevas filas al resultado
                if filas_nuevas:
                    filas_nuevas_df = pd.DataFrame(filas_nuevas)
                    merge_result = pd.concat([merge_result, filas_nuevas_df], ignore_index=True)
            
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

def construir_tabla_planilla(df, es_enero=False, monto_bono=0):
    # Asegurar que los campos necesarios existen y tienen valores por defecto
    df = df.copy()
    campos_requeridos = {
        'curso': '',
        'Curso Dictado': 0,
        'cantidad_cursos': 0,
        'Diseño de Examenes': 0,
        'Examen Clasif.': 0,
        'Servicio Actualización': 0  # Nueva columna
    }
    
    for campo, valor_default in campos_requeridos.items():
        if campo not in df.columns:
            df[campo] = valor_default
        else:
            df[campo] = df[campo].fillna(valor_default)
    
    # Construir diccionario base de columnas
    columnas_tabla = {
        'N°': range(1, len(df) + 1),
        'Docente': df['Docente'],
        'Sede': df['Sede'],
        'Categoria (Letra)': df['Categoria (Letra)'],
        'Categoria (Monto)': df['Categoria (Monto)'],
        'N_Ruc': df['N_Ruc'],
        'Curso': df['curso'],
        'Curso Dictado': df['Curso Dictado']
    }
    
    # Agregar columna Bono solo si es enero
    if es_enero:
        columnas_tabla['Bono'] = monto_bono
    
    # Continuar con el resto de columnas
    columnas_tabla.update({
        'Extra Curso': 0,
        'Cantidad Cursos': df['cantidad_cursos'],
        'Diseño de Examenes': df['Diseño de Examenes'],
        'Examen Clasif.': df['Examen Clasif.'],
        'Servicio Actualización': df['Servicio Actualización'],
        'Total Pago S/.': 0,
        'Estado': df['Estado']
    })
    
    tabla = pd.DataFrame(columnas_tabla)
    
    # AGREGAR SERVICIO DE ACTUALIZACIÓN A LA COLUMNA CURSO CUANDO CORRESPONDA
    def agregar_servicio_a_curso(row):
        curso_actual = str(row['Curso']).strip()
        servicio_actualizacion = row['Servicio Actualización']
        
        # Si tiene servicio de actualización (mayor a 0)
        if servicio_actualizacion > 0:
            texto_servicio = "Servicio de actualización de materiales de enseñanza"
            
            # Si ya tiene cursos académicos, agregar el servicio separado por ' / '
            if curso_actual and curso_actual != '' and curso_actual != 'nan':
                return f"{curso_actual} / {texto_servicio}"
            else:
                # Si no tiene carga académica, solo el servicio
                return texto_servicio
        
        # Si no tiene servicio de actualización, mantener el curso original
        return curso_actual
    
    tabla['Curso'] = tabla.apply(agregar_servicio_a_curso, axis=1)
    
    # Calcular total incluyendo Bono si existe
    if es_enero:
        tabla['Total Pago S/.'] = (tabla['Curso Dictado'] + tabla['Bono'] + tabla['Extra Curso'] + 
                                    tabla['Diseño de Examenes'] + tabla['Examen Clasif.'] + 
                                    tabla['Servicio Actualización'])
    else:
        tabla['Total Pago S/.'] = (tabla['Curso Dictado'] + tabla['Extra Curso'] + tabla['Diseño de Examenes'] + 
                                    tabla['Examen Clasif.'] + tabla['Servicio Actualización'])
    
    # Ordenar alfabéticamente por nombre de docente
    tabla = tabla.sort_values('Docente').reset_index(drop=True)
    # Reajustar la numeración después del ordenamiento
    tabla['N°'] = range(1, len(tabla) + 1)
    
    return tabla

def construir_tabla_coordinacion(ruta_coordinacion, normalizar_texto, datos_docentes):
    if not os.path.exists(ruta_coordinacion):
        return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])
    
    try:
        # Leer el archivo Excel
        coordinacion_df = cargar_excel_con_cache(ruta_coordinacion, sheet_name=0, header=0)
        
        # Detectar automáticamente las columnas importantes (misma lógica que en agregar_servicio_coordinacion)
        columna_docente = None
        columna_horas = None
        
        # Buscar columna de docente de manera más flexible
        for col in coordinacion_df.columns:
            col_lower = str(col).lower().strip()
            if any(patron in col_lower for patron in ['docente', 'profesor', 'teacher', 'instructor']):
                columna_docente = col
                break
        
        if columna_docente is None:
            print("⚠️ No se encontró columna de docente en archivo de coordinación")
            return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])
        
        # Buscar columna de horas de manera más flexible - PRIORIDAD A HORAS TOTALES
        prioridad_horas = [
            'horas totales', 'horas_totales', 'total', 'totales',  # Máxima prioridad
            'hora total', 'total horas', 'total_horas',
            'horas', 'hora', 'hour', 'tiempo', 'time',  # Prioridad media
            'horas semanales', 'horas_semanales', 'semanal'  # Menor prioridad
        ]
        
        for patron in prioridad_horas:
            for col in coordinacion_df.columns:
                col_lower = str(col).lower().strip()
                if patron in col_lower:
                    columna_horas = col
                    break
            if columna_horas:  # Si encontró una columna, salir del loop externo
                break
        
        if columna_horas is None:
            print("⚠️ No se encontró columna de horas en archivo de coordinación")
            return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])
        
        # Limpiar y procesar datos - REMOVER NaN Y VALORES VACÍOS
        coordinacion_df = coordinacion_df.dropna(subset=[columna_docente, columna_horas])
        coordinacion_df = coordinacion_df[coordinacion_df[columna_docente].astype(str).str.strip() != '']
        coordinacion_df = coordinacion_df[coordinacion_df[columna_docente].astype(str).str.lower() != 'nan']
        coordinacion_df[columna_docente] = coordinacion_df[columna_docente].astype(str).str.strip()
        
        # EXTRAER SOLO LOS NÚMEROS DE LA COLUMNA DE HORAS (ej: "32 horas" -> 32)
        def extraer_numero_horas(texto):
            import re
            if pd.isna(texto):
                return 0
            # Buscar el primer número en el texto
            numeros = re.findall(r'\d+', str(texto))
            return int(numeros[0]) if numeros else 0
        
        coordinacion_df[columna_horas] = coordinacion_df[columna_horas].apply(extraer_numero_horas)
        
        # Agrupar por docente (sumar horas si aparece múltiples veces)
        coordinacion_agrupado = coordinacion_df[[columna_docente, columna_horas]].groupby(columna_docente, as_index=False)[columna_horas].sum()
        coordinacion_agrupado = coordinacion_agrupado.rename(columns={columna_docente: 'Docente_Original', columna_horas: 'Horas_Total'})
        
        # Aplicar normalización para matching
        coordinacion_agrupado['docente_norm'] = coordinacion_agrupado['Docente_Original'].apply(normalizar_texto)
        
        # Normalizar nombres en datos_docentes
        datos_docentes_temp = datos_docentes.copy()
        datos_docentes_temp['docente_norm'] = datos_docentes_temp['Docente'].apply(normalizar_texto)
        
        # Hacer merge con datos_docentes para obtener categoría
        coordinacion_con_categoria = coordinacion_agrupado.merge(
            datos_docentes_temp[['docente_norm', 'Docente', 'Categoria (Letra)', 'Categoria (Monto)']],
            on='docente_norm', how='left'
        )
        
        # Usar el nombre corregido del merge, fallback al original
        coordinacion_con_categoria['Docente_Final'] = coordinacion_con_categoria['Docente'].fillna(
            coordinacion_con_categoria['Docente_Original']
        )
        coordinacion_con_categoria['Categoria (Letra)'] = coordinacion_con_categoria['Categoria (Letra)'].fillna('N/A')
        coordinacion_con_categoria['Categoria (Monto)'] = coordinacion_con_categoria['Categoria (Monto)'].fillna(0)
        
        # Calcular monto total
        coordinacion_con_categoria['Monto_Total'] = (
            coordinacion_con_categoria['Horas_Total'] * 
            coordinacion_con_categoria['Categoria (Monto)']
        )
        
        # Construir tabla final
        tabla_coordinacion = pd.DataFrame({
            'N°': range(1, len(coordinacion_con_categoria) + 1),
            'Docente': coordinacion_con_categoria['Docente_Final'],
            'Categoría (Letra)': coordinacion_con_categoria['Categoria (Letra)'],
            'Categoría (Monto)': coordinacion_con_categoria['Categoria (Monto)'],
            'Horas Totales': coordinacion_con_categoria['Horas_Total'],
            'Monto Total': coordinacion_con_categoria['Monto_Total']
        })
        
        # Ordenar alfabéticamente por docente
        tabla_coordinacion = tabla_coordinacion.sort_values('Docente').reset_index(drop=True)
        tabla_coordinacion['N°'] = range(1, len(tabla_coordinacion) + 1)
        
        print(f"✅ Tabla de coordinación creada con {len(tabla_coordinacion)} docentes")
        return tabla_coordinacion
        
    except Exception as e:
        print(f"⚠️ Error al crear tabla de coordinación: {e}")
        return pd.DataFrame(columns=['N°', 'Docente', 'Categoría (Letra)', 'Categoría por Hora', 'Horas Totales', 'Monto Total'])

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


# ================================= FUNCIONES CON USO DE MEMORIA CACHE ===================================================

# Cache simple para archivos Excel
_excel_cache = {}

# Cache para evitar relecturas innecesarias de planilla anterior
_cache_planilla_anterior = {}
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
    

def limpiar_cache_excel():
    global _excel_cache
    _excel_cache.clear()

# ----------------------- CARGAR ARCHIVO CON CACHE -----------------------

def cargar_excel_con_cache(ruta, sheet_name=0, header='infer'):
    cache_key = f"{ruta}_{sheet_name}_{header}"
    
    if cache_key not in _excel_cache:
        if header == 'infer':
            _excel_cache[cache_key] = pd.read_excel(ruta, sheet_name=sheet_name)
        else:
            _excel_cache[cache_key] = pd.read_excel(ruta, sheet_name=sheet_name, header=header)
    
    return _excel_cache[cache_key].copy()  # Retorna una copia para evitar modificaciones accidentales

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
    
def construir_tabla_planilla_con_cache(df, es_enero=False, monto_bono=0):
    try:
        # Crear hash basado en las columnas relevantes para la tabla y parámetros adicionales
        columnas_relevantes = ['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N_Ruc', 'curso', 'cantidad_cursos', 'Curso Dictado', 'Diseño de Examenes', 'Examen Clasif.', 'Estado']
        
        df_relevante = df[columnas_relevantes]
        key_tabla = hash(df_relevante.to_string() + str(es_enero) + str(monto_bono))
        
        # Verificar cache
        if key_tabla in _cache_tablas_construidas:
            return _cache_tablas_construidas[key_tabla].copy()
        
        # Si no está en cache, construir y guardar
        tabla = construir_tabla_planilla(df, es_enero, monto_bono)
        _cache_tablas_construidas[key_tabla] = tabla.copy()
        
        return tabla
        
    except Exception as e:
        # Fallback: usar función original si hay error en cache
        return construir_tabla_planilla(df, es_enero, monto_bono)

# ================================= EXPANSIÓN DE FILAS POR CURSO CON MODALIDAD ===================================================

def expandir_filas_por_curso(agrupar_df, datos_csv_procesados):
    """
    Expande el DataFrame agrupado creando una fila por cada curso/servicio individual.
    Cada fila incluye: curso individual, modalidad específica, tipo de servicio, horas y monto.
    
    Args:
        agrupar_df: DataFrame con docentes agrupados (una fila por docente)
        datos_csv_procesados: DataFrame original con información de cada curso
    
    Returns:
        DataFrame expandido con múltiples filas por docente
    """
    filas_expandidas = []
    
    for _, row_docente in agrupar_df.iterrows():
        # Usar el nombre ORIGINAL del CSV (columna 'docente' en minúscula)
        # en lugar del nombre oficial (columna 'Docente' con mayúscula)
        docente_nombre_csv = row_docente.get('docente', row_docente.get('Docente', None))
        
        # Validar que el nombre del docente no sea None o vacío
        if pd.isna(docente_nombre_csv) or not str(docente_nombre_csv).strip():
            print(f"⚠️ Saltando fila con nombre de docente vacío o None")
            continue
        
        # Convertir a string y limpiar
        docente_nombre_limpio = str(docente_nombre_csv).strip().upper()
        
        # Obtener todos los cursos de este docente del CSV original
        cursos_docente = datos_csv_procesados[
            datos_csv_procesados['docente'].fillna('').astype(str).str.strip().str.upper() == docente_nombre_limpio
        ]
        
        # Debug: mostrar si no se encuentran cursos
        if cursos_docente.empty:
            print(f"⚠️ No se encontraron cursos para: {docente_nombre_csv} (oficial: {row_docente.get('Docente', 'N/A')})")
        
        # Para cada curso académico del docente
        for _, curso_row in cursos_docente.iterrows():
            fila_curso = row_docente.copy()
            fila_curso['Curso_Individual'] = curso_row['Curso']
            fila_curso['Modalidad_Curso'] = curso_row['modalidad']
            fila_curso['Tipo_Servicio'] = 'CURSO_DICTADO'
            fila_curso['Horas_Servicio'] = 28
            # Validar categoria_monto antes de multiplicar
            categoria_monto = row_docente['Categoria (Monto)'] if not pd.isna(row_docente['Categoria (Monto)']) else 0
            fila_curso['Monto_Individual'] = categoria_monto * 28
            filas_expandidas.append(fila_curso)
        
        # Agregar fila para examen de clasificación si aplica
        examen_clasif = row_docente['Examen Clasif.'] if not pd.isna(row_docente['Examen Clasif.']) else 0
        if examen_clasif > 0:
            fila_examen = row_docente.copy()
            fila_examen['Curso_Individual'] = 'Examen de clasificación'
            fila_examen['Modalidad_Curso'] = 'VIRTUAL'
            fila_examen['Tipo_Servicio'] = 'EXAMEN_CLASIF'
            categoria_monto = row_docente['Categoria (Monto)'] if not pd.isna(row_docente['Categoria (Monto)']) else 1
            horas_examen = int(examen_clasif / categoria_monto) if categoria_monto > 0 else 0
            fila_examen['Horas_Servicio'] = horas_examen
            fila_examen['Monto_Individual'] = examen_clasif
            filas_expandidas.append(fila_examen)
        
        # Agregar fila para diseño de exámenes si aplica
        diseno_examenes = row_docente['Diseño de Examenes'] if not pd.isna(row_docente['Diseño de Examenes']) else 0
        if diseno_examenes > 0:
            fila_diseno = row_docente.copy()
            fila_diseno['Curso_Individual'] = 'Diseño de exámenes'
            fila_diseno['Modalidad_Curso'] = 'N/A'
            fila_diseno['Tipo_Servicio'] = 'DISENO_EXAMENES'
            categoria_monto = row_docente['Categoria (Monto)'] if not pd.isna(row_docente['Categoria (Monto)']) else 1
            horas_diseno = int(diseno_examenes / categoria_monto) if categoria_monto > 0 else 0
            fila_diseno['Horas_Servicio'] = horas_diseno
            fila_diseno['Monto_Individual'] = diseno_examenes
            filas_expandidas.append(fila_diseno)
        
        # Agregar fila para servicio de actualización si aplica
        servicio_act = row_docente['Servicio Actualización'] if not pd.isna(row_docente['Servicio Actualización']) else 0
        if servicio_act > 0:
            fila_servicio = row_docente.copy()
            fila_servicio['Curso_Individual'] = 'Servicio de actualización de materiales de enseñanza'
            fila_servicio['Modalidad_Curso'] = 'VIRTUAL'
            fila_servicio['Tipo_Servicio'] = 'SERVICIO_ACTUALIZACION'
            # Usar las horas reales del archivo de coordinación, NO calcularlas
            horas_total = row_docente.get('Horas_Total', 0)
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
        columnas_resto = [col for col in agrupar_df.columns]
        return pd.DataFrame(columns=columnas_nuevas + columnas_resto)
    
    df_expandido = pd.DataFrame(filas_expandidas)
    
    # ============= MEJORAS DE ORGANIZACIÓN =============
    
    # 1. Definir orden lógico para tipos de servicio
    orden_tipo_servicio = {
        'CURSO_DICTADO': 1,
        'DISENO_EXAMENES': 2,
        'EXAMEN_CLASIF': 3,
        'SERVICIO_ACTUALIZACION': 4
    }
    df_expandido['_orden_servicio'] = df_expandido['Tipo_Servicio'].map(orden_tipo_servicio).fillna(99)
    
    # Mapeo a nombres descriptivos en español
    nombres_servicio = {
        'CURSO_DICTADO': 'Curso Académico',
        'DISENO_EXAMENES': 'Diseño de Exámenes',
        'EXAMEN_CLASIF': 'Examen de Clasificación',
        'SERVICIO_ACTUALIZACION': 'Servicio de Actualización'
    }
    df_expandido['Tipo_Servicio_Desc'] = df_expandido['Tipo_Servicio'].map(nombres_servicio)
    
    # 2. Ordenar por: Docente (alfabético) -> Tipo de Servicio (orden lógico) -> Curso Individual
    df_expandido = df_expandido.sort_values(
        by=['Docente', '_orden_servicio', 'Curso_Individual'],
        ascending=[True, True, True]
    ).reset_index(drop=True)
    
    # 3. Agregar numeración global y por docente
    df_expandido.insert(0, 'N°', range(1, len(df_expandido) + 1))
    
    # Agregar contador de servicio por docente (1, 2, 3, etc. para cada docente)
    df_expandido['Servicio_Nro'] = df_expandido.groupby('Docente').cumcount() + 1
    
    # 4. Reorganizar columnas de forma más lógica
    # Primero: Identificación y servicio, luego detalles, al final datos administrativos repetidos
    columnas_principales = [
        'N°',
        'Docente',
        'Servicio_Nro',
        'Tipo_Servicio_Desc',
        'Curso_Individual',
        'Modalidad_Curso',
        'Horas_Servicio',
        'Monto_Individual'
    ]
    
    columnas_resumen = [
        'cantidad_cursos',
        'Curso Dictado',
        'Disenio_examenes',
        'Examen_clasif',
        'Servicio_actualizacion'
    ]
    
    columnas_administrativas = [
        'Sede',
        'Categoria_letra',
        'Categoria_monto',
        'N_Ruc',
        'Estado_docente',
        'Docente_idioma',
        'Numero_dni',
        'Numero_celular',
        'Domicilio_docente',
        'Correo_personal',
        'Nro_contrato'
    ]
    
    # Construir orden final de columnas
    columnas_ordenadas = columnas_principales + columnas_resumen + columnas_administrativas
    
    # Agregar cualquier columna que falte al final
    columnas_restantes = [col for col in df_expandido.columns 
                          if col not in columnas_ordenadas and col not in ['_orden_servicio', 'docente', 'curso', 'Tipo_Servicio']]
    
    columnas_finales = [col for col in columnas_ordenadas if col in df_expandido.columns] + columnas_restantes
    
    # Eliminar columnas temporales y redundantes
    df_expandido = df_expandido.drop(columns=['_orden_servicio', 'Tipo_Servicio'], errors='ignore')
    
    # Aplicar orden final
    df_expandido = df_expandido[columnas_finales]
    
    return df_expandido
