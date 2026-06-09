import pandas as pd
import os
import re
from fuzzywuzzy import process, fuzz
from core.fases.utils import TextUtils
from core.planillas import cache as _cache
from core.planillas import table_builders as _table_builders
from core.planillas.transformations import generar_siguiente_curso_intensivo_desde_fila


def traducir_dias(dias_raw: str) -> str:
    dias_dict = {'MONDAY': 'Lun', 'TUESDAY': 'Mar', 'WEDNESDAY': 'Mié', 'THURSDAY': 'Jue', 'FRIDAY': 'Vie', 'SATURDAY': 'Sáb', 'SUNDAY': 'Dom'}
    dias = dias_raw.strip('{}').split(',') if isinstance(dias_raw, str) else []
    return ', '.join(dias_dict.get(d.strip().upper(), d.strip()) for d in dias)


def extraer_numero_horas(texto):
    if pd.isna(texto):
        return 0
    numeros = re.findall(r'\d+', str(texto))
    return int(numeros[0]) if numeros else 0


def preparar_coordinacion_agrupada(coordinacion_df, normalizar_texto, mostrar_columna_docente=False, mostrar_columnas_disponibles=False):
    # Detectar automáticamente las columnas importantes.
    columna_docente = None
    columna_horas = None

    for col in coordinacion_df.columns:
        col_lower = str(col).lower().strip()
        if any(patron in col_lower for patron in ['docente', 'profesor', 'teacher', 'instructor']):
            columna_docente = col
            if mostrar_columna_docente:
                print(f"✅ Columna de docente encontrada: '{col}'")
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
        for col in coordinacion_df.columns:
            col_lower = str(col).lower().strip()
            if patron in col_lower:
                columna_horas = col
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
    coordinacion_agrupado = coordinacion_agrupado.rename(columns={columna_docente: 'Docente_Original', columna_horas: 'Horas_Total'})
    coordinacion_agrupado['docente_norm'] = coordinacion_agrupado['Docente_Original'].apply(normalizar_texto)

    return coordinacion_agrupado


def formatear_numero(valor, ancho_minimo):
    if pd.isna(valor):
        return ''

    valor_str = str(valor).strip()
    if not valor_str or valor_str.lower() == 'nan':
        return ''

    if '.' in valor_str:
        try:
            valor_str = str(int(float(valor_str)))
        except (ValueError, TypeError):
            pass

    if valor_str.isdigit():
        return valor_str.zfill(max(ancho_minimo, len(valor_str)))

    return valor_str


def crear_mapeo_fuzzy(nombres_input, nombres_base, umbral=65):
    mapeo = {}
    for nombre in nombres_input:
        if nombre not in mapeo:  # Evitar recálculos
            resultado = process.extractOne(nombre, nombres_base, scorer=fuzz.token_sort_ratio)
            mapeo[nombre] = resultado[0] if resultado[1] >= umbral else None
    return mapeo

# ------------- CÁLCULO DE MONTOS -------------

def agrupar_y_calcular(df, datos_docentes, col_curso):
    df_trabajo = df.copy()

    # Cada curso intensivo genera un curso adicional (siguiente ciclo) que también se paga.
    df_trabajo['_intensivo_adicional'] = (
        df_trabajo[col_curso]
        .fillna('')
        .astype(str)
        .str.contains('intensivo', case=False, na=False)
        .astype(int)
    )

    # Agrupar por docente: una fila por docente con todos sus cursos concatenados
    agrupado = (df_trabajo.groupby('docente').agg(
        curso=(col_curso, lambda x: ' / '.join(str(v).strip() for v in x if str(v).strip() and str(v).strip().lower() != 'nan')),
        cantidad_cursos_base=(col_curso, 'count'),
        cursos_intensivos_adicionales=('_intensivo_adicional', 'sum')
    ).reset_index())

    agrupado['cantidad_cursos'] = agrupado['cantidad_cursos_base'] + agrupado['cursos_intensivos_adicionales']
    agrupado.drop(columns=['cantidad_cursos_base', 'cursos_intensivos_adicionales'], inplace=True)

    nombres_base = datos_docentes['Docente'].tolist()

    # Crear mapeo fuzzy una sola vez (solo sobre docentes únicos)
    mapeo_docentes = crear_mapeo_fuzzy(agrupado['docente'].tolist(), nombres_base, umbral=65)

    # Aplicar mapeo usando el diccionario
    agrupado['Docente'] = agrupado['docente'].map(mapeo_docentes)

    agrupado = agrupado.merge(datos_docentes[['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N_Ruc', 'Estado']], on='Docente', how='left')
    agrupado['Curso Dictado'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 28
    agrupado['Diseño de Examenes'] = agrupado['Categoria (Monto)'] * agrupado['cantidad_cursos'] * 4
    return agrupado


def _agregar_docentes_faltantes_al_merge(
    merge_result,
    docentes_faltantes_df,
    datos_docentes,
    normalizar_texto,
    campo_nombre_validacion,
    extra_builder,
):
    if docentes_faltantes_df.empty or datos_docentes is None:
        return merge_result

    datos_docentes_temp = datos_docentes.copy()
    datos_docentes_temp['docente_norm'] = datos_docentes_temp['Docente'].apply(normalizar_texto)

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
            'Docente': docente_info['Docente'],
            'Sede': docente_info['Sede'],
            'Categoria (Letra)': docente_info['Categoria (Letra)'],
            'Categoria (Monto)': docente_info['Categoria (Monto)'],
            'N_Ruc': docente_info['N_Ruc'],
            'Estado': docente_info['Estado'],
            'curso': '',
            'cantidad_cursos': 0,
            'Curso Dictado': 0,
            'Diseño de Examenes': 0,
            'Examen Clasif.': 0,
            'docente_norm': row_faltante['docente_norm'],
        }
        nueva_fila.update(extra_builder(row_faltante))
        filas_nuevas.append(nueva_fila)

    if not filas_nuevas:
        return merge_result

    filas_nuevas_df = pd.DataFrame(filas_nuevas)
    return pd.concat([merge_result, filas_nuevas_df], ignore_index=True)


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
                TextUtils.normalizar_texto(valor)
                for valor in datos_preview.iloc[i].tolist()
            }
            if 'DOCENTE' in fila_normalizada and any(
                TextUtils.normalizar_texto(patron) in fila_normalizada
                for patron in patrones_columna_monto
            ):
                header_idx = i
                break

        # Leer el archivo Excel usando la fila de encabezados detectada
        datos_df = _cache.cargar_excel_con_cache(ruta_archivo, sheet_name=0, header=header_idx)
        
        # Verificar que existe la columna Docente (con fallback a búsqueda de similares)
        if 'Docente' not in datos_df.columns:
            posibles_docentes = [col for col in datos_df.columns if 'docente' in str(col).lower()]
            if posibles_docentes:
                datos_df = datos_df.rename(columns={posibles_docentes[0]: 'Docente'})
            else:
                df[nombre_columna_resultado] = 0
                return df
        
        # Limpiar valores NaN/VACÍOS en la columna Docente
        datos_df = datos_df.dropna(subset=['Docente'])
        datos_df = datos_df[datos_df['Docente'].astype(str).str.strip() != '']
        datos_df = datos_df[datos_df['Docente'].astype(str).str.lower() != 'nan']
        
        # Buscar la columna de Monto (con patrones prioritarios)
        columna_monto = None
        for patron in patrones_columna_monto:
            if patron in datos_df.columns:
                columna_monto = patron
                break
            # Búsqueda case-insensitive si es string
            if isinstance(patron, str):
                for col in datos_df.columns:
                    if str(col).lower().strip() == patron.lower().strip():
                        columna_monto = col
                        break
                if columna_monto:
                    break
        
        if columna_monto is None:
            print(f"⚠️ No se encontró columna de monto en {ruta_archivo}")
            df[nombre_columna_resultado] = 0
            return df
        
        # Procesar y normalizar datos
        datos_df['Docente'] = datos_df['Docente'].astype(str).str.strip()
        datos_df['docente_norm'] = datos_df['Docente'].apply(normalizar_texto_fn)
        
        # Asegurar que df tiene docente_norm
        if 'docente_norm' not in df.columns:
            df['docente_norm'] = df['Docente'].apply(normalizar_texto_fn)
        
        # Hacer el merge (left para mantener todos los de df)
        merge_result = df.merge(datos_df[['docente_norm', columna_monto]], on='docente_norm', how='left')
        
        # Agregar docentes que no están en df pero sí en datos_df
        docentes_no_en_df = datos_df[~datos_df['docente_norm'].isin(df['docente_norm'])]
        
        if not docentes_no_en_df.empty and datos_docentes is not None:
            print(f"📋 Agregando {len(docentes_no_en_df)} docentes del archivo {ruta_archivo}...")
            
            merge_result = _agregar_docentes_faltantes_al_merge(
                merge_result,
                docentes_no_en_df,
                datos_docentes,
                normalizar_texto_fn,
                campo_nombre_validacion='Docente',
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
        df['Servicio Actualización'] = 0
        df['Horas_Total'] = 0
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
            df['Servicio Actualización'] = 0
            df['Horas_Total'] = 0
            return df
        
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
        
        merge_result = _agregar_docentes_faltantes_al_merge(
            merge_result,
            docentes_coord_no_en_df,
            datos_docentes,
            normalizar_texto,
            campo_nombre_validacion='Docente_Correcto',
            extra_builder=lambda row_coord: {
                'Monto_Coordinacion': row_coord['Monto_Coordinacion'],
                'Horas_Total': row_coord['Horas_Total'],
                'Docente_Correcto': row_coord['Docente_Correcto'],
            },
        )
        
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
    return _procesar_archivo_docentes_con_monto(
        df,
        ruta_clasificacion,
        normalizar_texto,
        datos_docentes,
        patrones_columna_monto=['Monto Total', 'monto'],
        nombre_columna_resultado='Examen Clasif.',
        procesar_resultado_fn=None  # Default: llenar directo con monto
    )

# --------------------------- CONSTRUCCIÓN DE TABLAS ---------------------------

def construir_tabla_planilla(df, es_enero=False, monto_bono=0):
    return _table_builders.construir_tabla_planilla(df, es_enero, monto_bono)

def construir_tabla_coordinacion(ruta_coordinacion, normalizar_texto, datos_docentes):
    return _table_builders.construir_tabla_coordinacion(
        ruta_coordinacion,
        normalizar_texto,
        datos_docentes,
        _cache.cargar_excel_con_cache,
        preparar_coordinacion_agrupada,
    )

def construir_tabla_carga_academica(datos, estado_planilla):
    return _table_builders.construir_tabla_carga_academica(datos, estado_planilla, traducir_dias)

def filtrar_combinaciones_optimizado(datos, combinacion_anterior):
    if not combinacion_anterior:
        return datos
    
    # Convertir el set de combinaciones a DataFrame para hacer merge eficiente
    combinaciones_df = pd.DataFrame(list(combinacion_anterior), columns=['docente', 'Curso'])

    merged = datos.merge(combinaciones_df, on=['docente', 'Curso'], how='left', indicator=True)
    
    # Filtrar solo las que NO están en la planilla anterior
    datos_filtrados = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
    
    return datos_filtrados


# ================================= EXPANSIÓN DE FILAS POR CURSO CON MODALIDAD ===================================================

def expandir_filas_por_curso(agrupar_df, datos_csv_procesados):
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
        'Nro_Contrato'
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


def construir_tabla_planilla_generador_resumida(agrupar_df, tabla_generador, datos_csv_procesados):
    tabla_base = tabla_generador.copy()

    if 'cantidad_cursos' not in tabla_base.columns and 'Cantidad Cursos' in tabla_base.columns:
        tabla_base['cantidad_cursos'] = tabla_base['Cantidad Cursos']

    if tabla_base.empty:
        columnas = [
            'N°', 'Docente', 'N_Ruc', 'Categoria_letra', 'Categoria_monto', 'Sede',
            'Curso_Virtual', 'Curso_Presencial', 'cantidad_cursos', 'Curso Dictado',
            'Disenio_examenes', 'Examen_clasif', 'Horas_Total', 'Servicio_actualizacion',
            'Total_pago', 'Estado_docente', 'Docente_idioma', 'Tipo_documento', 'Numero_dni', 'Numero_celular',
            'Domicilio_docente', 'Correo_personal', 'Nro_Contrato'
        ]
        return pd.DataFrame(columns=columnas)

    columnas_texto = [
        'N_Ruc', 'Categoria_letra', 'Sede', 'Estado_docente', 'Docente_idioma', 'Tipo_documento',
        'Numero_dni', 'Numero_celular', 'Domicilio_docente', 'Correo_personal', 'Nro_Contrato'
    ]
    columnas_numericas = [
        'Categoria_monto', 'cantidad_cursos', 'Curso Dictado', 'Disenio_examenes',
        'Examen_clasif', 'Horas_Total', 'Servicio_actualizacion', 'Total_pago'
    ]

    for columna in columnas_texto:
        if columna not in tabla_base.columns:
            tabla_base[columna] = ''

    for columna in columnas_numericas:
        if columna not in tabla_base.columns:
            tabla_base[columna] = 0
        tabla_base[columna] = pd.to_numeric(tabla_base[columna], errors='coerce').fillna(0)

    # Evitar que filas sin nombre de docente se agrupen en una sola fila residual.
    tabla_base = tabla_base[tabla_base['Docente'].fillna('').astype(str).str.strip() != ''].copy()

    # Agrupar por Docente: una fila por docente con montos sumados
    _primer_valor = lambda x: next((v for v in x if pd.notna(v) and str(v).strip() and str(v).strip().lower() != 'nan'), '')
    _unir_unicos = lambda x: ' / '.join(dict.fromkeys(v for v in x if pd.notna(v) and str(v).strip() and str(v).strip().lower() != 'nan'))

    agrupado_resumen = tabla_base.groupby('Docente', as_index=False).agg({
        'N_Ruc': _primer_valor,
        'Categoria_letra': _primer_valor,
        'Categoria_monto': 'max',
        'Sede': _unir_unicos,
        'cantidad_cursos': 'sum',
        'Curso Dictado': 'sum',
        'Disenio_examenes': 'sum',
        'Examen_clasif': 'sum',
        'Horas_Total': 'sum',
        'Servicio_actualizacion': 'sum',
        'Total_pago': 'sum',
        'Estado_docente': _primer_valor,
        'Docente_idioma': _unir_unicos,
        'Tipo_documento': _primer_valor,
        'Numero_dni': _primer_valor,
        'Numero_celular': _primer_valor,
        'Domicilio_docente': _primer_valor,
        'Correo_personal': _primer_valor,
        'Nro_Contrato': _primer_valor
    })

    mapeo_docentes = (
        agrupar_df[['docente', 'Docente']]
        .dropna(subset=['Docente'])
        .assign(docente_key=lambda df: df['docente'].apply(TextUtils.normalizar_texto))
        .drop_duplicates(subset=['docente_key'], keep='last')
    )

    cursos_df = datos_csv_procesados.copy()
    cursos_df['docente_key'] = cursos_df['docente'].apply(TextUtils.normalizar_texto)
    cursos_df = cursos_df.merge(
        mapeo_docentes[['docente_key', 'Docente']],
        on='docente_key',
        how='left'
    )


    columnas_docente = [col for col in ['Docente_y', 'Docente', 'Docente_x'] if col in cursos_df.columns]
    if columnas_docente:
        cursos_df['Docente'] = (
            cursos_df[columnas_docente]
            .replace(r'^\s*$', pd.NA, regex=True)
            .bfill(axis=1)
            .iloc[:, 0]
        )
    elif 'docente' in cursos_df.columns:
        cursos_df['Docente'] = cursos_df['docente']
    else:
        cursos_df['Docente'] = pd.NA

    if 'Curso' not in cursos_df.columns:
        cursos_df['Curso'] = (
            cursos_df['idioma'].astype(str) + ' ' +
            cursos_df['nivel'].astype(str) + ' ' +
            cursos_df['ciclo'].astype(str)
        )

    cursos_df['modalidad'] = cursos_df.get('modalidad', '').fillna('').astype(str)
    cursos_df = cursos_df.dropna(subset=['Docente'])
    cursos_df = cursos_df[cursos_df['Docente'].fillna('').astype(str).str.strip() != ''].copy()

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

    cursos_por_docente = cursos_df.groupby('Docente', as_index=False).agg({
        'Curso_Virtual_tmp': _unir_todos,
        'Curso_Presencial_tmp': _unir_todos
    }).rename(columns={
        'Curso_Virtual_tmp': 'Curso_Virtual',
        'Curso_Presencial_tmp': 'Curso_Presencial'
    })

    agrupado_resumen = agrupado_resumen.merge(cursos_por_docente, on='Docente', how='left')
    agrupado_resumen['Curso_Virtual'] = agrupado_resumen['Curso_Virtual'].fillna('')
    agrupado_resumen['Curso_Presencial'] = agrupado_resumen['Curso_Presencial'].fillna('')
    agrupado_resumen['Numero_dni'] = agrupado_resumen['Numero_dni'].apply(lambda valor: formatear_numero(valor, 8))
    agrupado_resumen['Nro_Contrato'] = agrupado_resumen['Nro_Contrato'].apply(lambda valor: formatear_numero(valor, 4))

    agrupado_resumen = agrupado_resumen.sort_values('Docente').reset_index(drop=True)
    agrupado_resumen = agrupado_resumen.drop(columns=['N°'], errors='ignore')
    agrupado_resumen.insert(0, 'N°', range(1, len(agrupado_resumen) + 1))

    columnas_finales = [
        'N°', 'Docente', 'N_Ruc', 'Categoria_letra', 'Categoria_monto', 'Sede',
        'Curso_Virtual', 'Curso_Presencial', 'cantidad_cursos', 'Curso Dictado',
        'Disenio_examenes', 'Examen_clasif', 'Horas_Total', 'Servicio_actualizacion',
        'Total_pago', 'Estado_docente', 'Docente_idioma', 'Tipo_documento', 'Numero_dni', 'Numero_celular',
        'Domicilio_docente', 'Correo_personal', 'Nro_Contrato'
    ]

    for columna in columnas_finales:
        if columna not in agrupado_resumen.columns:
            agrupado_resumen[columna] = '' if columna not in columnas_numericas else 0

    return agrupado_resumen[columnas_finales]
