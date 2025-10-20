import pandas as pd
import os
import datetime
from .functions import *
from .excel_styles import *
from repositories import excel_repo, get_cache

def generar_planilla(data_path, excel_docentes, excel_exa_clasif, month, numero_carga, primera_planilla: str = None):
    año_actual = datetime.datetime.now().year
    
    try:
        # Cache centralizado para optimizar rendimiento
        cache = get_cache()
        cache.clear("planilla_generation")
        
        # Cargar datos base
        print("📂 Cargando datos de cursos...")
        datos = cargar_archivo(data_path)
        
        print("👥 Cargando datos de docentes...")
        if not excel_repo.exists(excel_docentes):
            raise FileNotFoundError(f"Archivo de docentes no encontrado: {excel_docentes}")
        datos_docentes = excel_repo.read_sheet(excel_docentes, sheet_name="list")

        # Procesamiento inicial
        datos = limpiar_docentes(datos, 'docente')
        datos_docentes = limpiar_docentes(datos_docentes, 'Docente')
        datos = datos[~datos['docente'].isin(['', ',', None])]
        
        # Copias para diferentes propósitos
        datos_csv_original = datos.copy()
        datos_csv_original_procesados = aplicar_transformaciones_base(datos_csv_original)
        datos = aplicar_transformaciones_base(datos)
        datos['detalles_curso'] = datos['Curso']

        # Manejo de segunda carga: filtrar combinaciones ya procesadas
        if numero_carga == 2 and primera_planilla and excel_repo.exists(primera_planilla):
            datos = _procesar_planilla_anterior(datos, primera_planilla)

        # Cálculos principales con cache inteligente
        cache_key_agrupacion = f"agrupacion_{month}_{numero_carga}_{hash(str(datos.values.tolist()))}"
        agrupar = cache.get(cache_key_agrupacion)
        if agrupar is None:
            print("🔄 Calculando agrupación (primera vez)...")
            agrupar = agrupar_y_calcular(datos, datos_docentes, 'detalles_curso')
            cache.set(cache_key_agrupacion, agrupar, ttl=3600)
        else:
            print("⚡ Usando agrupación desde cache...")
        
        # Agregar clasificación
        print("📝 Procesando examen de clasificación...")
        agrupar = agregar_examen_clasificacion_mejorada(agrupar, excel_exa_clasif, normalizar_texto)

        # Construir tabla principal
        TABLA = construir_tabla_planilla(agrupar)

        # Configuración de salida
        estado_planilla = "Primera planilla" if numero_carga == 1 else "Segunda planilla"
        numero_carga_letra = "Primera" if numero_carga == 1 else "Segunda"
        nombre_hoja_carga = f"{numero_carga_letra} carga académica"
        
        carpeta_salida = f"{month} {año_actual}"
        os.makedirs(carpeta_salida, exist_ok=True)
        nombre_salida = f"{estado_planilla} {month} {año_actual}.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)

        # Generar todas las hojas del Excel
        print("💾 Generando archivo Excel...")
        hojas_datos = _construir_hojas_excel(
            TABLA, datos_csv_original_procesados, datos_docentes, datos,
            excel_exa_clasif, numero_carga_letra, month, numero_carga,
            estado_planilla, nombre_hoja_carga
        )

        # Escribir archivo Excel
        if not excel_repo.write_excel(hojas_datos, ruta_salida):
            raise Exception(f"Error al escribir archivo Excel: {ruta_salida}")
        
        # Aplicar formato y estilos
        _aplicar_formato_excel(ruta_salida, numero_carga_letra, month, año_actual, nombre_hoja_carga, numero_carga)
        
        print(f"✅ {nombre_salida} generado correctamente.")
        return f"{nombre_salida} generado correctamente."

    except Exception as e:
        error_msg = f"❌ Error: {e}"
        print(error_msg)
        return error_msg
    finally:
        if 'cache' in locals():
            cache.clear("planilla_generation")


def _procesar_planilla_anterior(datos, primera_planilla):
    """Procesa la planilla anterior para filtrar combinaciones ya procesadas."""
    try:
        print("📋 Leyendo planilla anterior...")
        
        # Detectar automáticamente la fila de encabezados correcta
        planilla_anterior_df = None
        headers_to_try = [1, 2, 6, 0]
        
        for header_row in headers_to_try:
            try:
                test_df = excel_repo.read_sheet(primera_planilla, sheet_name="Primera carga académica", header=header_row)
                if not test_df.empty:
                    columnas = [str(col) for col in test_df.columns if str(col) != 'nan' and 'Unnamed' not in str(col)]
                    
                    # Verificar si es la fila correcta
                    keywords_found = sum(1 for col in columnas if any(keyword in col.lower() for keyword in ['docente', 'profesor', 'instructor', 'curso', 'materia', 'asignatura']))
                    
                    if keywords_found >= 1 or len(columnas) >= 5:
                        planilla_anterior_df = test_df
                        print(f"✅ Usando header={header_row} (encontradas {keywords_found} palabras clave)")
                        break
            except:
                continue
        
        if planilla_anterior_df is None:
            print("⚠️ No se pudo determinar la estructura de la planilla anterior")
            return datos
        
        # Detectar columnas de docente y curso
        col_docente, col_curso = _detectar_columnas_planilla(planilla_anterior_df)
        
        if col_docente and col_curso:
            print(f"✅ Usando columnas: Docente='{col_docente}' y Curso='{col_curso}'")
            
            # Filtrar combinaciones válidas
            planilla_filtrada = planilla_anterior_df.dropna(subset=[col_docente, col_curso])
            combinacion_anterior = set(zip(planilla_filtrada[col_docente], planilla_filtrada[col_curso]))
            combinacion_anterior = {(d, c) for d, c in combinacion_anterior if str(d) != 'nan' and str(c) != 'nan'}
            
            if combinacion_anterior:
                datos = limpiar_docentes(datos, 'docente')
                datos = filtrar_combinaciones_optimizado(datos, combinacion_anterior)
                print(f"🔄 Filtrados {len(combinacion_anterior)} combinaciones de la planilla anterior")
            else:
                print("⚠️ No se encontraron combinaciones válidas en la planilla anterior")
        else:
            print("⚠️ No se encontraron columnas de docente/curso, continuando sin filtrar")
    
    except Exception as e:
        print(f"⚠️ Error al procesar planilla anterior: {e}")
        print("⚠️ Continuando sin filtrar combinaciones anteriores")
    
    return datos


def _detectar_columnas_planilla(df):
    columnas_disponibles = list(df.columns)
    col_docente = None
    col_curso = None
    
    # Buscar por palabras clave primero
    for col in columnas_disponibles:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        
        if any(keyword in col_lower for keyword in ['docente', 'profesor', 'instructor']) and col_docente is None:
            col_docente = col
        elif any(keyword in col_lower for keyword in ['curso', 'materia', 'asignatura', 'inglés', 'francés', 'alemán', 'idioma']) and col_curso is None:
            col_curso = col
    
    # Si no se encontró docente por palabras clave, buscar por formato de nombre
    if col_docente is None:
        dias_semana = ['lun', 'mar', 'mie', 'jue', 'vie', 'sab', 'dom', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
        for col in columnas_disponibles:
            col_str = str(col).strip()
            if (',' in col_str and 
                any(char.isalpha() for char in col_str) and 
                len(col_str) > 10 and
                not any(dia in col_str.lower() for dia in dias_semana)):
                col_docente = col
                break
    
    return col_docente, col_curso


def _construir_hojas_excel(TABLA, datos_csv_original_procesados, datos_docentes, datos, excel_exa_clasif, numero_carga_letra, month, numero_carga,
                            estado_planilla, nombre_hoja_carga):
    """Construye todas las hojas del archivo Excel."""
    hojas_datos = {}
    
    # Hoja principal
    hojas_datos[f"{numero_carga_letra} Planilla {month}"] = TABLA
    
    # Hoja de clasificación
    if excel_repo.exists(excel_exa_clasif):
        try:
            print("📊 Leyendo archivo de clasificación...")
            clasif_df = excel_repo.read_sheet(excel_exa_clasif, header=1)
            if isinstance(clasif_df, pd.DataFrame) and not clasif_df.empty:
                hojas_datos["Examen de clasificación"] = clasif_df
            else:
                print("⚠️ Archivo de clasificación vacío o inválido")
        except Exception as e:
            print(f"⚠️ Error al leer archivo de clasificación: {e}")
    
    # Hoja generador
    agrupar_gen = agrupar_y_calcular(datos_csv_original_procesados, datos_docentes, 'Curso')
    agrupar_gen = agregar_examen_clasificacion_mejorada(agrupar_gen, excel_exa_clasif, normalizar_texto)
    TABLA_GENERADOR = construir_tabla_planilla(agrupar_gen)
    
    # Agregar columnas extra al generador
    columnas_extra = ['Idioma', 'Nro. Documento', 'Celular', 'Dirección', 'Correo personal', 'N° Contrato']
    TABLA_GENERADOR = TABLA_GENERADOR.merge(
        datos_docentes[['Docente'] + columnas_extra], 
        on='Docente', 
        how='left'
    ).rename(columns={
        'Categoria (Letra)': 'Categoria_letra',
        'Categoria (Monto)': 'Categoria_monto',
        'Diseño de Examenes': 'Disenio_examenes',
        'Examen Clasif.': 'Examen_clasif',
        'Total Pago S/.': 'Total_pago',
        'Estado': 'Estado_docente',
        'Idioma': 'Docente_idioma',
        'N°. Ruc': 'N_Ruc',
        'Nro. Documento': 'Numero_dni',
        'Celular': 'Numero_celular',
        'Dirección': 'Domicilio_docente',
        'Correo personal': 'Correo_personal',
        'N° Contrato': 'Nro_Contrato'
    })
    hojas_datos["Planilla_Generador"] = TABLA_GENERADOR
    
    # Carga académica
    df_carga = construir_tabla_carga_academica(datos, estado_planilla)
    hojas_datos[nombre_hoja_carga] = df_carga

    # Hojas adicionales para segunda carga
    if numero_carga == 2:
        df_carga_consol = construir_tabla_carga_academica(datos_csv_original_procesados, 'Consolidado')
        hojas_datos["Carga académica consolidada"] = df_carga_consol

        agrupar_consol = agrupar_y_calcular(datos_csv_original_procesados, datos_docentes, 'Curso')
        agrupar_consol = agregar_examen_clasificacion_mejorada(agrupar_consol, excel_exa_clasif, normalizar_texto)
        TABLA_CONSOLIDADA = construir_tabla_planilla(agrupar_consol)
        hojas_datos["Planilla consolidada"] = TABLA_CONSOLIDADA

    return hojas_datos


def _aplicar_formato_excel(ruta_salida, numero_carga_letra, month, año_actual, nombre_hoja_carga, numero_carga):
    """Aplica formato y estilos al archivo Excel."""
    from openpyxl import load_workbook
    
    wb = load_workbook(ruta_salida)

    # Definir títulos para cada hoja
    titulo_hojas = [
        ("Examen de clasificación", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nEXAMEN DE CLASIFICACIÓN - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),
        (f"{numero_carga_letra} Planilla {month}", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\n{numero_carga_letra.upper()} PLANILLA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),
        (nombre_hoja_carga, 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\n{numero_carga_letra.upper()} CARGA ACADÉMICA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),
    ]

    if numero_carga == 2:
        titulo_hojas.extend([
            ("Carga académica consolidada", 
                f"CENTRO DE IDIOMAS - FLCH - UNMSM\nCARGA ACADÉMICA CONSOLIDADA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),
            ("Planilla consolidada", 
                f"CENTRO DE IDIOMAS - FLCH - UNMSM\nPLANILLA CONSOLIDADA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL")
        ])

    # Aplicar formato
    procesar_formato_multiple_hojas(wb, titulo_hojas, numero_carga_letra, month)

    # Ordenar hojas
    hojas_ordenadas = [
        "Examen de clasificación",
        nombre_hoja_carga,
        f"{numero_carga_letra} Planilla {month}",
        "Planilla_Generador"
    ]
    if numero_carga == 2:
        hojas_ordenadas += ["Carga académica consolidada", "Planilla consolidada"]
        
    wb = ordenar_hojas_excel(wb, hojas_ordenadas)        
    wb.save(ruta_salida)

def agregar_examen_clasificacion_mejorada(df, ruta_clasificacion, normalizar_texto):
    if not excel_repo.exists(ruta_clasificacion):
        print("⚠️ Archivo de clasificación no encontrado")
        df['Examen Clasif.'] = 0
        return df
    
    try:
        print("📊 Leyendo archivo de clasificación...")
        clasif_df = excel_repo.read_sheet(ruta_clasificacion, header=1)
        
        # Detectar columna de docente
        if 'Docente' not in clasif_df.columns:
            posibles_docentes = [col for col in clasif_df.columns if 'docente' in str(col).lower()]
            if posibles_docentes:
                clasif_df = clasif_df.rename(columns={posibles_docentes[0]: 'Docente'})
            else:
                print("⚠️ No se encontró columna de docente en el archivo de clasificación")
                df['Examen Clasif.'] = 0
                return df
        
        # Detectar columna de monto
        columna_monto = None
        for col in clasif_df.columns:
            if str(col).lower().strip() == 'monto':
                columna_monto = col
                break
        
        if columna_monto is None:
            print("⚠️ No se encontró columna de monto en el archivo de clasificación")
            df['Examen Clasif.'] = 0
            return df
        
        # Normalizar nombres para hacer el merge
        clasif_df['Docente'] = clasif_df['Docente'].astype(str).str.strip()
        clasif_df['docente_norm'] = clasif_df['Docente'].apply(normalizar_texto)
        df['docente_norm'] = df['Docente'].apply(normalizar_texto)
        
        # Realizar merge y limpiar
        merge_result = df.merge(clasif_df[['docente_norm', columna_monto]], on='docente_norm', how='left')
        merge_result['Examen Clasif.'] = merge_result[columna_monto].fillna(0)
        merge_result.drop(columns=['docente_norm', columna_monto], inplace=True, errors='ignore')
        
        print(f"✅ Clasificación procesada: {len(clasif_df)} registros")
        return merge_result
        
    except Exception as e:
        print(f"⚠️ Error al procesar archivo de clasificación: {e}")
        df['Examen Clasif.'] = 0
        return df