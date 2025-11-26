import pandas as pd
import os
import datetime
from openpyxl import load_workbook
from .functions import *
from .excel_styles import *

def generar_planilla(data_path, excel_docentes, excel_exa_clasif, month, numero_carga, primera_planilla: str = None):
    año_actual = datetime.datetime.now().year
    
    try:
        limpiar_cache_excel()
        limpiar_cache_procesamiento()
        
        datos = cargar_archivo(data_path)
        datos_docentes = pd.read_excel(excel_docentes, sheet_name="list")

        datos = limpiar_docentes(datos, 'docente')
        datos_docentes = limpiar_docentes(datos_docentes, 'Docente')

        datos = datos[~datos['docente'].isin(['', ',', None])]
        datos_csv_original = datos.copy()

        datos_csv_original_procesados = aplicar_transformaciones_base(datos_csv_original)

        datos = aplicar_transformaciones_base(datos)
        datos['detalles_curso'] = datos['Curso']

        if numero_carga == 2 and primera_planilla and os.path.exists(primera_planilla):
            try:
                planilla_anterior_df = leer_planilla_anterior_con_cache(primera_planilla)
                if planilla_anterior_df.empty:
                    raise ValueError("Error al leer la planilla anterior")
                    
                datos = limpiar_docentes(datos, 'docente')

                combinacion_anterior = set(zip(planilla_anterior_df['Docente'], planilla_anterior_df['Curso']))
                datos = filtrar_combinaciones_optimizado(datos, combinacion_anterior)

            except Exception as e:
                print(f"⚠️ Error al comparar con la primera planilla: {e}")

        agrupar = agrupar_y_calcular_con_cache(datos, datos_docentes, 'detalles_curso')
        agrupar = agregar_examen_clasificacion(agrupar, excel_exa_clasif, normalizar_texto)

        TABLA = construir_tabla_planilla_con_cache(agrupar)

        estado_planilla = "Primera planilla" if numero_carga == 1 else "Segunda planilla"
        if numero_carga == 1:
            numero_carga_letra = "Primera"
        else:
            numero_carga_letra = "Segunda"
        nombre_hoja_carga = f"{numero_carga_letra} carga académica"

        carpeta_salida = f"{month} {año_actual}"
        os.makedirs(carpeta_salida, exist_ok=True)
        nombre_salida = f"{estado_planilla} {month} {año_actual}.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)

        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            TABLA.to_excel(writer, sheet_name=f"{numero_carga_letra} Planilla {month}", index=False)
            columnas_extra = ['Idioma', 'Nro. Documento', 'Celular', 'Dirección', 'Correo personal', 'N° Contrato']

            if os.path.exists(excel_exa_clasif):
                # Usar cache para evitar lecturas múltiples
                clasif_df = cargar_excel_con_cache(excel_exa_clasif, sheet_name=0, header=1)
                clasif_df.to_excel(writer, sheet_name="Examen de clasificación", index=False)
            
            # Construir hoja Planilla_Generador usando datos ya procesados
            agrupar_gen = agrupar_y_calcular_con_cache(datos_csv_original_procesados, datos_docentes, 'Curso')
            agrupar_gen = agregar_examen_clasificacion(agrupar_gen, excel_exa_clasif, normalizar_texto)

            TABLA_GENERADOR = construir_tabla_planilla_con_cache(agrupar_gen)
            TABLA_GENERADOR = TABLA_GENERADOR.merge(datos_docentes[['Docente'] + columnas_extra], on='Docente', how='left').rename(columns={
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
                'N° Contrato': 'Nro_contrato'
            }).to_excel(writer, sheet_name="Planilla_Generador", index=False)


            df_carga = construir_tabla_carga_academica(datos, estado_planilla)
            df_carga.to_excel(writer, sheet_name=nombre_hoja_carga, index=False)

            if numero_carga == 2:
                df_carga_consol = construir_tabla_carga_academica(datos_csv_original_procesados, 'Consolidado')
                df_carga_consol.to_excel(writer, sheet_name="Carga académica consolidada", index=False)

                # Optimización: Reutilizar cálculos ya realizados en lugar de duplicar
                # Esta operación es idéntica a agrupar_gen, el cache la detectará automáticamente
                agrupar_consol = agrupar_y_calcular_con_cache(datos_csv_original_procesados, datos_docentes, 'Curso')
                agrupar_consol = agregar_examen_clasificacion(agrupar_consol, excel_exa_clasif, normalizar_texto)
                TABLA_CONSOLIDADA = construir_tabla_planilla_con_cache(agrupar_consol)
                TABLA_CONSOLIDADA.to_excel(writer, sheet_name="Planilla consolidada", index=False)
        
        wb = load_workbook(ruta_salida)

        titulo_hojas = [
            ("Examen de clasificación", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nEXAMEN DE CLASIFICACIÓN - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),
            
            (f"{numero_carga_letra} Planilla {month}", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\n{numero_carga_letra.upper()} PLANILLA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            (nombre_hoja_carga, 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\n{numero_carga_letra.upper()} CARGA ACADÉMICA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            ("Carga académica consolidada", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nCARGA ACADÉMICA CONSOLIDADA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            ("Planilla consolidada", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nPLANILLA CONSOLIDADA - PERIODO {month.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL")
        ]

        # Procesar formato de todas las hojas de manera optimizada
        procesar_formato_multiple_hojas(wb, titulo_hojas, numero_carga_letra, month)
                        

        hojas_ordenadas = [
            "Examen de clasificación",
            nombre_hoja_carga,
            f"{numero_carga_letra} Planilla {month}",
            "Planilla_Generador"
        ]
        if numero_carga == 2:
            hojas_ordenadas += [
                "Carga académica consolidada",
                "Planilla consolidada"
            ]
        wb = ordenar_hojas_excel(wb, hojas_ordenadas)        
        wb.save(ruta_salida)
        return f"{nombre_salida} generado correctamente."

    except Exception as e:
        return f"❌ Error: {e}"
    finally:
        limpiar_cache_planilla()
        limpiar_cache_procesamiento()
