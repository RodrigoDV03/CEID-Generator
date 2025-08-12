import openpyxl
import pandas as pd
import os
import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side
from utils.functions import *

def generar_planilla(ruta_cursos: str, ruta_docentes: str, ruta_clasificacion: str, mes_seleccionado: str, numero_carga: int, ruta_planilla_anterior: str = None):

    año_actual = datetime.datetime.now().year
    
    try:
        datos = cargar_archivo(ruta_cursos)
        datos_docentes = pd.read_excel(ruta_docentes, sheet_name="list")

        datos = limpiar_docentes(datos, 'docente')
        datos_docentes = limpiar_docentes(datos_docentes, 'Docente')

        datos = datos[~datos['docente'].isin(['', ',', None])]
        datos_csv_original = datos.copy()

        if numero_carga == 2 and ruta_planilla_anterior and os.path.exists(ruta_planilla_anterior):
            try:
                df_raw = pd.read_excel(ruta_planilla_anterior, sheet_name="Primera carga académica", header=None)

                fila_header = None
                for idx, fila in df_raw.iterrows():
                    if "Docente" in fila.values and "Curso" in fila.values:
                        fila_header = idx
                        break

                if fila_header is None:
                    raise ValueError("Columna 'Docente' no encontrada en la planilla anterior.")

                planilla_anterior_df = pd.read_excel(
                    ruta_planilla_anterior, sheet_name="Primera carga académica", header=fila_header
                )

                datos['nivel'] = datos.apply(ajustar_nivel, axis=1)
                datos['Curso'] = datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)
                datos = limpiar_docentes(datos, 'docente')

                combinacion_anterior = set(zip(planilla_anterior_df['Docente'], planilla_anterior_df['Curso']))
                datos = datos[~datos.apply(lambda row: (row['docente'], row['Curso']) in combinacion_anterior, axis=1)]

            except Exception as e:
                print(f"⚠️ Error al comparar con la primera planilla: {e}")



        datos['nivel'] = datos.apply(ajustar_nivel, axis=1)
        datos['detalles_curso'] = datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

        agrupar = agrupar_y_calcular(datos, datos_docentes, 'detalles_curso')
        agrupar = agregar_clasificacion(agrupar, ruta_clasificacion, normalizar_texto)

        TABLA = construir_tabla(agrupar)

        estado_planilla = "Primera planilla" if numero_carga == 1 else "Segunda planilla"
        if numero_carga == 1:
            numero_carga_letra = "Primera"
        else:
            numero_carga_letra = "Segunda"
        nombre_hoja_carga = f"{numero_carga_letra} carga académica"

        carpeta_salida = f"{mes_seleccionado} {año_actual}"
        os.makedirs(carpeta_salida, exist_ok=True)
        nombre_salida = f"{estado_planilla} {mes_seleccionado} {año_actual}.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)

        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            TABLA.to_excel(writer, sheet_name=f"{numero_carga_letra} Planilla {mes_seleccionado}", index=False)
            columnas_extra = ['Idioma', 'Nro. Documento', 'Celular', 'Dirección', 'Correo personal', 'N° Contrato']

            if os.path.exists(ruta_clasificacion):
                clasif_df = pd.read_excel(ruta_clasificacion, header=1)
                clasif_df.to_excel(writer, sheet_name="Examen de clasificación", index=False)
            
            # Construir hoja Planilla_Generador
            datos_csv_original['nivel'] = datos_csv_original.apply(ajustar_nivel, axis=1)
            datos_csv_original['Curso'] = datos_csv_original[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

            agrupar_gen = agrupar_y_calcular(datos_csv_original, datos_docentes, 'Curso')
            agrupar_gen = agregar_clasificacion(agrupar_gen, ruta_clasificacion, normalizar_texto)

            TABLA_GENERADOR = construir_tabla(agrupar_gen)
            TABLA_GENERADOR.merge(
                datos_docentes[['Docente'] + columnas_extra], on='Docente', how='left'
            ).rename(columns={
                'Cantidad Cursos': 'Cantidad_cursos',
                'Examen Clasif.': 'Examen_clasif',
                'Categoria (Letra)': 'Categoria_letra',
                'Categoria (Monto)': 'Categoria_monto',
                'Sub Total Pago S/.': 'Subtotal_pago',
                'Estado': 'Estado_docente',
                'Idioma': 'Docente_idioma',
                'N°. Ruc': 'N_Ruc',
                'Nro. Documento': 'Numero_dni',
                'Celular': 'Numero_celular',
                'Dirección': 'Domicilio_docente',
                'Correo personal': 'Correo_personal',
                'N° Contrato': 'Nro_contrato'
            }).to_excel(writer, sheet_name="Planilla_Generador", index=False)


            df_carga = crear_df_carga(datos, estado_planilla)
            df_carga.to_excel(writer, sheet_name=nombre_hoja_carga, index=False)

            if numero_carga == 2:
                df_carga_consol = crear_df_carga(datos_csv_original, 'Consolidado')
                df_carga_consol.to_excel(writer, sheet_name="Carga académica consolidada", index=False)

                datos_csv_original['nivel'] = datos_csv_original.apply(ajustar_nivel, axis=1)
                datos_csv_original['Curso'] = datos_csv_original[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

                agrupar_consol = agrupar_y_calcular(datos_csv_original, datos_docentes, 'Curso')
                agrupar_consol = agregar_clasificacion(agrupar_consol, ruta_clasificacion, normalizar_texto)
                TABLA_CONSOLIDADA = construir_tabla(agrupar_consol)
                TABLA_CONSOLIDADA.to_excel(writer, sheet_name="Planilla consolidada", index=False)
        
        wb = load_workbook(ruta_salida)
        titulo_fusionado = (
            "CENTRO DE IDIOMAS - FLCH - UNMSM\n"
            f"{numero_carga_letra.upper()} CARGA ACADÉMICA - PERIODO {mes_seleccionado.upper()} {año_actual}\n"
            "MODALIDAD: VIRTUAL Y PRESENCIAL"
        )

        hojas_con_titulo = [
            ("Examen de clasificación", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nEXAMEN DE CLASIFICACIÓN - PERIODO {mes_seleccionado.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),
            
            (f"{numero_carga_letra} Planilla {mes_seleccionado}", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\n{numero_carga_letra.upper()} PLANILLA - PERIODO {mes_seleccionado.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            (nombre_hoja_carga, 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\n{numero_carga_letra.upper()} CARGA ACADÉMICA - PERIODO {mes_seleccionado.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            ("Carga académica consolidada", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nCARGA ACADÉMICA CONSOLIDADA - PERIODO {mes_seleccionado.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL"),

            ("Planilla consolidada", 
            f"CENTRO DE IDIOMAS - FLCH - UNMSM\nPLANILLA CONSOLIDADA - PERIODO {mes_seleccionado.upper()} {año_actual}\nMODALIDAD: VIRTUAL Y PRESENCIAL")
        ]

        for hoja, titulo_fusionado in hojas_con_titulo:
            if hoja in wb.sheetnames:
                ws = wb[hoja]
                ws.insert_rows(1)
                max_col = ws.max_column
                rango = f"A1:{get_column_letter(max_col)}1"
                ws.merge_cells(rango)
                ws["A1"].value = titulo_fusionado
                ws["A1"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                ws["A1"].fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
                ws["A1"].font = Font(bold=True, color="ffffff", size=22)

                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
                    for cell in row:
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center')

                for col in range(1, max_col + 1):
                    celda = ws.cell(row=2, column=col)
                    celda.alignment = Alignment(horizontal='center', vertical='center')
                    celda.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
                    celda.font = Font(bold=True, color="ffffff", size=12)

                if hoja == f"{numero_carga_letra} Planilla {mes_seleccionado}":
                    fila_total = ws.max_row + 1

                    ws.merge_cells(f"A{fila_total}:G{fila_total}")
                    celda_total = ws[f"A{fila_total}"]
                    celda_total.alignment = Alignment(horizontal="center", vertical="center")
                    celda_total.fill = openpyxl.styles.PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")

                    columnas_sumar = ['H', 'I', 'J', 'K', 'L', 'M']
                    for col in columnas_sumar:
                        celda = ws[f"{col}{fila_total}"]
                        celda.value = f"=SUM({col}3:{col}{fila_total-1})"
                        celda.font = Font(bold=True)
                        celda.alignment = Alignment(horizontal="center", vertical="center")
                        celda.border = thin_border

                    columnas_moneda = ['E', 'H', 'K', 'L', 'M']
                    for col in columnas_moneda:
                        for row in range(3, fila_total + 1):
                            celda = ws[f"{col}{row}"]
                            celda.number_format = '"S/ "#,##0.00'
                        

        hojas_ordenadas = [
            "Examen de clasificación",
            nombre_hoja_carga,
            f"{numero_carga_letra} Planilla {mes_seleccionado}",
            "Planilla_Generador"
        ]
        if numero_carga == 2:
            hojas_ordenadas += [
                "Carga académica consolidada",
                "Planilla consolidada"
            ]
        wb = ordenar_hojas_excel(wb, hojas_ordenadas)        
        wb.save(ruta_salida)
        return f"✅ {nombre_salida} generado correctamente."

    except Exception as e:
        return f"❌ Error: {e}"
