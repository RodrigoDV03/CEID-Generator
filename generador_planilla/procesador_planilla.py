import pandas as pd
import os
import datetime
from fuzzywuzzy import process, fuzz
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from .utils import *

def traducir_dias(dias_raw: str) -> str:
    dias_dict = {
        'MONDAY': 'Lun', 'TUESDAY': 'Mar', 'WEDNESDAY': 'Mié',
        'THURSDAY': 'Jue', 'FRIDAY': 'Vie', 'SATURDAY': 'Sáb', 'SUNDAY': 'Dom'
    }
    dias = dias_raw.strip('{}').split(',') if isinstance(dias_raw, str) else []
    return ', '.join(dias_dict.get(d.strip().upper(), d.strip()) for d in dias)

def generar_planilla(ruta_cursos: str, ruta_docentes: str, ruta_clasificacion: str, mes_seleccionado: str, numero_carga: int):

    try:
        extension = os.path.splitext(ruta_cursos)[-1].lower()
        if extension == ".csv":
            try:
                datos = pd.read_csv(ruta_cursos, sep=',')
            except Exception:
                datos = pd.read_csv(ruta_cursos)
        elif extension in [".xls", ".xlsx"]:
            datos = pd.read_excel(ruta_cursos)
        else:
            raise ValueError("Formato de cursos no soportado. Usa CSV o Excel.")
        
    
        datos_docentes = pd.read_excel(ruta_docentes, sheet_name="list")

        datos['docente'] = datos['docente'].astype(str).str.strip()
        datos_docentes['Docente'] = datos_docentes['Docente'].astype(str).str.strip()

        datos = datos[~datos['docente'].isin(['', ',', None])]
        datos['detalles_curso'] = datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

        agrupar = (
            datos.groupby('docente')
                 .agg(curso=('detalles_curso', lambda x: ' / '.join(x)),
                      cantidad_cursos=('detalles_curso', 'count'))
                 .reset_index()
        )

        nombres_base = datos_docentes['Docente'].tolist()
        agrupar['Docente'] = [
            process.extractOne(n, nombres_base, scorer=fuzz.token_sort_ratio)[0]
            if process.extractOne(n, nombres_base, scorer=fuzz.token_sort_ratio)[1] >= 85 else None
            for n in agrupar['docente']
        ]

        agrupar = agrupar.merge(
            datos_docentes[['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N°. Ruc', 'Contrato o tercero']],
            on='Docente', how='left'
        )

        agrupar['Curso Dictado'] = agrupar['Categoria (Monto)'] * agrupar['cantidad_cursos'] * 28
        agrupar['Diseño de Examenes'] = agrupar['Categoria (Monto)'] * agrupar['cantidad_cursos'] * 4

        # Examen de clasificación (manejo opcional)
        if os.path.exists(ruta_clasificacion):
            try:
                clasif_df = pd.read_excel(ruta_clasificacion, header=1)
                clasif_df['Docente'] = clasif_df['Docente'].astype(str).str.strip()
                clasif_df['docente_norm'] = clasif_df['Docente'].apply(normalizar_texto)

                agrupar['docente_norm'] = agrupar['Docente'].apply(normalizar_texto)
                agrupar = agrupar.merge(clasif_df[['docente_norm', 'Monto']], on='docente_norm', how='left')
                agrupar['Examen Clasif.'] = agrupar['Monto'].fillna(0)
                agrupar.drop(columns=['docente_norm', 'Monto'], inplace=True)
            except Exception as e:
                print(f"⚠️ Error al leer archivo de clasificación: {e}")
                agrupar['Examen Clasif.'] = 0
        else:
            print("⚠️ No se encontró el archivo de clasificación. Se asumirá monto 0 para todos.")
            agrupar['Examen Clasif.'] = 0

        TABLA = pd.DataFrame({
            'N°': range(1, len(agrupar) + 1),
            'Docente': agrupar['Docente'],
            'Sede': agrupar['Sede'],
            'Categoria (Letra)': agrupar['Categoria (Letra)'],
            'Categoria (Monto)': agrupar['Categoria (Monto)'],
            'N°. Ruc': agrupar['N°. Ruc'],
            'Curso': agrupar['curso'],
            'Curso Dictado': agrupar['Curso Dictado'],
            'Extra Curso': 0,
            'Cantidad Cursos': agrupar['cantidad_cursos'],
            'Diseño de Examenes': agrupar['Diseño de Examenes'],
            'Examen Clasif.': agrupar['Examen Clasif.'],
            'Sub Total Pago S/.': 0,
            'Contrato o Tercero': agrupar['Contrato o tercero']
        })
        TABLA['Sub Total Pago S/.'] = (
            TABLA['Curso Dictado'] + TABLA['Extra Curso'] +
            TABLA['Diseño de Examenes'] + TABLA['Examen Clasif.']
        )

        año_actual = datetime.datetime.now().year
        carpeta_salida = f"{mes_seleccionado} {año_actual}"
        os.makedirs(carpeta_salida, exist_ok=True)
        nombre_salida = f"Planilla {mes_seleccionado} {año_actual}.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)

        estado_planilla = "Primera planilla" if numero_carga == 1 else "Segunda planilla"
        nombre_hoja_carga = f"{numero_carga} carga académica"

        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            TABLA.to_excel(writer, sheet_name=f"Planilla {mes_seleccionado}", index=False)
            columnas_extra = ['Nro. Documento', 'Celular', 'Dirección', 'Correo personal']
            hoja_generador = (
                TABLA
                .merge(datos_docentes[['Docente'] + columnas_extra], on='Docente', how='left')
                .rename(columns={
                    'Cantidad Cursos': 'Cantidad_cursos',
                    'Examen Clasif.': 'Examen_clasif',
                    'Categoria (Letra)': 'Categoria_letra',
                    'Categoria (Monto)': 'Categoria_monto',
                    'Sub Total Pago S/.': 'Subtotal_pago',
                    'Contrato o Tercero': 'Contrato_o_tercero',
                    'N°. Ruc': 'N_Ruc',
                    'Nro. Documento': 'Numero_dni',
                    'Celular': 'Numero_celular',
                    'Dirección': 'Domicilio_docente',
                    'Correo personal': 'Correo_personal'
                })
            )
            hoja_generador['Numero_dni'] = hoja_generador['Numero_dni'].apply(lambda x: str(int(float(x))).zfill(8) if pd.notna(x) else '')
            hoja_generador.to_excel(writer, sheet_name="Planilla_Generador", index=False)

            df_carga = pd.DataFrame({
                'N°': range(1, len(datos) + 1),
                'Dias': datos['dias'].apply(traducir_dias),
                'H. Inicio': datos['horainicio'].astype(str).str[:5],
                'H. Fin': datos['horafin'].astype(str).str[:5],
                'Idioma': datos['idioma'],
                'Nivel': datos['nivel'].str.replace('Ã¡', 'á', regex=False),
                'Ciclo': datos['ciclo'],
                'Curso': datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1),
                'Sede': datos['sede'],
                'Sec.': '',
                'Matr.': datos['matriculados'],
                'Docente': datos['docente'],
                'Modalidad': datos['modalidad'],
                'Estado Planilla': estado_planilla
            })
            df_carga = df_carga.sort_values(by='Docente', ascending=True).reset_index(drop=True)
            df_carga.to_excel(writer, sheet_name=nombre_hoja_carga, index=False)

        wb = load_workbook(ruta_salida)
        titulo_fusionado = (
            "CENTRO DE IDIOMAS - FLCH - UNMSM\n"
            f"{numero_carga} CARGA ACADÉMICA - PERIODO {mes_seleccionado.upper()} {año_actual}\n"
            "MODALIDAD: VIRTUAL Y PRESENCIAL"
        )

        hojas_con_titulo = [f"Planilla {mes_seleccionado}", nombre_hoja_carga]
        for hoja in hojas_con_titulo:
            ws = wb[hoja]
            ws.insert_rows(1)
            max_col = ws.max_column
            rango = f"A1:{get_column_letter(max_col)}1"
            ws.merge_cells(rango)
            ws["A1"].value = titulo_fusionado
            ws["A1"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws["A1"].font = Font(bold=True)

        wb.save(ruta_salida)
        return f"✅ Archivo generado correctamente: '{ruta_salida}'"

    except Exception as e:
        return f"❌ Error: {e}"