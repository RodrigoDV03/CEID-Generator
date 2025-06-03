import pandas as pd
import os
import datetime
import tempfile
from fuzzywuzzy import process, fuzz
from .utils import normalizar_texto

def generar_planilla(ruta_cursos, ruta_clasificacion, mes_seleccionado):
    try:
        extension = os.path.splitext(ruta_cursos)[-1].lower()
        if extension == ".csv":
            try:
                datos = pd.read_csv(ruta_cursos, sep=',')
            except Exception:
                datos = pd.read_csv(ruta_cursos)  # fallback a coma

            # Guardar como Excel temporal
            temp_excel = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            datos.to_excel(temp_excel.name, index=False)
            ruta_cursos_excel = temp_excel.name
            datos = pd.read_excel(ruta_cursos_excel)
        elif extension in [".xls", ".xlsx"]:
            datos = pd.read_excel(ruta_cursos)
        else:
            raise ValueError("Formato no soportado. Usa CSV o Excel.")



        archivo_docentes = os.path.join(os.path.dirname(__file__), "docentes.xlsx")
        hoja_docentes = "list"
        datos2 = pd.read_excel(archivo_docentes, sheet_name=hoja_docentes)
        clasif_df = pd.read_excel(ruta_clasificacion, header=1)

        # Limpiar columnas clave
        datos['docente'] = datos['docente'].astype(str).str.strip()
        datos2['Docente'] = datos2['Docente'].astype(str).str.strip()
        clasif_df['Docente'] = clasif_df['Docente'].astype(str).str.strip()

        # Eliminar filas vacías
        datos = datos[~datos['docente'].isin(['', ',', None])]
        datos['detalles_curso'] = datos[['idioma', 'nivel', 'ciclo']].astype(str).agg(' '.join, axis=1)

        # Agrupamos
        agrupar = datos.groupby('docente').agg(
            curso=('detalles_curso', lambda x: ' / '.join(x)),
            cantidad_cursos=('detalles_curso', 'count')
        ).reset_index()

        # Fuzzy matching con nombres base
        nombres_base = datos2['Docente'].tolist()
        docente_equivalente = []
        for nombre in agrupar['docente']:
            match, score = process.extractOne(nombre, nombres_base, scorer=fuzz.token_sort_ratio)
            docente_equivalente.append(match if score >= 85 else None)
        agrupar['Docente'] = docente_equivalente

        # Merge con base de datos de docentes
        datos2_filtrado = datos2[['Docente', 'Sede', 'Categoria (Letra)', 'Categoria (Monto)', 'N°. Ruc', 'Contrato o tercero']]
        agrupar = agrupar.merge(datos2_filtrado, on='Docente', how='left')

        # Cálculo de montos
        agrupar['Curso Dictado'] = agrupar['Categoria (Monto)'] * agrupar['cantidad_cursos'] * 28
        agrupar['Diseño de Examenes'] = agrupar['Categoria (Monto)'] * agrupar['cantidad_cursos'] * 4

        # Normalización de nombres para clasificación
        agrupar['docente_norm'] = agrupar['Docente'].apply(normalizar_texto)
        clasif_df['docente_norm'] = clasif_df['Docente'].apply(normalizar_texto)

        # Merge para Examen Clasif.
        agrupar = agrupar.merge(
            clasif_df[['docente_norm', 'Monto']],
            on='docente_norm',
            how='left'
        )

        agrupar['Examen Clasif.'] = agrupar['Monto'].fillna(0)
        agrupar.drop(columns=['docente_norm', 'Monto'], inplace=True)

        # Tabla final
        TABLA = pd.DataFrame({
            'N°': range(1, len(agrupar) + 1),
            'Docente': agrupar['Docente'],
            'Sede': agrupar['Sede'],
            'Categoria(Letra)': agrupar['Categoria (Letra)'],
            'Categoria(Monto)': agrupar['Categoria (Monto)'],
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

        # Subtotal
        TABLA['Sub Total Pago S/.'] = (
            TABLA['Curso Dictado'].fillna(0) +
            TABLA['Extra Curso'].fillna(0) +
            TABLA['Diseño de Examenes'].fillna(0) +
            TABLA['Examen Clasif.'].fillna(0)
        )

        año_actual = datetime.datetime.now().year
        nombre_salida = f"Planilla_{mes_seleccionado}_{año_actual}.xlsx"

        with pd.ExcelWriter(nombre_salida, engine='openpyxl') as writer:
            # Hoja principal
            TABLA.to_excel(writer, sheet_name=f"Planilla {mes_seleccionado}", index=False)

            # Agregar columnas adicionales desde datos2
            columnas_extra = ['DNI', 'Celular', 'Direccion', 'Correo personal']
            datos2_extras = datos2[['Docente'] + columnas_extra]

            # Merge con datos extendidos
            hoja_generador = TABLA.merge(datos2_extras, on='Docente', how='left')

            # Renombrar columnas
            hoja_generador = hoja_generador.rename(columns={
                'Cantidad Cursos': 'Cantidad_cursos',
                'Examen Clasif.': 'Examen_clasif',
                'Categoria(Letra)': 'Categoria_letra',
                'Categoria(Monto)': 'Categoria_monto',
                'Sub Total Pago S/.': 'Subtotal_pago',
                'Contrato o Tercero': 'Contrato_o_tercero',
                'N°. Ruc': 'N_Ruc',
                'DNI': 'Numero_dni',
                'Celular': 'Numero_celular',
                'Direccion': 'Domicilio_docente',
                'Correo personal': 'Correo_personal'
            })
            hoja_generador['Numero_dni'] = hoja_generador['Numero_dni'].apply(lambda x: str(int(float(x))).zfill(8) if pd.notna(x) else '')
            hoja_generador.to_excel(writer, sheet_name="Planilla_Generador", index=False)



        return f"Archivo generado correctamente: '{nombre_salida}'"
    except Exception as e:
        return f"Error: {e}"