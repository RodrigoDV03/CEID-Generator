import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fuzzywuzzy import process, fuzz
import os
import datetime
import unicodedata
import tempfile

archivo_cursos_path = ""
archivo_clasif_path = ""

def normalizar_texto(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize('NFKD', texto)
    return ''.join([c for c in texto if not unicodedata.combining(c)])

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
            'N° Ruc': agrupar['N°. Ruc'],
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

            # Hoja con columnas renombradas
            hoja_generador = TABLA.rename(columns={
                'Cantidad Cursos': 'Cantidad_cursos',
                'Examen Clasif.': 'Examen_clasif',
                'Categoria(Letra)': 'Categoria_letra',
                'Categoria(Monto)': 'Categoria_monto',
                'Sub Total Pago S/.': 'Subtotal_pago',
                'Contrato o Tercero': 'Contrato_o_tercero'
            })
            hoja_generador.to_excel(writer, sheet_name="Planilla_Generador", index=False)



        return f"Archivo generado correctamente: '{nombre_salida}'"
    except Exception as e:
        return f"Error: {e}"


# --- Interfaz gráfica ---
ventana = tk.Tk()
ventana.title("Generador de Planillas y Documentos")
ventana.geometry("520x380")

mes_var = tk.StringVar(value="Enero")
tk.Label(ventana, text="Selecciona el mes para el nombre de archivo:").pack(pady=(15,5))
mes_opciones = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
tk.OptionMenu(ventana, mes_var, *mes_opciones).pack()

# Label y botón para archivo de cursos
label_cursos = tk.Label(ventana, text="Archivo de cursos: No seleccionado", wraplength=450, fg="gray")
label_cursos.pack(pady=(15, 5))

def seleccionar_cursos():
    global archivo_cursos_path
    archivo = filedialog.askopenfilename(title="Selecciona el archivo de cursos",
                                         filetypes=[("Archivos Excel o CSV", "*.xlsx *.xls *.csv")])
    if archivo:
        archivo_cursos_path = archivo
        nombre = os.path.basename(archivo)
        label_cursos.config(text=f"Archivo de cursos: {nombre}", fg="black")

tk.Button(ventana, text="Seleccionar archivo de cursos", command=seleccionar_cursos).pack()

# Label y botón para archivo de clasificación
label_clasif = tk.Label(ventana, text="Archivo de clasificación: No seleccionado", wraplength=450, fg="gray")
label_clasif.pack(pady=(20, 5))

def seleccionar_clasif():
    global archivo_clasif_path
    archivo = filedialog.askopenfilename(title="Selecciona el archivo de clasificación",
                                         filetypes=[("Archivos Excel", "*.xlsx *.xls")])
    if archivo:
        archivo_clasif_path = archivo
        nombre = os.path.basename(archivo)
        label_clasif.config(text=f"Archivo de clasificación: {nombre}", fg="black")

tk.Button(ventana, text="Seleccionar archivo de clasificación", command=seleccionar_clasif).pack()

# Botón para procesar todo
def procesar_desde_gui():
    if not archivo_cursos_path or not archivo_clasif_path:
        messagebox.showerror("Error", "Debes seleccionar ambos archivos antes de procesar.")
        return
    resultado = generar_planilla(archivo_cursos_path, archivo_clasif_path, mes_var.get())
    messagebox.showinfo("Resultado", resultado)

tk.Button(ventana, text="Procesar y generar archivo", command=procesar_desde_gui).pack(pady=25)

ventana.mainloop()
