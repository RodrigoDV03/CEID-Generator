import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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


# --- VARIABLES GLOBALES ---
archivo_cursos_path = ""
archivo_clasif_path = ""

# --- VENTANA PRINCIPAL ---
ventana = tk.Tk()
ventana.title("Generador de Planilla - CEID")
ventana.geometry("580x480")
ventana.configure(bg="#f4f6fa")
ventana.resizable(False, False)

# --- ESTILOS Y COLORES ---
PRIMARY = "#2d415a"
ACCENT = "#4a90e2"
DISABLED = "#bbb"
BG = "#f4f6fa"
GRAY = "#a0a8b8"

FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_LABEL = ("Segoe UI", 11)
FONT_TEXT = ("Segoe UI", 10)
FONT_BUTTON = ("Segoe UI", 10, "bold")
FONT_FOOTER = ("Segoe UI", 9)

# --- ENCABEZADO ---
tk.Label(
    ventana,
    text="Generador de Planillas y Documentos",
    font=FONT_TITLE,
    bg=BG,
    fg=PRIMARY
).pack(pady=(20, 10))

ttk.Separator(ventana, orient="horizontal").pack(fill="x", padx=30)

# --- SELECCIÓN DE MES ---
frame_mes = tk.Frame(ventana, bg=BG)
frame_mes.pack(pady=(20, 5), fill="x", padx=30)

tk.Label(
    frame_mes,
    text="Selecciona el mes para el nombre del archivo:",
    font=FONT_LABEL,
    bg=BG,
    fg=PRIMARY
).pack(side="left")

mes_var = tk.StringVar(value="Enero")
meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
tk.OptionMenu(frame_mes, mes_var, *meses).pack(side="left", padx=(10, 0))

# --- SECCIÓN DE ARCHIVO DE CURSOS ---
frame_cursos = tk.LabelFrame(
    ventana,
    text="Archivo de Cursos",
    font=FONT_TEXT,
    bg=BG,
    fg=PRIMARY,
    padx=10, pady=10
)
frame_cursos.pack(fill="x", padx=30, pady=(18, 5))

label_cursos = tk.Label(
    frame_cursos,
    text="📂 No seleccionado",
    fg="gray",
    bg=BG,
    font=FONT_TEXT,
    wraplength=420
)
label_cursos.pack(side="left", padx=(0, 10))

def seleccionar_cursos():
    global archivo_cursos_path
    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo de cursos",
        filetypes=[("Archivos Excel o CSV", "*.xlsx *.xls *.csv")]
    )
    if archivo:
        archivo_cursos_path = archivo
        label_cursos.config(text=f"📁 {os.path.basename(archivo)}", fg=PRIMARY)

tk.Button(
    frame_cursos,
    text="Seleccionar...",
    font=FONT_BUTTON,
    bg=ACCENT,
    fg="white",
    activebackground="#357ABD",
    activeforeground="white",
    relief="flat",
    command=seleccionar_cursos
).pack(side="right")

# --- SECCIÓN DE ARCHIVO DE CLASIFICACIÓN ---
frame_clasif = tk.LabelFrame(
    ventana,
    text="Archivo de Clasificación",
    font=FONT_TEXT,
    bg=BG,
    fg=PRIMARY,
    padx=10, pady=10
)
frame_clasif.pack(fill="x", padx=30, pady=(10, 5))

label_clasif = tk.Label(
    frame_clasif,
    text="📂 No seleccionado",
    fg="gray",
    bg=BG,
    font=FONT_TEXT,
    wraplength=420
)
label_clasif.pack(side="left", padx=(0, 10))

def seleccionar_clasif():
    global archivo_clasif_path
    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo de clasificación",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        archivo_clasif_path = archivo
        label_clasif.config(text=f"📁 {os.path.basename(archivo)}", fg=PRIMARY)

tk.Button(
    frame_clasif,
    text="Seleccionar...",
    font=FONT_BUTTON,
    bg=ACCENT,
    fg="white",
    activebackground="#357ABD",
    activeforeground="white",
    relief="flat",
    command=seleccionar_clasif
).pack(side="right")

# --- BOTÓN DE PROCESAMIENTO ---
def procesar_desde_gui():
    if not archivo_cursos_path or not archivo_clasif_path:
        messagebox.showerror("Error", "⚠️ Debes seleccionar ambos archivos antes de procesar.")
        return
    resultado = generar_planilla(archivo_cursos_path, archivo_clasif_path, mes_var.get())
    if resultado.startswith("Error"):
        messagebox.showerror("Error", resultado)
    else:
        messagebox.showinfo("Éxito", resultado)

tk.Button(
    ventana,
    text="🚀 Procesar y generar archivo",
    font=("Segoe UI", 12, "bold"),
    bg=PRIMARY,
    fg="white",
    activebackground="#466a8f",
    activeforeground="white",
    relief="flat",
    command=procesar_desde_gui,
    padx=12, pady=10
).pack(pady=30)

# --- PIE DE PÁGINA ---
tk.Label(
    ventana,
    text="CEID Generator - v1.0",
    font=FONT_FOOTER,
    bg=BG,
    fg=GRAY
).pack(side="bottom", pady=(0, 12))

ventana.mainloop()