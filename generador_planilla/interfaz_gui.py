import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from .procesador_planilla import *

def iniciar_interfaz_planilla(callback_volver=None):
    archivo_cursos_path = ""
    archivo_docentes_path = ""
    archivo_clasif_path = ""

    ventana = tk.Tk()
    ventana.title("Generador de Planilla - CEID")
    ventana.geometry("580x650")
    ventana.configure(bg="#f4f6fa")
    ventana.resizable(False, False)

    PRIMARY = "#2d415a"
    ACCENT = "#4a90e2"
    BG = "#f4f6fa"
    GRAY = "#a0a8b8"
    FONT_TITLE = ("Segoe UI", 18, "bold")
    FONT_LABEL = ("Segoe UI", 11)
    FONT_TEXT = ("Segoe UI", 10)
    FONT_BUTTON = ("Segoe UI", 10, "bold")
    FONT_FOOTER = ("Segoe UI", 9)

    # Título
    tk.Label(ventana, text="Generador de Planilla - CEID", font=FONT_TITLE, bg=BG, fg=PRIMARY).pack(pady=(20, 10))
    ttk.Separator(ventana, orient="horizontal").pack(fill="x", padx=30)

    # Mes
    frame_mes = tk.Frame(ventana, bg=BG)
    frame_mes.pack(pady=(20, 5), fill="x", padx=30)
    tk.Label(frame_mes, text="Selecciona el mes:", font=FONT_LABEL, bg=BG, fg=PRIMARY).pack(side="left")
    mes_var = tk.StringVar(value="Enero")
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    tk.OptionMenu(frame_mes, mes_var, *meses).pack(side="left", padx=(10, 0))

    # Número de Carga
    frame_carga = tk.Frame(ventana, bg=BG)
    frame_carga.pack(pady=(5, 10), fill="x", padx=30)
    tk.Label(frame_carga, text="Número de planilla:", font=FONT_LABEL, bg=BG, fg=PRIMARY).pack(side="left")
    carga_var = tk.IntVar(value=1)
    opciones_carga = [("1 (Primera carga)", 1), ("2 (Segunda carga)", 2)]
    for texto, valor in opciones_carga:
        tk.Radiobutton(frame_carga, text=texto, variable=carga_var, value=valor, bg=BG, font=FONT_TEXT, fg=PRIMARY, selectcolor=BG).pack(side="left", padx=5)

    # Cursos
    frame_cursos = tk.LabelFrame(ventana, text="Archivo crudo (.csv)", font=FONT_TEXT, bg=BG, fg=PRIMARY, padx=10, pady=10)
    frame_cursos.pack(fill="x", padx=30, pady=(10, 5))
    label_cursos = tk.Label(frame_cursos, text="📂 No seleccionado", fg="gray", bg=BG, font=FONT_TEXT)
    label_cursos.pack(side="left", padx=(0, 10))

    def seleccionar_cursos():
        nonlocal archivo_cursos_path
        archivo = filedialog.askopenfilename(title="Selecciona archivo de cursos", filetypes=[("Excel o CSV", "*.xlsx *.xls *.csv")])
        if archivo:
            archivo_cursos_path = archivo
            label_cursos.config(text=f"📁 {os.path.basename(archivo)}", fg=PRIMARY)

    tk.Button(frame_cursos, text="Seleccionar...", font=FONT_BUTTON, bg=ACCENT, fg="white", command=seleccionar_cursos).pack(side="right")

    # Docentes
    frame_docentes = tk.LabelFrame(ventana, text="Lista de docentes (.xlsx)", font=FONT_TEXT, bg=BG, fg=PRIMARY, padx=10, pady=10)
    frame_docentes.pack(fill="x", padx=30, pady=(10, 5))
    label_docentes = tk.Label(frame_docentes, text="📂 No seleccionado", fg="gray", bg=BG, font=FONT_TEXT)
    label_docentes.pack(side="left", padx=(0, 10))

    def seleccionar_docentes():
        nonlocal archivo_docentes_path
        archivo = filedialog.askopenfilename(title="Selecciona archivo de docentes", filetypes=[("Excel", "*.xlsx *.xls")])
        if archivo:
            archivo_docentes_path = archivo
            label_docentes.config(text=f"📁 {os.path.basename(archivo)}", fg=PRIMARY)

    tk.Button(frame_docentes, text="Seleccionar...", font=FONT_BUTTON, bg=ACCENT, fg="white", command=seleccionar_docentes).pack(side="right")

    # Clasificación
    frame_clasif = tk.LabelFrame(ventana, text="Archivo de Examen de Clasificación (.xlsx)", font=FONT_TEXT, bg=BG, fg=PRIMARY, padx=10, pady=10)
    frame_clasif.pack(fill="x", padx=30, pady=(10, 5))
    label_clasif = tk.Label(frame_clasif, text="📂 No seleccionado", fg="gray", bg=BG, font=FONT_TEXT)
    label_clasif.pack(side="left", padx=(0, 10))

    def seleccionar_clasif():
        nonlocal archivo_clasif_path
        archivo = filedialog.askopenfilename(title="Selecciona excel de examen de clasificación", filetypes=[("Excel", "*.xlsx *.xls")])
        if archivo:
            archivo_clasif_path = archivo
            label_clasif.config(text=f"📁 {os.path.basename(archivo)}", fg=PRIMARY)

    tk.Button(frame_clasif, text="Seleccionar...", font=FONT_BUTTON, bg=ACCENT, fg="white", command=seleccionar_clasif).pack(side="right")

    # Procesar
    def procesar():
        if not archivo_cursos_path or not archivo_docentes_path or not archivo_clasif_path:
            messagebox.showerror("Error", "⚠️ Debes seleccionar ambos archivos.")
            return
        resultado = generar_planilla(
            archivo_cursos_path,
            archivo_docentes_path,
            archivo_clasif_path,
            mes_var.get(),
            carga_var.get()
        )
        if resultado.startswith("Error"):
            messagebox.showerror("Error", resultado)
        else:
            messagebox.showinfo("Éxito", resultado)

    tk.Button(ventana, text="🚀 Procesar y generar archivo", font=("Segoe UI", 12, "bold"), bg=PRIMARY, fg="white", command=procesar).pack(pady=30)

    def volver():
        ventana.destroy()
        if callback_volver:
            callback_volver()

    tk.Button(
        ventana, text="⬅ Volver al Menú Principal",
        command=volver,
        font=("Segoe UI", 10),
        bg="#cccccc", fg="#222",
        relief="flat", padx=8, pady=4
    ).pack(pady=(10, 15), side="bottom")



    tk.Label(ventana, text="CEID Generator - v1.0", font=FONT_FOOTER, bg=BG, fg=GRAY).pack(side="bottom", pady=(0, 12))

    ventana.mainloop()
