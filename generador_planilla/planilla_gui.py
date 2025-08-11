import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from .generador_planilla import *
from utils.constants import *

def iniciar_interfaz_planilla(callback_volver=None):
    archivo_cursos_path = ""
    archivo_docentes_path = ""
    archivo_clasif_path = ""
    archivo_planilla_anterior_path = ""

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Generador de Planilla - CEID")
    root.geometry("750x800")
    root.after(100, lambda: root.state("zoomed"))
    root.configure(fg_color="#f4f5f7")


    def titulo(texto):
        return ctk.CTkLabel(root, text=texto, font=FONT_TITLE, text_color=PRIMARY_COLOR)

    def seccion_titulo(texto):
        return ctk.CTkLabel(root, text=texto, font=FONT_SECTION, text_color=PRIMARY_COLOR)

    def etiqueta(texto):
        return ctk.CTkLabel(root, text=texto, font=FONT_TEXT)

    # --- TÍTULO PRINCIPAL ---
    titulo("📄 Generador de Planilla - CEID").pack(pady=(25, 15))

    # --- SECCIÓN MES Y CARGA ---
    marco_opciones = ctk.CTkFrame(root, fg_color=BG_COLOR, corner_radius=12)
    marco_opciones.pack(pady=10, padx=40, fill="x")

    etiqueta("Selecciona el mes:").pack(in_=marco_opciones, anchor="w", padx=15, pady=(15, 0))
    mes_var = ctk.StringVar(value="Enero")
    ctk.CTkOptionMenu(marco_opciones, variable=mes_var, values=[
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    ).pack(padx=15, pady=(0, 15), fill="x")

    etiqueta("Número de carga:").pack(in_=marco_opciones, anchor="w", padx=15)
    carga_var = ctk.IntVar(value=1)
    ctk.CTkRadioButton(marco_opciones, text="1 (Primera carga)", variable=carga_var, value=1).pack(anchor="w", padx=25)
    ctk.CTkRadioButton(marco_opciones, text="2 (Segunda carga)", variable=carga_var, value=2).pack(anchor="w", padx=25, pady=(0, 15))

    # --- PLANILLA ANTERIOR (solo si es segunda carga) ---
    marco_planilla_anterior = ctk.CTkFrame(root, fg_color=BG_COLOR, corner_radius=12)
    seccion_titulo("📑 Planilla anterior - Solo para segunda carga").pack(in_=marco_planilla_anterior, anchor="w", padx=15, pady=(10, 5))
    label_planilla_estado = ctk.CTkLabel(marco_planilla_anterior, text="📂 No seleccionado", text_color=GRAY_COLOR)
    label_planilla_estado.pack(side="left", padx=15)

    def seleccionar_planilla_anterior():
        archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if archivo:
            nonlocal archivo_planilla_anterior_path
            archivo_planilla_anterior_path = archivo
            label_planilla_estado.configure(text=f"📁 {os.path.basename(archivo)}", text_color="black")

    btn_planilla_anterior = ctk.CTkButton(marco_planilla_anterior, text="Seleccionar archivo", command=seleccionar_planilla_anterior, width=160)
    btn_planilla_anterior.pack(side="right", padx=15)

    def actualizar_visibilidad_planilla_anterior(*args):
        if carga_var.get() == 2:
            marco_planilla_anterior.pack(padx=40, pady=10, fill="x")
        else:
            marco_planilla_anterior.pack_forget()
            nonlocal archivo_planilla_anterior_path
            archivo_planilla_anterior_path = ""
            label_planilla_estado.configure(text="📂 No seleccionado", text_color="gray")

    carga_var.trace_add("write", actualizar_visibilidad_planilla_anterior)
    actualizar_visibilidad_planilla_anterior()

    # --- SELECTORES DE ARCHIVO ---
    def crear_selector_archivo(texto_label, extensiones, actualizar_func):
        frame = ctk.CTkFrame(root, fg_color=BG_COLOR, corner_radius=12)
        frame.pack(padx=40, pady=10, fill="x")

        seccion_titulo(texto_label).pack(in_=frame, anchor="w", padx=15, pady=(10, 5))
        label_estado = ctk.CTkLabel(frame, text="📂 No seleccionado", text_color=GRAY_COLOR)
        label_estado.pack(side="left", padx=15)

        def seleccionar():
            archivo = filedialog.askopenfilename(filetypes=[extensiones])
            if archivo:
                actualizar_func(archivo)
                label_estado.configure(text=f"📁 {os.path.basename(archivo)}", text_color=BLACK_COLOR)

        ctk.CTkButton(frame, text="Seleccionar archivo", command=seleccionar, width=160).pack(side="right", padx=15)

    def actualizar_cursos(path): nonlocal archivo_cursos_path; archivo_cursos_path = path
    def actualizar_docentes(path): nonlocal archivo_docentes_path; archivo_docentes_path = path
    def actualizar_clasif(path): nonlocal archivo_clasif_path; archivo_clasif_path = path

    crear_selector_archivo("📘 Archivo de cursos", ("Archivos", "*.csv"), actualizar_cursos)
    crear_selector_archivo("👨‍🏫 Lista de docentes", ("Archivos Excel", "*.xlsx *.xls"), actualizar_docentes)
    crear_selector_archivo("🧪 Docentes para Examen de Clasificación", ("Archivos Excel", "*.xlsx *.xls"), actualizar_clasif)

    # --- BOTÓN DE PROCESAR ---
    def procesar():
        if not archivo_cursos_path or not archivo_docentes_path or not archivo_clasif_path:
            messagebox.showerror("Faltan archivos", "⚠️ Debes seleccionar los tres archivos requeridos.")
            return
        resultado = generar_planilla(
            archivo_cursos_path,
            archivo_docentes_path,
            archivo_clasif_path,
            mes_var.get(),
            carga_var.get(),
            archivo_planilla_anterior_path if carga_var.get() == 2 else None
        )
        if resultado.startswith("Error"):
            messagebox.showerror("Error", resultado)
        else:
            messagebox.showinfo("Éxito", resultado)

    ctk.CTkButton(
        root, text="🚀 Generar Planilla",
        command=procesar, height=45, font=ctk.CTkFont(size=15, weight="bold"),
        fg_color=ACCENT_COLOR, hover_color=HOVER_COLOR, text_color=WHITE_COLOR
    ).pack(pady=30, padx=80, fill="x")

    # --- BOTÓN VOLVER ---
    def volver():
        root.destroy()
        if callback_volver:
            callback_volver()

    ctk.CTkButton(
        root, text="⬅ Volver al menú", command=volver,
        fg_color=BUTTON_BG_COLOR, hover_color=BUTTON_HOVER_BG_COLOR, text_color=WHITE_COLOR, font=FONT_BUTTON
    ).pack(pady=(5, 15))

    ctk.CTkLabel(root, text="CEID Generator - GENERADOR DE PLANILLA", font=FONT_FOOTER, text_color=TEXT_COLOR).pack(pady=(0, 10))

    root.mainloop()
