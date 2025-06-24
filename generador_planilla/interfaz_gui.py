import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from .procesador_planilla import *

def iniciar_interfaz_planilla(callback_volver=None):
    archivo_cursos_path = ""
    archivo_docentes_path = ""
    archivo_clasif_path = ""

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    ventana = ctk.CTk()
    ventana.title("Generador de Planilla - CEID")
    ventana.geometry("650x720")
    ventana.resizable(False, False)

    def titulo(texto):
        return ctk.CTkLabel(ventana, text=texto, font=ctk.CTkFont(size=22, weight="bold"))

    def seccion_titulo(texto):
        return ctk.CTkLabel(ventana, text=texto, font=ctk.CTkFont(size=16, weight="bold"))

    def etiqueta(texto):
        return ctk.CTkLabel(ventana, text=texto, font=ctk.CTkFont(size=13))

    # Título
    titulo("📄 Generador de Planilla - CEID").pack(pady=(25, 10))

    # Sección Mes y Carga
    marco_opciones = ctk.CTkFrame(ventana)
    marco_opciones.pack(pady=10, padx=30, fill="x")

    # Mes
    etiqueta("Selecciona el mes:").pack(in_=marco_opciones, anchor="w", padx=10, pady=(10, 0))
    mes_var = ctk.StringVar(value="Enero")
    ctk.CTkOptionMenu(marco_opciones, variable=mes_var, values=[
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]).pack(padx=10, pady=(0, 10), fill="x")

    # Número de planilla
    etiqueta("Número de carga:").pack(in_=marco_opciones, anchor="w", padx=10)
    carga_var = ctk.IntVar(value=1)
    ctk.CTkRadioButton(marco_opciones, text="1 (Primera carga)", variable=carga_var, value=1).pack(anchor="w", padx=20)
    ctk.CTkRadioButton(marco_opciones, text="2 (Segunda carga)", variable=carga_var, value=2).pack(anchor="w", padx=20, pady=(0, 10))

    # Función para crear selector de archivo
    def crear_selector_archivo(texto_label, extensiones, actualizar_func):
        frame = ctk.CTkFrame(ventana)
        frame.pack(padx=30, pady=10, fill="x")

        seccion_titulo(texto_label).pack(in_=frame, anchor="w", padx=10, pady=(5, 2))
        label_estado = ctk.CTkLabel(frame, text="📂 No seleccionado", text_color="gray")
        label_estado.pack(side="left", padx=10)

        def seleccionar():
            archivo = filedialog.askopenfilename(filetypes=[extensiones])
            if archivo:
                actualizar_func(archivo)
                label_estado.configure(text=f"📁 {os.path.basename(archivo)}", text_color="black")

        ctk.CTkButton(frame, text="Seleccionar archivo", command=seleccionar, width=160).pack(side="right", padx=10)

    # Curso
    def actualizar_cursos(path): nonlocal archivo_cursos_path; archivo_cursos_path = path
    crear_selector_archivo("Archivo de cursos (.csv / .xlsx)", ("Archivos", "*.csv *.xlsx *.xls"), actualizar_cursos)

    # Docentes
    def actualizar_docentes(path): nonlocal archivo_docentes_path; archivo_docentes_path = path
    crear_selector_archivo("Lista de docentes (.xlsx)", ("Archivos Excel", "*.xlsx *.xls"), actualizar_docentes)

    # Clasificación
    def actualizar_clasif(path): nonlocal archivo_clasif_path; archivo_clasif_path = path
    crear_selector_archivo("Docentes para Examen de Clasificación (.xlsx)", ("Archivos Excel", "*.xlsx *.xls"), actualizar_clasif)

    # Botón de procesar
    def procesar():
        if not archivo_cursos_path or not archivo_docentes_path or not archivo_clasif_path:
            messagebox.showerror("Faltan archivos", "⚠️ Debes seleccionar los tres archivos requeridos.")
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

    ctk.CTkButton(
        ventana, text="🚀 Generar Planilla",
        command=procesar, height=40, font=ctk.CTkFont(size=14, weight="bold")
    ).pack(pady=30, padx=60, fill="x")

    def volver():
        ventana.destroy()
        if callback_volver:
            callback_volver()

    ctk.CTkButton(
        ventana, text="⬅ Volver al menú",
        command=volver, fg_color="#cccccc", text_color="#222"
    ).pack(pady=(10, 20))

    # Footer
    ctk.CTkLabel(ventana, text="CEID Generator - v1.0 · Área de Sistemas", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(0, 10))

    ventana.mainloop()
