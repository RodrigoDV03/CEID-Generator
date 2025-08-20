import customtkinter as ctk
import os
import pandas as pd
import sys
import threading
from tkinter import filedialog, messagebox
from .control_pagos import *
from utils.constants import *

def iniciar_interfaz_control_pagos(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")


    root = ctk.CTk()
    root.title("Control de Pagos | Generador de Archivos CEID")
    root.geometry("800x750")
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))

    hoja_var = ctk.StringVar()
    numero_armada = ctk.StringVar()
    ruta_planilla = None
    ruta_control = None

    # Título principal
    titulo(root, "Control de Pagos | Generador de Archivos CEID")

    # ARCHIVO EXCEL
    frame_excel = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_excel.pack(padx=30, pady=10, fill="x")

    etiqueta(root, "📄 Seleccionar planilla del mes:").pack(in_=frame_excel, anchor="w", padx=10, pady=(10, 2))
    label_excel = ctk.CTkLabel(frame_excel, text="📂 No seleccionado", text_color=GRAY_COLOR)
    label_excel.pack(side="left", padx=10, pady=(0, 10))

    def seleccionar_archivo():
        nonlocal ruta_planilla
        ruta = filedialog.askopenfilename(title="Seleccionar planilla del mes", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color=BLACK_COLOR)
                boton_gen.configure(state="normal")
                ruta_planilla = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    crear_boton_archivo(frame_excel, label_excel, seleccionar_archivo)

    # ARCHIVO DOCENTE CONTRATO
    frame_contrato = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_contrato.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Seleccionar excel de docentes de contrato:").pack(in_=frame_contrato, anchor="w", padx=10, pady=(7, 0))
    label_docente = ctk.CTkLabel(frame_contrato, text="📂 No seleccionado", text_color=GRAY_COLOR)

    def seleccionar_docente():
        nonlocal ruta_control
        ruta_control = filedialog.askopenfilename(title="Seleccionar excel de docentes de contrato", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta_control:
            try:
                pd.ExcelFile(ruta_control)
                label_docente.configure(text=f"📁 {os.path.basename(ruta_control)}", text_color=BLACK_COLOR)
                boton_gen.ruta_control = ruta_control
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    crear_boton_archivo(frame_contrato, label_docente, seleccionar_docente)

    # Número de armada
    frame_armada = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_armada.pack(padx=30, pady=10, fill="x")
    etiqueta(root, "🔢 Número de armada:").pack(in_=frame_armada, anchor="w", padx=10, pady=(10, 2))
    armadas = ["Primera", "Segunda", "Tercera"]
    crear_option_menu(frame_armada, numero_armada, armadas)

    # BOTÓN GENERAR
    def generar():
        if not ruta_planilla or not hoja_var.get() or not ruta_control or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return
        def tarea():
            try:
                actualizar_control_pagos(ruta_planilla, ruta_control, numero_armada.get())
                messagebox.showinfo("Éxito", f"Control de pagos actualizado.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        threading.Thread(target=tarea).start()

    boton_gen = boton_generador(root, "Actualizar control de pagos", generar)

    # CONSOLA EMBEBIDA
    consola_frame = ctk.CTkFrame(root, height=500, fg_color=WHITE_COLOR)
    consola_frame.pack(padx=30, pady=(10, 20), fill="both", expand=False)
    consola_text = ctk.CTkTextbox(consola_frame, height=400, wrap="word")
    consola_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)
    consola_text.configure(state="disabled")

    class TextRedirector:
        def __init__(self, text_widget):
            self.text_widget = text_widget

        def write(self, message):
            if self.text_widget.winfo_exists():
                self.text_widget.configure(state="normal")
                self.text_widget.insert("end", message)
                self.text_widget.see("end")
                self.text_widget.configure(state="disabled")

        def flush(self):
            pass

    sys.stdout = TextRedirector(consola_text)
    sys.stderr = TextRedirector(consola_text)

    # BOTÓN VOLVER
    boton_volver(root, callback_volver).pack(pady=(5, 15))

    # FOOTER
    footer(root, "CEID Generator - CONTROL DE AVANCE DE PAGOS")

    root.mainloop()