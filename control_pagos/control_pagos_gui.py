import customtkinter as ctk
import os
import pandas as pd
import sys
import threading
from tkinter import filedialog
from .control_pagos import actualizar_control_pagos
from utils.gui_constants import *
from utils import custom_modals as messagebox

def iniciar_interfaz_control_pagos(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Control de Pagos | Generador de Archivos CEID")
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))

    hoja_var = ctk.StringVar()
    numero_armada = ctk.StringVar()
    ruta_planilla = None
    ruta_control = None

    # --- TÍTULO PRINCIPAL ---
    titulo(root, "Control de Pagos | Generador de Archivos CEID")

    # --- CARD PRINCIPAL ---
    card = ctk.CTkFrame(root, fg_color=SECTION_COLOR, corner_radius=15)
    card.pack(padx=40, pady=20, fill="both", expand=False)

    # --- PLANILLA DEL MES ---
    etiqueta(card, "Seleccionar planilla del mes:")

    def seleccionar_archivo():
        nonlocal ruta_planilla
        ruta = filedialog.askopenfilename(
            title="Seleccionar planilla del mes",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color=WHITE_COLOR)
                boton_gen.configure(state="normal")
                ruta_planilla = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    boton_excel, label_excel = crear_boton_archivo(card, "📂 No seleccionado", seleccionar_archivo)

    # --- ARCHIVO DOCENTES CONTRATO ---
    etiqueta(card, "Seleccionar Excel de docentes de contrato:")

    def seleccionar_docente():
        nonlocal ruta_control
        ruta_control = filedialog.askopenfilename(
            title="Seleccionar excel de docentes de contrato",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta_control:
            try:
                pd.ExcelFile(ruta_control)
                label_docente.configure(text=f"📁 {os.path.basename(ruta_control)}", text_color=WHITE_COLOR)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    boton_docente, label_docente = crear_boton_archivo(card, "📂 No seleccionado", seleccionar_docente)

    # --- NÚMERO DE ARMADA ---
    etiqueta(card, "Número de armada:")
    crear_option_menu(card, numero_armada, ["Primera", "Segunda", "Tercera"])

    # --- BOTÓN GENERAR ---
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

    boton_gen = boton_generador(card, "Actualizar Control de Pagos", generar)

    # --- CONSOLA EMBEBIDA ---
    consola_frame = ctk.CTkFrame(root, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(padx=40, pady=(10, 20), fill="both", expand=True)
    consola_text = ctk.CTkTextbox(consola_frame, height=120, wrap="word", fg_color=CONSOLE_BG, text_color=WHITE_COLOR)
    consola_text.pack(padx=10, pady=10, fill="both", expand=True)
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
        def flush(self): pass

    sys.stdout = TextRedirector(consola_text)
    sys.stderr = TextRedirector(consola_text)

    # --- BOTÓN VOLVER ---
    boton_volver(root, callback_volver)

    # --- FOOTER ---
    footer(root)

    root.mainloop()