import sys
import threading
import customtkinter as ctk
import pandas as pd
import os
from tkinter import filedialog, messagebox
from datetime import datetime
from .generador_fase_final import *
from utils.constants import *

def iniciar_interfaz_fase_final(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Fase Final | Generador de Archivos CEID")
    root.geometry("900x800")
    root.resizable(True, True)
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))


    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))
    numero_armada = ctk.StringVar()

    ruta_planilla = None
    excel_control_pagos = None
    carpeta_destino = None

    # Título principal
    titulo(root, "Fase Final | Generador de Archivos CEID")

    # ARCHIVO EXCEL
    frame_excel = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_excel.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Seleccionar Planilla del mes:").pack(in_=frame_excel, anchor="w", padx=10, pady=(7, 0))
    label_excel = ctk.CTkLabel(frame_excel, text="📂 No seleccionado", text_color=GRAY_COLOR)
    label_excel.pack(side="left", padx=10, pady=(0, 7))

    def seleccionar_archivo():
        nonlocal ruta_planilla
        ruta = filedialog.askopenfilename(title="Seleccionar planilla del mes", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_menu.configure(values=hojas)
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color=BLACK_COLOR)
                boton_gen.configure(state="normal")
                ruta_planilla = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    crear_boton_archivo(frame_excel, label_excel, seleccionar_archivo)

    # Hoja de trabajo
    frame_hoja = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_hoja.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Hoja de trabajo:").pack(in_=frame_hoja, anchor="w", padx=10, pady=(7, 0))
    hoja_menu = crear_option_menu(frame_hoja, hoja_var, [])
    hoja_menu.pack(padx=10, pady=(0, 7), fill="x")

    # ARCHIVO DOCENTE CONTRATO
    frame_docente = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_docente.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Seleccionar excel de docentes de contrato:").pack(in_=frame_docente, anchor="w", padx=10, pady=(7, 0))
    label_docente = ctk.CTkLabel(frame_docente, text="📂 No seleccionado", text_color=GRAY_COLOR)

    def seleccionar_docente():
        nonlocal excel_control_pagos
        excel_control_pagos = filedialog.askopenfilename(title="Seleccionar excel de docentes de contrato", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if excel_control_pagos:
            try:
                pd.ExcelFile(excel_control_pagos)
                label_docente.configure(text=f"📁 {os.path.basename(excel_control_pagos)}", text_color=BLACK_COLOR)
                boton_gen.excel_control_pagos = excel_control_pagos
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    crear_boton_archivo(frame_docente, label_docente, seleccionar_docente)

    # MES Y AÑO
    frame_fecha = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_fecha.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Mes:").pack(in_=frame_fecha, anchor="w", padx=10, pady=(7, 0))

    crear_option_menu(frame_fecha, mes_var, meses)
    etiqueta(root, "Año:").pack(in_=frame_fecha, anchor="w", padx=10, pady=(7, 0))

    crear_option_menu(frame_fecha, año_var, años)

    # Número de armada
    frame_armada = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_armada.pack(padx=20, pady=8, fill="x")
    etiqueta(root, "Número de armada:").pack(in_=frame_armada, anchor="w", padx=10, pady=(7, 0))
    crear_option_menu(frame_armada, numero_armada, ["primera", "segunda", "tercera"])

    # CARPETA DE DESTINO
    frame_destino = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_destino.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Seleccionar carpeta de destino:").pack(in_=frame_destino, anchor="w", padx=10, pady=(7, 0))
    label_destino = ctk.CTkLabel(frame_destino, text="📂 No seleccionado", text_color=GRAY_COLOR)
    label_destino.pack(side="left", padx=10, pady=(0, 7))

    def seleccionar_salida():
        nonlocal carpeta_destino
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if ruta:
            carpeta_destino = ruta
            label_destino.configure(text=f"📂 {os.path.basename(ruta)}", text_color=BLACK_COLOR)

    crear_boton_archivo(frame_destino, label_destino, seleccionar_salida)

    # BOTÓN GENERAR
    def iniciar_generacion():
        if not ruta_planilla or not excel_control_pagos or not hoja_var.get() or not carpeta_destino or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return
        def tarea():
            try:
                procesar_planilla_fase_final(ruta_planilla, excel_control_pagos, hoja_var.get(), carpeta_destino, mes_var.get(), año_var.get(), numero_armada.get())
                messagebox.showinfo("Éxito", f"Documentos de fase final generados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        threading.Thread(target=tarea).start()

    boton_gen = boton_generador(root, "Generar Documentos", iniciar_generacion)

    # CONSOLA EMBEBIDA
    consola_frame = ctk.CTkFrame(root, height=150, fg_color=WHITE_COLOR)
    consola_frame.pack(padx=20, pady=(10, 20), fill="both", expand=False)
    consola_text = ctk.CTkTextbox(consola_frame, height=120, wrap="word")
    consola_text.pack(padx=10, pady=(0, 7), fill="both", expand=True)
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
    footer(root, "CEID Generator - FASE FINAL")

    root.mainloop()