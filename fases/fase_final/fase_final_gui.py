import sys
import threading
import customtkinter as ctk
import pandas as pd
import os
from tkinter import filedialog, messagebox
from datetime import datetime
from .generador_fase_final import procesar_planilla_fase_final
from utils.gui_constants import *

def iniciar_interfaz_fase_final(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Fase Final | Generador de Archivos CEID")
    root.resizable(True, True)
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))

    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))
    numero_armada = ctk.StringVar()
    tipo_fase_final = ctk.StringVar(value="planilla docente (con contrato)")

    planilla_path = None
    excel_control_pagos = None
    carpeta_destino = None

    # Título principal
    titulo(root, "Fase Final | Generador de Archivos CEID")

    # --- CARD PRINCIPAL ---
    card = ctk.CTkFrame(root, fg_color=SECTION_COLOR, corner_radius=15)
    card.pack(padx=40, pady=20, fill="both", expand=False)

    # --- NUEVO BLOQUE PARA TIPO DE CONFORMIDAD ---
    etiqueta(card, "Elaborar fase final para:")
    tipo_menu = crear_option_menu(card, tipo_fase_final, ["planilla docente (con contrato)", "planilla docente (sin contrato)", "administrativo"])

    # ARCHIVO EXCEL
    etiqueta(card, "Seleccionar planilla del mes:")

    def seleccionar_archivo():
        nonlocal planilla_path
        ruta = filedialog.askopenfilename(title="Seleccionar planilla del mes", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color=WHITE_COLOR)
                boton_gen.configure(state="normal")
                planilla_path = ruta
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    boton_excel, label_excel = crear_boton_archivo(card, "📂 No seleccionado", seleccionar_archivo)

    # ARCHIVO DOCENTE CONTRATO (SOLO SI DOCENTE)
    frame_docente = ctk.CTkFrame(card, fg_color="transparent")
    
    etiqueta(frame_docente, "Seleccionar excel de docentes de contrato:")

    def seleccionar_docente():
        nonlocal excel_control_pagos
        excel_control_pagos = filedialog.askopenfilename(title="Seleccionar excel de docentes de contrato", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if excel_control_pagos:
            try:
                pd.ExcelFile(excel_control_pagos)
                label_docente.configure(text=f"📁 {os.path.basename(excel_control_pagos)}", text_color=WHITE_COLOR)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    boton_docente, label_docente = crear_boton_archivo(frame_docente, "📂 No seleccionado", seleccionar_docente)

    def actualizar_visibilidad(*args):
        if tipo_fase_final.get() == "planilla docente (con contrato)":
            frame_docente.pack(fill="x", padx=0, pady=0, after=boton_excel.master)
        else:
            frame_docente.pack_forget()

    tipo_fase_final.trace_add("write", actualizar_visibilidad)
    actualizar_visibilidad()  # inicializar estado

    # MES Y AÑO
    etiqueta(card, "Mes:")
    crear_option_menu(card, mes_var, meses)

    etiqueta(card, "Año:")
    crear_option_menu(card, año_var, años)

    # Número de armada
    etiqueta(card, "Número de armada:")
    crear_option_menu(card, numero_armada, ["primera", "segunda", "tercera"])

    # CARPETA DE DESTINO
    etiqueta(card, "Seleccionar carpeta de destino:")

    def seleccionar_salida():
        nonlocal carpeta_destino
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if ruta:
            carpeta_destino = ruta
            label_destino.configure(text=f"📂 {os.path.basename(ruta)}", text_color=WHITE_COLOR)

    boton_destino, label_destino = crear_boton_archivo(card, "📂 No seleccionado", seleccionar_salida)

    # BOTÓN GENERAR
    def iniciar_generacion():
        if not planilla_path or not hoja_var.get() or not carpeta_destino or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return

        if tipo_fase_final.get() == "planilla docente (con contrato)" and not excel_control_pagos:
            messagebox.showerror("Error", "Debe seleccionar el excel de docentes de contrato.")
            return

        def tarea():
            try:
                procesar_planilla_fase_final(
                    planilla_path,
                    excel_control_pagos if tipo_fase_final.get() == "planilla docente (con contrato)" else None,
                    hoja_var.get(),
                    carpeta_destino,
                    mes_var.get(),
                    año_var.get(),
                    numero_armada.get(),
                    tipo_fase_final.get()
                )
                messagebox.showinfo("Éxito", f"Documentos de fase final generados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        threading.Thread(target=tarea).start()

    boton_gen = boton_generador(card, "Generar Documentos", iniciar_generacion)

    # CONSOLA EMBEBIDA
    consola_frame = ctk.CTkFrame(root, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(padx=40, pady=(10, 20), fill="both", expand=False)
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

    # BOTÓN VOLVER
    boton_volver(root, callback_volver)

    # FOOTER
    footer(root)

    root.mainloop()