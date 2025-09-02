import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import os, sys, threading
import pandas as pd
from utils.gui_constants import *
from .generador_fase_inicial import procesar_planilla_fase_inicial

def iniciar_interfaz_fase_inicial(callback_volver=None):
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Fase Inicial | Generador de Archivos CEID")
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))

    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))
    numero_armada = ctk.StringVar()
    tipo_fase_inicial = ctk.StringVar(value="planilla docente")

    planilla_path = None
    carpeta_destino = None

    # --- TÍTULO PRINCIPAL ---
    titulo(root, "Fase Inicial | Generador de Archivos CEID")

    # --- CARD PRINCIPAL ---
    card = ctk.CTkFrame(root, fg_color=SECTION_COLOR, corner_radius=15)
    card.pack(padx=40, pady=20, fill="both", expand=False)

    # Tipo
    etiqueta(card, "Elaborar fase inicial para:")
    tipo_menu = crear_option_menu(card, tipo_fase_inicial, ["planilla docente", "administrativo"])

    # Archivo Excel
    etiqueta(card, "Seleccionar planilla del mes:")

    def seleccionar_archivo():
        nonlocal planilla_path
        ruta = filedialog.askopenfilename(title="Seleccionar planilla del mes",
                                          filetypes=[("Archivos Excel", "*.xlsx *.xls")])
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

    # Mes y Año
    etiqueta(card, "Mes:")
    crear_option_menu(card, mes_var, meses)

    etiqueta(card, "Año:")
    crear_option_menu(card, año_var, años)

    # Número de armada
    etiqueta(card, "Número de armada:")
    crear_option_menu(card, numero_armada, ["primera", "segunda", "tercera", "sin armada"])

    # Carpeta destino
    etiqueta(card, "Carpeta de destino:")

    def seleccionar_salida():
        nonlocal carpeta_destino
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if carpeta:
            carpeta_destino = carpeta
            label_destino.configure(text=f"📁 {os.path.basename(carpeta)}", text_color=WHITE_COLOR)

    boton_destino, label_destino = crear_boton_archivo(card, "📂 No seleccionado", seleccionar_salida)

    # Botón generar
    def generar():
        if not planilla_path or not hoja_var.get() or not carpeta_destino or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return
        def tarea():
            try:
                procesar_planilla_fase_inicial(planilla_path, hoja_var.get(), carpeta_destino, mes_var.get(), año_var.get(), numero_armada.get(), tipo_fase_inicial.get())
                messagebox.showinfo("Éxito", f"Documentos de fase inicial generados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        threading.Thread(target=tarea).start()

    boton_gen = boton_generador(card, "Generar Documentos", generar)

    # Consola
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

    # Botón volver
    boton_volver(root, callback_volver)

    # Footer
    footer(root)

    root.mainloop()