import sys
import threading
import customtkinter as ctk
from matplotlib.pylab import pad
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

    ruta_excel = None
    archivo_docente = None
    carpeta_destino = None

    # Título principal
    titulo(root, "Fase Final | Generador de Archivos CEID").pack(pady=(20, 8))

    # ARCHIVO EXCEL
    frame_excel = ctk.CTkFrame(root, fg_color=WHITE_COLOR)
    frame_excel.pack(padx=20, pady=8, fill="x")

    etiqueta(root, "Seleccionar Planilla del mes:").pack(in_=frame_excel, anchor="w", padx=10, pady=(7, 0))
    label_excel = ctk.CTkLabel(frame_excel, text="📂 No seleccionado", text_color=GRAY_COLOR)
    label_excel.pack(side="left", padx=10, pady=(0, 7))

    def seleccionar_archivo():
        nonlocal ruta_excel
        ruta = filedialog.askopenfilename(title="Seleccionar planilla del mes", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_menu.configure(values=hojas)
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=f"📁 {os.path.basename(ruta)}", text_color="black")
                boton_generar.configure(state="normal")
                ruta_excel = ruta
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
        nonlocal archivo_docente
        archivo_docente = filedialog.askopenfilename(title="Seleccionar excel de docentes de contrato", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if archivo_docente:
            try:
                pd.ExcelFile(archivo_docente)
                label_docente.configure(text=f"📁 {os.path.basename(archivo_docente)}", text_color=BLACK_COLOR)
                boton_generar.archivo_docente = archivo_docente
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
        if not ruta_excel or not archivo_docente or not hoja_var.get() or not carpeta_destino or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return
        def tarea():
            try:
                procesar_planilla(ruta_excel, archivo_docente, hoja_var.get(), carpeta_destino, mes_var.get(), año_var.get(), numero_armada.get())
                messagebox.showinfo("Éxito", f"Documentos de fase final generados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        threading.Thread(target=tarea).start()

    boton_generar = ctk.CTkButton(
        root, text="🚀 Generar Documentos", command=iniciar_generacion,
        state="disabled", height=40, fg_color=ACCENT_COLOR, hover_color=HOVER_COLOR, text_color=WHITE_COLOR, font=FONT_BUTTON
    )
    boton_generar.pack(pady=25, padx=60, fill="x")

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

    def volver():
        root.destroy()
        if callback_volver:
            callback_volver()

    ctk.CTkButton(
        root, text="⬅ Volver al menú", command=volver,
        fg_color=BUTTON_BG_COLOR, hover_color=BUTTON_HOVER_BG_COLOR, text_color=WHITE_COLOR, font=FONT_BUTTON
    ).pack(pady=(5, 15))

    ctk.CTkLabel(root, text="CEID Generator - FASE FINAL", font=FONT_FOOTER, text_color=TEXT_COLOR).pack(pady=(0, 10))

    root.mainloop()