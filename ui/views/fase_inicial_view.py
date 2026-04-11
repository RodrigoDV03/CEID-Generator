import customtkinter as ctk
from tkinter import filedialog
from datetime import datetime
import os, sys, threading
import pandas as pd

from ui.components import TextRedirector
from utils.gui_constants import *
from utils import custom_modals as messagebox
from core.fases.fase_inicial.generador_fase_inicial import procesar_planilla_fase_inicial


def mostrar_fase_inicial(app):

    app.limpiar_contenido()
    app.titulo_header.configure(text="Fase Inicial")

    # =========================
    # VARIABLES
    # =========================
    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    numero_armada = ctk.StringVar()
    tipo_fase_inicial = ctk.StringVar(value="planilla docente")

    planilla_path = None
    carpeta_destino = None

    # =========================
    # CONTENEDOR
    # =========================
    contenedor = ctk.CTkScrollableFrame(app.main, fg_color="transparent")
    contenedor.pack(fill="both", expand=True, padx=30, pady=20)

    # =====================================================
    # ① CONFIGURACIÓN
    # =====================================================
    frame_conf = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_conf.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_conf, text="Configuración", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    ctk.CTkLabel(frame_conf, text="Tipo de procesamiento", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_conf, variable=tipo_fase_inicial,
                      values=["planilla docente", "administrativo"])\
        .pack(anchor="w", padx=20, pady=(0, 10))

    # =====================================================
    # ② ARCHIVO EXCEL
    # =====================================================
    frame_excel = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_excel.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_excel, text="Planilla del mes", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_excel = ctk.CTkLabel(frame_excel, text="No seleccionado", text_color=TEXT_LIGHT)
    label_excel.pack(anchor="w", padx=20)

    def seleccionar_archivo():
        nonlocal planilla_path
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                hoja_var.set("Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0])
                label_excel.configure(text=os.path.basename(ruta), text_color="#4CAF50")
                planilla_path = ruta
                validar()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    ctk.CTkButton(
        frame_excel,
        text="Seleccionar archivo",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_archivo
    ).pack(padx=20, pady=10, anchor="e")

    # =====================================================
    # ③ CONFIG EXTRA
    # =====================================================
    frame_extra = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_extra.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_extra, text="Parámetros", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    # Mes
    ctk.CTkLabel(frame_extra, text="Mes", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20)

    ctk.CTkOptionMenu(frame_extra, variable=mes_var, values=meses)\
        .pack(anchor="w", padx=20, pady=5)

    # Armada
    ctk.CTkLabel(frame_extra, text="Número de armada", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20)

    ctk.CTkOptionMenu(frame_extra, variable=numero_armada,
                      values=["primera", "segunda", "tercera", "sin armada"])\
        .pack(anchor="w", padx=20, pady=5)

    # =====================================================
    # ④ DESTINO
    # =====================================================
    frame_destino = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_destino.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_destino, text="Carpeta destino", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_destino = ctk.CTkLabel(frame_destino, text="No seleccionado", text_color=TEXT_LIGHT)
    label_destino.pack(anchor="w", padx=20)

    def seleccionar_salida():
        nonlocal carpeta_destino
        carpeta = filedialog.askdirectory()
        if carpeta:
            carpeta_destino = carpeta
            label_destino.configure(text=os.path.basename(carpeta), text_color="#4CAF50")
            validar()

    ctk.CTkButton(
        frame_destino,
        text="Seleccionar carpeta",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_salida
    ).pack(anchor="e", padx=20, pady=10)

    # =====================================================
    # VALIDACIÓN
    # =====================================================
    def validar():
        if planilla_path and carpeta_destino and numero_armada.get():
            boton_gen.configure(state="normal")
        else:
            boton_gen.configure(state="disabled")

    # =====================================================
    # GENERAR
    # =====================================================
    def generar():

        if not planilla_path or not carpeta_destino:
            messagebox.showerror("Error", "Faltan datos")
            return

        def tarea():
            try:
                procesar_planilla_fase_inicial(
                    planilla_path,
                    hoja_var.get(),
                    carpeta_destino,
                    mes_var.get(),
                    numero_armada.get(),
                    tipo_fase_inicial.get()
                )
                messagebox.showinfo("Éxito", "Documentos generados correctamente")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        threading.Thread(target=tarea).start()

    boton_gen = ctk.CTkButton(
        contenedor,
        text="Generar documentos",
        height=50,
        fg_color=ACCENT_COLOR,
        hover_color=PRIMARY_COLOR,
        state="disabled",
        command=generar
    )
    boton_gen.pack(pady=20)

    # =====================================================
    # CONSOLA
    # =====================================================
    consola_frame = ctk.CTkFrame(contenedor, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(fill="both", pady=10)

    consola_text = ctk.CTkTextbox(
        consola_frame,
        height=180,
        fg_color=CONSOLE_BG,
        text_color=WHITE_COLOR
    )
    consola_text.pack(fill="both", expand=True, padx=10, pady=10)
    consola_text.configure(state="disabled")

    sys.stdout = TextRedirector(consola_text)
    sys.stderr = TextRedirector(consola_text)