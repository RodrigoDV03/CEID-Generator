import sys
import threading
import customtkinter as ctk
import pandas as pd
import os
from tkinter import filedialog
from datetime import datetime

from core.fases.fase_final.generador_fase_final import procesar_planilla_fase_final
from ui.components import TextRedirector
from utils.gui_constants import *
from utils import custom_modals as messagebox


def mostrar_fase_final(app):

    app.limpiar_contenido()
    app.titulo_header.configure(text="Fase Final")

    # =========================
    # VARIABLES
    # =========================
    hoja_var = ctk.StringVar()
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    numero_armada = ctk.StringVar()
    tipo_fase_final = ctk.StringVar(value="planilla docente (con contrato)")

    planilla_path = None
    excel_control_pagos = None
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
    ctk.CTkOptionMenu(
        frame_conf,
        variable=tipo_fase_final,
        values=[
            "planilla docente (con contrato)",
            "planilla docente (sin contrato)",
            "administrativo"
        ]
    ).pack(anchor="w", padx=20, pady=(0, 10))

    # =====================================================
    # ② PLANILLA
    # =====================================================
    frame_excel = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_excel.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_excel, text="Planilla del mes", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_excel = ctk.CTkLabel(frame_excel, text="No seleccionado", text_color=TEXT_LIGHT)
    label_excel.pack(anchor="w", padx=20)

    def seleccionar_planilla():
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
                messagebox.showerror("Error", str(e))

    ctk.CTkButton(
        frame_excel,
        text="Seleccionar archivo",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_planilla
    ).pack(anchor="e", padx=20, pady=10)

    # =====================================================
    # ③ ARCHIVO CONTRATO (CONDICIONAL)
    # =====================================================
    frame_docente = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)

    ctk.CTkLabel(frame_docente, text="Excel docentes con contrato", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_docente = ctk.CTkLabel(frame_docente, text="No seleccionado", text_color=TEXT_LIGHT)
    label_docente.pack(anchor="w", padx=20)

    def seleccionar_docente():
        nonlocal excel_control_pagos
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if ruta:
            try:
                pd.ExcelFile(ruta)
                excel_control_pagos = ruta
                label_docente.configure(text=os.path.basename(ruta), text_color="#4CAF50")
                validar()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    ctk.CTkButton(
        frame_docente,
        text="Seleccionar archivo",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_docente
    ).pack(anchor="e", padx=20, pady=10)

    def actualizar_visibilidad(*_):
        if tipo_fase_final.get() == "planilla docente (con contrato)":
            frame_docente.pack(fill="x", pady=10)
        else:
            frame_docente.pack_forget()

    tipo_fase_final.trace_add("write", actualizar_visibilidad)
    actualizar_visibilidad()

    # =====================================================
    # ④ PARÁMETROS
    # =====================================================
    frame_params = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_params.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_params, text="Parámetros", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    ctk.CTkLabel(frame_params, text="Mes", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_params, variable=mes_var, values=meses)\
        .pack(anchor="w", padx=20, pady=(0, 10))

    ctk.CTkLabel(frame_params, text="Número de armada", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_params, variable=numero_armada,
                      values=["primera", "segunda", "tercera", "sin armada"])\
        .pack(anchor="w", padx=20, pady=(0, 10))

    # =====================================================
    # ⑤ DESTINO
    # =====================================================
    frame_destino = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_destino.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_destino, text="Carpeta destino", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_destino = ctk.CTkLabel(frame_destino, text="No seleccionado", text_color=TEXT_LIGHT)
    label_destino.pack(anchor="w", padx=20)

    def seleccionar_destino():
        nonlocal carpeta_destino
        ruta = filedialog.askdirectory()
        if ruta:
            carpeta_destino = ruta
            label_destino.configure(text=os.path.basename(ruta), text_color="#4CAF50")
            validar()

    ctk.CTkButton(
        frame_destino,
        text="Seleccionar carpeta",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_destino
    ).pack(anchor="e", padx=20, pady=10)

    # =====================================================
    # VALIDACIÓN
    # =====================================================
    def validar():
        if planilla_path and carpeta_destino and numero_armada.get():
            if tipo_fase_final.get() == "planilla docente (con contrato)" and not excel_control_pagos:
                boton_gen.configure(state="disabled")
            else:
                boton_gen.configure(state="normal")
        else:
            boton_gen.configure(state="disabled")

    # =====================================================
    # GENERAR
    # =====================================================
    def generar():
        def tarea():
            try:
                procesar_planilla_fase_final(
                    planilla_path,
                    excel_control_pagos if tipo_fase_final.get() == "planilla docente (con contrato)" else None,
                    hoja_var.get(),
                    carpeta_destino,
                    mes_var.get(),
                    numero_armada.get(),
                    tipo_fase_final.get()
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
        height=150,
        fg_color=CONSOLE_BG,
        text_color=WHITE_COLOR
    )
    consola_text.pack(fill="both", expand=True, padx=10, pady=10)
    consola_text.configure(state="disabled")

    sys.stdout = TextRedirector(consola_text)
    sys.stderr = TextRedirector(consola_text)