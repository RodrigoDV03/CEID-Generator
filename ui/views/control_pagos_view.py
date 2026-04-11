import customtkinter as ctk
import os
import pandas as pd
import sys
import threading
from tkinter import filedialog

from core.control_pagos.control_pagos import actualizar_control_pagos
from ui.components import TextRedirector
from utils.gui_constants import *
from utils import custom_modals as messagebox


def mostrar_control_pagos(app):

    app.limpiar_contenido()
    app.titulo_header.configure(text="Control de Pagos")

    numero_armada = ctk.StringVar()
    ruta_planilla = None
    ruta_control = None

    contenedor = ctk.CTkScrollableFrame(app.main, fg_color="transparent")
    contenedor.pack(fill="both", expand=True, padx=30, pady=20)

    # =====================================================
    # ① PLANILLA
    # =====================================================
    frame_planilla = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_planilla.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_planilla, text="Planilla del mes", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_excel = ctk.CTkLabel(frame_planilla, text="No seleccionado", text_color=TEXT_LIGHT)
    label_excel.pack(anchor="w", padx=20)

    def validar():
        if ruta_planilla and ruta_control and numero_armada.get():
            boton_gen.configure(state="normal")
        else:
            boton_gen.configure(state="disabled")

    def seleccionar_archivo():
        nonlocal ruta_planilla
        ruta = filedialog.askopenfilename(
            title="Seleccionar planilla del mes",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta:
            try:
                hojas = pd.ExcelFile(ruta).sheet_names
                _ = "Planilla_Generador" if "Planilla_Generador" in hojas else hojas[0]
                label_excel.configure(text=os.path.basename(ruta), text_color="#4CAF50")
                ruta_planilla = ruta
                validar()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron leer las hojas:\n{e}")

    ctk.CTkButton(
        frame_planilla,
        text="Seleccionar archivo",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_archivo
    ).pack(anchor="e", padx=20, pady=10)

    # =====================================================
    # ② CONTROL
    # =====================================================
    frame_control = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_control.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_control, text="Excel de control de pagos", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    label_docente = ctk.CTkLabel(frame_control, text="No seleccionado", text_color=TEXT_LIGHT)
    label_docente.pack(anchor="w", padx=20)

    def seleccionar_docente():
        nonlocal ruta_control
        ruta = filedialog.askopenfilename(
            title="Seleccionar excel de docentes de contrato",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta:
            try:
                pd.ExcelFile(ruta)
                ruta_control = ruta
                label_docente.configure(text=os.path.basename(ruta_control), text_color="#4CAF50")
                validar()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    ctk.CTkButton(
        frame_control,
        text="Seleccionar archivo",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_docente
    ).pack(anchor="e", padx=20, pady=10)

    # =====================================================
    # ③ PARÁMETROS
    # =====================================================
    frame_params = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_params.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_params, text="Parámetros", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    ctk.CTkLabel(frame_params, text="Número de armada", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))

    ctk.CTkOptionMenu(
        frame_params,
        variable=numero_armada,
        values=["Primera", "Segunda", "Tercera"],
        command=lambda _: validar(),
    ).pack(anchor="w", padx=20, pady=(0, 10))

    # =====================================================
    # ④ GENERAR
    # =====================================================
    def generar():
        if not ruta_planilla or not ruta_control or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return

        def tarea():
            try:
                actualizar_control_pagos(ruta_planilla, ruta_control, numero_armada.get())
                messagebox.showinfo("Éxito", f"Control de pagos actualizado.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        threading.Thread(target=tarea).start()

    boton_gen = ctk.CTkButton(
        contenedor,
        text="Actualizar Control de Pagos",
        height=50,
        fg_color=ACCENT_COLOR,
        hover_color=PRIMARY_COLOR,
        state="disabled",
        command=generar
    )
    boton_gen.pack(pady=20)

    # =====================================================
    # ⑤ CONSOLA
    # =====================================================
    consola_frame = ctk.CTkFrame(contenedor, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(fill="both", expand=True, pady=10)

    consola_text = ctk.CTkTextbox(consola_frame, height=120, wrap="word", fg_color=CONSOLE_BG, text_color=WHITE_COLOR)
    consola_text.pack(padx=10, pady=10, fill="both", expand=True)
    consola_text.configure(state="disabled")

    sys.stdout = TextRedirector(consola_text)
    sys.stderr = TextRedirector(consola_text)