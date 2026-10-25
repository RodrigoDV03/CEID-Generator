import customtkinter as ctk
import os
import pandas as pd
from tkinter import filedialog

from core.control_pagos.control_pagos import actualizar_control_pagos
from ui.console_stream import ConsolaEnVivo
from ui.progress import BarraProgreso
from utils.gui_constants import *
from utils import custom_modals as messagebox


def mostrar_control_pagos(app):

    app.limpiar_contenido()
    app.titulo_header.configure(text="Control de Pagos")

    numero_armada = ctk.StringVar()
    ruta_planilla = None
    ruta_control = None

    root = app.root if hasattr(app, "root") else app.winfo_toplevel()

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
                pd.ExcelFile(ruta)
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
        values=["Primera", "Segunda", "Tercera", "Cuarta", "Quinta", "Sexta"],
        command=lambda _: validar(),
    ).pack(anchor="w", padx=20, pady=(0, 10))

    # =====================================================
    # ④ GENERAR
    # =====================================================
    def mostrar_resultado(resultado):
        boton_gen.configure(state="normal")

        resumen = (
            f"{resultado.actualizados} de {resultado.total_filas} docente(s) del control de pagos "
            f"fueron actualizados.\n"
            f"- Contratos terminados: {len(resultado.contratos_terminados)}\n"
            f"- Contratos por acabar (saldo bajo): {len(resultado.contratos_por_acabar)}\n"
            f"- Contratos con excedente: {len(resultado.contratos_con_excedente)}"
        )

        if resultado.sin_coincidencia:
            barra.completar("Completado con advertencias")
            lista = "\n".join(f"- {nombre}" for nombre in resultado.sin_coincidencia)
            messagebox.showwarning(
                "Control de pagos actualizado con advertencias",
                f"{resumen}\n\n"
                f"{len(resultado.sin_coincidencia)} docente(s) del archivo de control NO se encontraron "
                f"en la planilla y NO se actualizaron:\n{lista}\n\n"
                "Revisa que el nombre esté escrito igual en ambos archivos."
            )
        else:
            barra.completar("Completado")
            messagebox.showinfo("Éxito", f"Control de pagos actualizado.\n\n{resumen}")

    def generar():
        if not ruta_planilla or not ruta_control or not numero_armada.get():
            messagebox.showerror("Error", "Por favor, complete todos los campos antes de generar.")
            return

        armada = numero_armada.get()

        boton_gen.configure(state="disabled")
        barra.iniciar()

        def tarea():
            return actualizar_control_pagos(
                ruta_planilla,
                ruta_control,
                armada,
                on_progreso=lambda actual, total: barra.actualizar(actual, total, "filas revisadas"),
            )

        def al_error(e):
            boton_gen.configure(state="normal")
            barra.completar("Error")
            messagebox.showerror("Error", f"No se pudo procesar el Excel: {e}")

        consola.ejecutar(tarea, on_exito=mostrar_resultado, on_error=al_error)

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
    # ⑤ PROGRESO
    # =====================================================
    barra = BarraProgreso(contenedor, root)
    barra.frame.pack(fill="x", padx=20, pady=(0, 10))

    # =====================================================
    # ⑥ CONSOLA
    # =====================================================
    consola_frame = ctk.CTkFrame(contenedor, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(fill="both", expand=True, pady=10)

    consola_text = ctk.CTkTextbox(consola_frame, height=120, wrap="word", fg_color=CONSOLE_BG, text_color=WHITE_COLOR)
    consola_text.pack(padx=10, pady=10, fill="both", expand=True)
    consola_text.configure(state="disabled")

    consola = ConsolaEnVivo(consola_text, root)
