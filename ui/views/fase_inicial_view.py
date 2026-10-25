import customtkinter as ctk
from tkinter import filedialog
import os
import pandas as pd

from ui.console_stream import ConsolaEnVivo
from ui.progress import BarraProgreso
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
    mes_var = ctk.StringVar(value=mes_actual())
    numero_armada = ctk.StringVar()
    tipo_fase_inicial = ctk.StringVar(value="planilla docente")

    planilla_path = None
    carpeta_destino = None

    root = app.root if hasattr(app, "root") else app.winfo_toplevel()

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
                    values=["primera", "segunda", "tercera", "cuarta", "quinta", "sexta", "sin armada"],
                    command=lambda _: validar())\
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
    def mostrar_resultado(resultado):
        boton_gen.configure(state="normal")

        if resultado.sin_firma:
            barra.completar("Completado con advertencias")
            lista = "\n".join(f"- {nombre}" for nombre in resultado.sin_firma)
            messagebox.showwarning(
                "Documentos generados con advertencias",
                f"Se generaron {resultado.exitosos} de {resultado.total_procesados} documentos.\n\n"
                f"{len(resultado.sin_firma)} cotización(es) se generaron SIN firma "
                f"porque no se encontró el archivo de firma del docente:\n{lista}"
            )
        elif resultado.con_error:
            barra.completar("Completado con errores")
            lista = "\n".join(f"- {nombre}" for nombre in resultado.con_error)
            messagebox.showwarning(
                "Documentos generados con errores",
                f"Se generaron {resultado.exitosos} de {resultado.total_procesados} documentos.\n\n"
                f"No se pudieron generar los documentos de:\n{lista}\n\n"
                "Revisa la consola para el detalle de cada error."
            )
        else:
            barra.completar("Completado")
            messagebox.showinfo("Éxito", "Documentos generados correctamente")

    def generar():

        if not planilla_path or not carpeta_destino:
            messagebox.showerror("Error", "Faltan datos")
            return

        # Leer los valores de la pantalla en el hilo principal.
        hoja = hoja_var.get()
        mes = mes_var.get()
        armada = numero_armada.get()
        tipo = tipo_fase_inicial.get()

        boton_gen.configure(state="disabled")
        barra.iniciar()

        def tarea():
            return procesar_planilla_fase_inicial(
                planilla_path,
                hoja,
                carpeta_destino,
                mes,
                armada,
                tipo,
                on_progreso=lambda actual, total: barra.actualizar(actual, total, "docentes procesados"),
            )

        def al_error(e):
            boton_gen.configure(state="normal")
            barra.completar("Error")
            messagebox.showerror("Error", str(e))

        consola.ejecutar(tarea, on_exito=mostrar_resultado, on_error=al_error)

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
    # PROGRESO
    # =====================================================
    barra = BarraProgreso(contenedor, root)
    barra.frame.pack(fill="x", padx=20, pady=(0, 10))

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

    consola = ConsolaEnVivo(consola_text, root)
