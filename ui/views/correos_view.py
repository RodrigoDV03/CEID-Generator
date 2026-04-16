import customtkinter as ctk
import os
import sys
import threading
from tkinter import filedialog
from datetime import datetime

from services.correo_service import (
    generar_data_correo_service,
    enviar_previsualizaciones_service,
    previsualizar_correos_service,
)
from ui.modals.preview_correos_modal import PreviewCorreosModal
from ui.components import TextRedirector
from utils.gui_constants import *
from utils import custom_modals as messagebox


def mostrar_correos(app):

    app.limpiar_contenido()
    app.titulo_header.configure(text="Envío de correos")

    # =========================
    # VARIABLES
    # =========================
    modo_var = ctk.StringVar(value="Masivo (solo orden)")
    tipo_var = ctk.StringVar(value="Docente")
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))

    mes_inicio_contrato_var = ctk.StringVar(value=mes_var.get())
    mes_fin_contrato_var = ctk.StringVar(value=mes_var.get())

    ruta_excel = ""
    pdfs = []
    pdf_orden = ""
    pdf_contrato = ""

    data_envio = []
    data_envio_individual_contenedor = [None]  # Usar contenedor para nonlocal
    previsualizaciones = []

    # =========================
    # CONTENEDOR
    # =========================
    contenedor = ctk.CTkScrollableFrame(app.main, fg_color="transparent")
    contenedor.pack(fill="both", expand=True, padx=30, pady=20)

    # =====================================================
    # HELPERS
    # =====================================================
    def set_estado(label, texto, color=TEXT_LIGHT):
        label.configure(text=texto, text_color=color)

    def actualizar_estado_pdf(indicador, ruta):
        texto = os.path.basename(ruta) if ruta else "No seleccionado"
        color = "#4CAF50" if ruta else TEXT_LIGHT
        set_estado(indicador, texto, color)

    def es_modo_contrato():
        return modo_var.get() == "Primera vez con contrato (individual)"

    def es_modo_reconocimiento_deuda():
        return modo_var.get() == "Reconocimiento de deuda (solo orden)"

    # =====================================================
    # ① CONFIG
    # =====================================================
    frame_conf = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_conf.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_conf, text="Configuración", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    ctk.CTkLabel(frame_conf, text="Modo de envío", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    menu_modo = ctk.CTkOptionMenu(frame_conf, variable=modo_var,
                                 values=[
                                     "Masivo (solo orden)",
                                     "Reconocimiento de deuda (solo orden)",
                                     "Primera vez con contrato (individual)"
                                 ])
    menu_modo.pack(anchor="w", padx=20, pady=(0, 10))

    ctk.CTkLabel(frame_conf, text="Tipo de destinatario", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_conf, variable=tipo_var,
                      values=["Docente", "Administrativo"])\
        .pack(anchor="w", padx=20, pady=(0, 10))

    ctk.CTkLabel(frame_conf, text="Mes", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_conf, variable=mes_var, values=meses)\
        .pack(anchor="w", padx=20, pady=(0, 10))

    ctk.CTkLabel(frame_conf, text="Año", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    ctk.CTkOptionMenu(frame_conf, variable=año_var, values=años)\
        .pack(anchor="w", padx=20, pady=(0, 10))

    ctk.CTkLabel(frame_conf, text="Mes de inicio (contrato)", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    menu_mes_inicio = ctk.CTkOptionMenu(frame_conf, variable=mes_inicio_contrato_var, values=meses)
    menu_mes_inicio.pack(anchor="w", padx=20, pady=(0, 10))

    ctk.CTkLabel(frame_conf, text="Mes de fin (contrato)", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    menu_mes_fin = ctk.CTkOptionMenu(frame_conf, variable=mes_fin_contrato_var, values=meses)
    menu_mes_fin.pack(anchor="w", padx=20, pady=(0, 10))
    # Función para actualizar estado de menús según modo
    def actualizar_modo(*_):
        es_masivo = not es_modo_contrato()
        estado = "disabled" if es_masivo else "normal"
        menu_mes_inicio.configure(state=estado)
        menu_mes_fin.configure(state=estado)

    # =====================================================
    # ② ARCHIVOS
    # =====================================================
    frame_files = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_files.pack(fill="x", pady=10)

    ctk.CTkLabel(frame_files, text="Archivos", font=FONT_SECTION, text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(15, 10))

    ctk.CTkLabel(frame_files, text="Excel con datos", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    lbl_excel = ctk.CTkLabel(frame_files, text="No seleccionado", text_color=TEXT_LIGHT)
    lbl_excel.pack(anchor="w", padx=20)

    def seleccionar_excel():
        nonlocal ruta_excel
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if ruta:
            ruta_excel = ruta
            set_estado(lbl_excel, os.path.basename(ruta), "#4CAF50")
            validar_generar()

    ctk.CTkButton(frame_files, text="Seleccionar Excel",
                  fg_color=PRIMARY_COLOR, hover_color=ACCENT_COLOR,
                  command=seleccionar_excel)\
        .pack(anchor="e", padx=20, pady=10)

    ctk.CTkLabel(frame_files, text="PDFs para enviar", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    lbl_pdfs = ctk.CTkLabel(frame_files, text="0 PDFs", text_color=TEXT_LIGHT)
    lbl_pdfs.pack(anchor="w", padx=20)

    def seleccionar_pdfs():
        nonlocal pdfs
        rutas = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if rutas:
            pdfs = list(rutas)
            set_estado(lbl_pdfs, f"{len(pdfs)} PDFs", "#4CAF50")
            validar_generar()

    btn_pdfs = ctk.CTkButton(frame_files, text="Seleccionar PDFs",
                             fg_color=PRIMARY_COLOR, hover_color=ACCENT_COLOR,
                             command=seleccionar_pdfs)
    btn_pdfs.pack(anchor="e", padx=20, pady=5)

    frame_contrato = ctk.CTkFrame(frame_files, fg_color="transparent")

    ctk.CTkLabel(frame_contrato, text="Orden de servicio (PDF)", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    lbl_orden = ctk.CTkLabel(frame_contrato, text="No seleccionado", text_color=TEXT_LIGHT)
    lbl_orden.pack(anchor="w", padx=20)

    def seleccionar_pdf_orden():
        nonlocal pdf_orden
        ruta = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if ruta:
            pdf_orden = ruta
            actualizar_estado_pdf(lbl_orden, pdf_orden)
            validar_generar()

    ctk.CTkButton(
        frame_contrato,
        text="Seleccionar orden",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_pdf_orden
    ).pack(anchor="e", padx=20, pady=5)

    ctk.CTkLabel(frame_contrato, text="Contrato firmado (PDF)", text_color=TEXT_COLOR)\
        .pack(anchor="w", padx=20, pady=(10, 0))
    lbl_contrato = ctk.CTkLabel(frame_contrato, text="No seleccionado", text_color=TEXT_LIGHT)
    lbl_contrato.pack(anchor="w", padx=20)

    def seleccionar_pdf_contrato():
        nonlocal pdf_contrato
        ruta = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if ruta:
            pdf_contrato = ruta
            actualizar_estado_pdf(lbl_contrato, pdf_contrato)
            validar_generar()

    ctk.CTkButton(
        frame_contrato,
        text="Seleccionar contrato",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=seleccionar_pdf_contrato
    ).pack(anchor="e", padx=20, pady=(5, 10))

    def actualizar_campos_modo(*_):
        es_contrato = es_modo_contrato()

        if es_contrato:
            frame_contrato.pack(fill="x", pady=(5, 0))
            lbl_pdfs.configure(text_color=TEXT_LIGHT)
            btn_pdfs.configure(state="disabled")
        else:
            frame_contrato.pack_forget()
            lbl_pdfs.configure(text_color=TEXT_LIGHT)
            btn_pdfs.configure(state="normal")

    modo_var.trace_add("write", actualizar_modo)
    modo_var.trace_add("write", actualizar_campos_modo)

    # =====================================================
    # VALIDACIÓN
    # =====================================================
    def validar_generar():
        if es_modo_contrato():
            btn_generar.configure(state="normal" if ruta_excel and pdf_orden and pdf_contrato else "disabled")
        else:
            btn_generar.configure(state="normal" if ruta_excel and pdfs else "disabled")

    root = app.root if hasattr(app, "root") else app.winfo_toplevel()

    # =====================================================
    # ③ VALIDAR
    # =====================================================
    frame_validar = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_validar.pack(fill="x", pady=10)

    estado_lbl = ctk.CTkLabel(frame_validar, text="Sin validar", text_color=TEXT_LIGHT)
    estado_lbl.pack(anchor="w", padx=20, pady=(12, 5))

    def abrir_modal_previsualizacion():
        if not previsualizaciones:
            messagebox.showwarning("Advertencia", "No hay correos para previsualizar")
            return

        def on_save(previews_editadas):
            previsualizaciones[:] = previews_editadas
            set_estado(estado_lbl, "Previsualización actualizada", "#FFD700")

        def on_send(previews_editadas):
            previsualizaciones[:] = previews_editadas

            def tarea_envio():
                try:
                    resumen = enviar_previsualizaciones_service(previsualizaciones)
                    root.after(
                        0,
                        lambda: messagebox.showinfo(
                            "Éxito",
                            f"Envío completado: {resumen['exitosos']} exitosos, {resumen['fallidos']} fallidos",
                        ),
                    )
                except Exception as e:
                    root.after(0, lambda: messagebox.showerror("Error", str(e)))

            threading.Thread(target=tarea_envio, daemon=True).start()

        PreviewCorreosModal(root, previsualizaciones, on_save=on_save, on_send=on_send)

    def generar_data():
        nonlocal data_envio

        def tarea():
            try:
                data_envio_resultado, data_envio_individual_resultado = generar_data_correo_service(
                    es_modo_contrato=es_modo_contrato(),
                    tipo_var=tipo_var.get(),
                    ruta_excel=ruta_excel,
                    pdfs=pdfs,
                    pdf_orden=pdf_orden,
                )

                if es_modo_contrato():
                    if not data_envio_individual_resultado:
                        root.after(0, lambda: set_estado(estado_lbl, "Error al procesar correo", ACCENT_COLOR))
                        return

                    data_envio_individual_contenedor[0] = data_envio_individual_resultado
                    data_envio[:] = []
                    lista_para_preview = [data_envio_individual_resultado]
                    texto_estado = "Correo individual listo"
                else:
                    data_envio[:] = data_envio_resultado
                    data_envio_individual_contenedor[0] = None
                    lista_para_preview = data_envio
                    texto_estado = f"{len(data_envio)} correos listos"

                previsualizaciones_resultado = previsualizar_correos_service(
                    data_envio=lista_para_preview,
                    mes=mes_var.get(),
                    tipo_var=tipo_var.get(),
                    anio=int(año_var.get()),
                    es_reconocimiento_deuda=es_modo_reconocimiento_deuda(),
                    es_modo_contrato=es_modo_contrato(),
                    mes_inicio=mes_inicio_contrato_var.get(),
                    mes_fin=mes_fin_contrato_var.get(),
                    pdf_contrato=pdf_contrato,
                )

                def aplicar_resultado():
                    set_estado(estado_lbl, texto_estado, "#4CAF50")
                    previsualizaciones[:] = previsualizaciones_resultado
                    abrir_modal_previsualizacion()

                root.after(0, aplicar_resultado)

            except Exception as e:
                root.after(0, lambda: messagebox.showerror("Error", str(e)))

        threading.Thread(target=tarea, daemon=True).start()

    btn_generar = ctk.CTkButton(
        frame_validar,
        text="Validar",
        fg_color=ACCENT_COLOR,
        state="disabled",
        command=generar_data,
    )
    btn_generar.pack(padx=20, pady=(5, 12))

    def al_cambiar_modo(*_):
        previsualizaciones.clear()
        set_estado(estado_lbl, "Sin validar", TEXT_LIGHT)
        validar_generar()

    modo_var.trace_add("write", al_cambiar_modo)

    actualizar_modo()
    actualizar_campos_modo()
    validar_generar()

    # =====================================================
    # CONSOLA
    # =====================================================
    consola_frame = ctk.CTkFrame(contenedor, fg_color=CONSOLE_BG)
    consola_frame.pack(fill="both", expand=True, pady=10)

    consola = ctk.CTkTextbox(consola_frame, fg_color=CONSOLE_BG, text_color=WHITE_COLOR)
    consola.pack(fill="both", expand=True, padx=10, pady=10)
    consola.configure(state="disabled")

    sys.stdout = TextRedirector(consola)
    sys.stderr = TextRedirector(consola)