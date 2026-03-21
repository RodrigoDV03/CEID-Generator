import customtkinter as ctk
import os
import sys
import threading
from tkinter import filedialog
from datetime import datetime

from core.correos.envio_correos import (
    TipoCorreo,
    procesar_correos_docente_gmail,
    procesar_correos_administrativos_gmail,
    enviar_lote_desde_gui_docentes,
    enviar_lote_desde_gui_administrativos,
    procesar_correo_individual_contrato_primera_vez,
    enviar_correo_contrato_primera_vez_desde_gui
)

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
    data_envio_individual = None

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

    def es_modo_contrato():
        return modo_var.get() == "Primera vez con contrato (individual)"

    def obtener_tipo_correo():
        return TipoCorreo.DOCENTE if tipo_var.get() == "Docente" else TipoCorreo.ADMINISTRATIVO

    def obtener_tipo_texto():
        return "docente" if tipo_var.get() == "Docente" else "administrativo"

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
                                 values=["Masivo (solo orden)", "Primera vez con contrato (individual)"])
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
    años = [str(a) for a in range(datetime.now().year - 2, datetime.now().year + 3)]
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
        es_masivo = modo_var.get() == "Masivo (solo orden)"
        estado = "disabled" if es_masivo else "normal"
        menu_mes_inicio.configure(state=estado)
        menu_mes_fin.configure(state=estado)

    modo_var.trace_add("write", actualizar_modo)
    actualizar_modo()  # Llamar una vez para establecer estado inicial

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

    # =====================================================
    # VALIDACIÓN
    # =====================================================
    def validar_generar():
        if es_modo_contrato():
            btn_generar.configure(state="normal" if ruta_excel and pdf_orden and pdf_contrato else "disabled")
        else:
            btn_generar.configure(state="normal" if ruta_excel and pdfs else "disabled")

    def validar_envio():
        if es_modo_contrato():
            btn_enviar.configure(state="normal" if data_envio_individual else "disabled")
        else:
            btn_enviar.configure(state="normal" if data_envio else "disabled")

    # =====================================================
    # ③ VALIDAR
    # =====================================================
    frame_validar = ctk.CTkFrame(contenedor, fg_color=CARD_COLOR, corner_radius=15)
    frame_validar.pack(fill="x", pady=10)

    estado_lbl = ctk.CTkLabel(frame_validar, text="Sin validar", text_color=TEXT_LIGHT)
    estado_lbl.pack(anchor="w", padx=20)

    def generar_data():
        nonlocal data_envio, data_envio_individual

        def tarea():
            try:
                if es_modo_contrato():
                    data_envio_individual = procesar_correo_individual_contrato_primera_vez(
                        ruta_excel, "list", pdf_orden, obtener_tipo_correo()
                    )
                    if data_envio_individual:
                        set_estado(estado_lbl, "Correo individual listo", "#4CAF50")
                else:
                    if tipo_var.get() == "Docente":
                        data_envio[:] = procesar_correos_docente_gmail(ruta_excel, "list", pdfs)
                    else:
                        data_envio[:] = procesar_correos_administrativos_gmail(ruta_excel, "list", pdfs)

                    set_estado(estado_lbl, f"{len(data_envio)} correos listos", "#4CAF50")

                validar_envio()

            except Exception as e:
                messagebox.showerror("Error", str(e))

        threading.Thread(target=tarea).start()

    btn_generar = ctk.CTkButton(frame_validar, text="Validar",
                                fg_color=ACCENT_COLOR,
                                state="disabled",
                                command=generar_data)
    btn_generar.pack(padx=20, pady=10)

    # =====================================================
    # CONSOLA
    # =====================================================
    consola_frame = ctk.CTkFrame(contenedor, fg_color=CONSOLE_BG)
    consola_frame.pack(fill="both", expand=True, pady=10)

    consola = ctk.CTkTextbox(consola_frame, fg_color=CONSOLE_BG, text_color=WHITE_COLOR)
    consola.pack(fill="both", expand=True, padx=10, pady=10)
    consola.configure(state="disabled")

    class Redirector:
        def write(self, txt):
            consola.configure(state="normal")
            consola.insert("end", txt)
            consola.see("end")
            consola.configure(state="disabled")
        def flush(self): pass

    sys.stdout = Redirector()
    sys.stderr = Redirector()

    # =====================================================
    # ④ ENVIAR
    # =====================================================
    def enviar():
        def tarea():
            try:
                if es_modo_contrato():
                    enviar_correo_contrato_primera_vez_desde_gui(
                        data_envio_individual,
                        pdf_contrato,
                        mes_var.get(),
                        mes_inicio_contrato_var.get(),
                        mes_fin_contrato_var.get(),
                        obtener_tipo_texto(),
                        int(año_var.get())
                    )
                else:
                    if tipo_var.get() == "Docente":
                        enviar_lote_desde_gui_docentes(data_envio, mes_var.get(), int(año_var.get()))
                    else:
                        enviar_lote_desde_gui_administrativos(data_envio, mes_var.get(), int(año_var.get()))

                messagebox.showinfo("Éxito", "Correos enviados correctamente")

            except Exception as e:
                messagebox.showerror("Error", str(e))

        threading.Thread(target=tarea).start()

    btn_enviar = ctk.CTkButton(
        contenedor,
        text="Enviar correos",
        height=50,
        fg_color=ACCENT_COLOR,
        state="disabled",
        command=enviar
    )
    btn_enviar.pack(pady=20)