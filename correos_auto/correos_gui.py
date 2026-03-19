import customtkinter as ctk
import os
import pandas as pd
import sys
import threading
from tkinter import filedialog
from datetime import datetime

from .envio_correos import (
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


def iniciar_interfaz_correos(callback_volver=None):

    # =========================
    # CONFIG GENERAL
    # =========================
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Envío de correos – CEID")
    root.after(100, lambda: root.state("zoomed"))
    root.configure(fg_color=BG_COLOR)

    # =========================
    # VARIABLES DE ESTADO
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
    # FUNCIONES AUXILIARES
    # =========================
    def set_estado(label, texto, color="#B0B0B0"):
        label.configure(text=texto, text_color=color)

    def es_modo_contrato() -> bool:
        return modo_var.get() == "Primera vez con contrato (individual)"

    def obtener_tipo_correo() -> TipoCorreo:
        return TipoCorreo.DOCENTE if tipo_var.get() == "Docente" else TipoCorreo.ADMINISTRATIVO

    def obtener_tipo_texto() -> str:
        return "docente" if tipo_var.get() == "Docente" else "administrativo"

    def actualizar_ayuda_visual():
        if es_modo_contrato():
            txt = (
                "Flujo individual con contrato:\n"
                "1) Selecciona Excel + PDF de orden + PDF de contrato.\n"
                "2) Clic en 'Validar envío individual'.\n"
                "3) Revisa el destinatario detectado.\n"
                "4) Clic en 'Enviar correo individual (orden + contrato)'."
            )
        else:
            txt = (
                "Flujo masivo:\n"
                "1) Selecciona Excel + lote de PDFs de orden.\n"
                "2) Clic en 'Validar lote de correos'.\n"
                "3) Clic en 'Enviar correos masivos'."
            )
        lbl_ayuda.configure(text=txt)

    def validar_generar():
        if es_modo_contrato():
            valido = bool(ruta_excel and pdf_orden and pdf_contrato)
        else:
            valido = bool(ruta_excel and pdfs)
        btn_generar.configure(state="normal" if valido else "disabled")

    def validar_envio():
        if es_modo_contrato():
            btn_enviar.configure(state="normal" if data_envio_individual else "disabled")
        else:
            btn_enviar.configure(state="normal" if data_envio else "disabled")

    # =========================
    # HEADER
    # =========================
    titulo(root, "Envío de Correos – CEID")

    contenedor = ctk.CTkScrollableFrame(root, fg_color=BG_COLOR)
    contenedor.pack(fill="both", expand=True, padx=40, pady=20)

    # =====================================================
    # ① CONFIGURACIÓN DEL ENVÍO
    # =====================================================
    frame_conf = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_conf.pack(fill="x", pady=10)

    etiqueta(frame_conf, "① Configuración del envío")

    ctk.CTkLabel(frame_conf, text="Modo de envío", text_color=WHITE_COLOR).pack(anchor="w", padx=30, pady=(0, 2))
    menu_modo = crear_option_menu(
        frame_conf,
        modo_var,
        ["Masivo (solo orden)", "Primera vez con contrato (individual)"]
    )

    ctk.CTkLabel(frame_conf, text="Tipo de personal", text_color=WHITE_COLOR).pack(anchor="w", padx=30, pady=(0, 2))
    menu_tipo = crear_option_menu(frame_conf, tipo_var, ["Docente", "Administrativo"])

    ctk.CTkLabel(frame_conf, text="Mes de la orden", text_color=WHITE_COLOR).pack(anchor="w", padx=30, pady=(0, 2))
    menu_mes = crear_option_menu(frame_conf, mes_var, meses)

    años = [str(a) for a in range(datetime.now().year - 2, datetime.now().year + 3)]
    ctk.CTkLabel(frame_conf, text="Año de la orden", text_color=WHITE_COLOR).pack(anchor="w", padx=30, pady=(0, 2))
    menu_anio = crear_option_menu(frame_conf, año_var, años)

    ctk.CTkLabel(frame_conf, text="Mes inicio del contrato (solo modo individual)", text_color=WHITE_COLOR).pack(anchor="w", padx=30, pady=(0, 2))
    menu_mes_inicio = crear_option_menu(frame_conf, mes_inicio_contrato_var, meses)

    ctk.CTkLabel(frame_conf, text="Mes fin del contrato (solo modo individual)", text_color=WHITE_COLOR).pack(anchor="w", padx=30, pady=(0, 2))
    menu_mes_fin = crear_option_menu(frame_conf, mes_fin_contrato_var, meses)

    # =====================================================
    # ② ARCHIVOS REQUERIDOS
    # =====================================================
    frame_archivos = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_archivos.pack(fill="x", pady=10)

    etiqueta(frame_archivos, "② Archivos requeridos")

    # ---- EXCEL ----
    lbl_excel = ctk.CTkLabel(frame_archivos, text="📂 Excel no seleccionado", text_color="#B0B0B0")
    lbl_excel.pack(anchor="w", padx=30, pady=5)

    def seleccionar_excel():
        nonlocal ruta_excel
        ruta = filedialog.askopenfilename(
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta:
            try:
                pd.ExcelFile(ruta)
                ruta_excel = ruta
                set_estado(lbl_excel, f"✅ {os.path.basename(ruta)}", "#4CAF50")
                validar_generar()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")

    ctk.CTkButton(
        frame_archivos,
        text="Seleccionar Excel (BBDD)",
        command=seleccionar_excel,
        width=180
    ).pack(anchor="w", padx=30)

    # ---- PDFs ----
    lbl_pdfs = ctk.CTkLabel(frame_archivos, text="📂 Ningún PDF seleccionado", text_color="#B0B0B0")
    lbl_pdfs.pack(anchor="w", padx=30, pady=5)

    def seleccionar_pdfs():
        nonlocal pdfs
        rutas = filedialog.askopenfilenames(
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if rutas:
            pdfs = list(rutas)
            set_estado(lbl_pdfs, f"✅ {len(pdfs)} PDFs seleccionados", "#4CAF50")
            validar_generar()

    btn_pdfs = ctk.CTkButton(
        frame_archivos,
        text="Seleccionar PDFs",
        command=seleccionar_pdfs,
        width=180
    )
    btn_pdfs.pack(anchor="w", padx=30, pady=(0, 10))

    # ---- PDF ORDEN (modo individual con contrato) ----
    lbl_pdf_orden = ctk.CTkLabel(frame_archivos, text="📂 PDF de orden no seleccionado", text_color="#B0B0B0")
    lbl_pdf_orden.pack(anchor="w", padx=30, pady=5)

    def seleccionar_pdf_orden():
        nonlocal pdf_orden
        ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if ruta:
            pdf_orden = ruta
            set_estado(lbl_pdf_orden, f"✅ Orden: {os.path.basename(ruta)}", "#4CAF50")
            validar_generar()

    btn_pdf_orden = ctk.CTkButton(
        frame_archivos,
        text="Seleccionar PDF de orden",
        command=seleccionar_pdf_orden,
        width=180
    )
    btn_pdf_orden.pack(anchor="w", padx=30, pady=(0, 10))

    # ---- PDF CONTRATO (modo individual con contrato) ----
    lbl_pdf_contrato = ctk.CTkLabel(frame_archivos, text="📂 PDF de contrato no seleccionado", text_color="#B0B0B0")
    lbl_pdf_contrato.pack(anchor="w", padx=30, pady=5)

    def seleccionar_pdf_contrato():
        nonlocal pdf_contrato
        ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if ruta:
            pdf_contrato = ruta
            set_estado(lbl_pdf_contrato, f"✅ Contrato: {os.path.basename(ruta)}", "#4CAF50")
            validar_generar()

    btn_pdf_contrato = ctk.CTkButton(
        frame_archivos,
        text="Seleccionar PDF de contrato",
        command=seleccionar_pdf_contrato,
        width=180
    )
    btn_pdf_contrato.pack(anchor="w", padx=30, pady=(0, 10))

    def actualizar_modo(*_):
        nonlocal data_envio, data_envio_individual
        data_envio = []
        data_envio_individual = None
        btn_enviar.configure(state="disabled")

        if es_modo_contrato():
            menu_mes_inicio.configure(state="normal")
            menu_mes_fin.configure(state="normal")
            btn_pdfs.configure(state="disabled")
            btn_pdf_orden.configure(state="normal")
            btn_pdf_contrato.configure(state="normal")
            btn_generar.configure(text="Validar envío individual")
            btn_enviar.configure(text="Enviar individual: orden + contrato")
            set_estado(lbl_pdfs, "📂 PDFs para lote: no aplica en modo individual", "#B0B0B0")
            set_estado(lbl_estado_data, "⏳ Aún no se han generado datos", "#B0B0B0")
        else:
            menu_mes_inicio.configure(state="disabled")
            menu_mes_fin.configure(state="disabled")
            btn_pdfs.configure(state="normal")
            btn_pdf_orden.configure(state="disabled")
            btn_pdf_contrato.configure(state="disabled")
            btn_generar.configure(text="Validar lote de correos")
            btn_enviar.configure(text="Enviar correos masivos")
            set_estado(lbl_estado_data, "⏳ Aún no se han generado datos", "#B0B0B0")

        actualizar_ayuda_visual()
        validar_generar()

    # =====================================================
    # ③ GENERAR Y VALIDAR DATOS
    # =====================================================
    frame_gen = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_gen.pack(fill="x", pady=10)

    etiqueta(frame_gen, "③ Generar y validar datos")

    lbl_ayuda = ctk.CTkLabel(
        frame_gen,
        text="",
        justify="left",
        wraplength=980,
        text_color=WHITE_COLOR
    )
    lbl_ayuda.pack(anchor="w", padx=30, pady=(0, 8))

    lbl_estado_data = ctk.CTkLabel(
        frame_gen,
        text="⏳ Aún no se han generado datos",
        text_color="#B0B0B0"
    )
    lbl_estado_data.pack(anchor="w", padx=30)

    def generar_data():
        nonlocal data_envio, data_envio_individual
        btn_generar.configure(state="disabled")
        set_estado(lbl_estado_data, "⏳ Generando datos...", "#FFA726")

        def tarea():
            nonlocal data_envio, data_envio_individual
            try:
                if es_modo_contrato():
                    data_envio = []
                    data_envio_individual = procesar_correo_individual_contrato_primera_vez(
                        ruta_excel=ruta_excel,
                        hoja="list",
                        pdf_orden_path=pdf_orden,
                        tipo=obtener_tipo_correo()
                    )

                    if not data_envio_individual:
                        set_estado(lbl_estado_data, "⚠️ No se pudo validar la orden individual", "#FF7043")
                    else:
                        set_estado(
                            lbl_estado_data,
                            f"✅ Listo: {data_envio_individual['nombre']} - {data_envio_individual['correo']}",
                            "#4CAF50"
                        )
                else:
                    data_envio_individual = None
                    if tipo_var.get() == "Docente":
                        data_envio = procesar_correos_docente_gmail(ruta_excel, "list", pdfs)
                    else:
                        data_envio = procesar_correos_administrativos_gmail(ruta_excel, "list", pdfs)

                    if not data_envio:
                        set_estado(lbl_estado_data, "⚠️ No se encontraron coincidencias", "#FF7043")
                    else:
                        set_estado(
                            lbl_estado_data,
                            f"✅ {len(data_envio)} correos listos para envío",
                            "#4CAF50"
                        )

                validar_envio()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo procesar:\n{e}")
            finally:
                btn_generar.configure(state="normal")

        threading.Thread(target=tarea).start()

    btn_generar = ctk.CTkButton(
        frame_gen,
        text="Validar lote de correos",
        state="disabled",
        command=generar_data,
        width=200
    )
    btn_generar.pack(anchor="w", padx=30, pady=10)

    modo_var.trace_add("write", actualizar_modo)

    # =====================================================
    # CONSOLA
    # =====================================================
    consola_frame = ctk.CTkFrame(contenedor, fg_color=CONSOLE_BG, corner_radius=12)
    consola_frame.pack(fill="both", expand=True, pady=10)

    consola = ctk.CTkTextbox(
        consola_frame,
        fg_color=CONSOLE_BG,
        text_color=WHITE_COLOR,
        wrap="word"
    )
    consola.pack(fill="both", expand=True, padx=10, pady=10)
    consola.configure(state="disabled")

    class Redirector:
        def write(self, msg):
            consola.configure(state="normal")
            consola.insert("end", msg)
            consola.see("end")
            consola.configure(state="disabled")
        def flush(self): pass

    sys.stdout = Redirector()
    sys.stderr = Redirector()

    # =====================================================
    # ④ ENVÍO FINAL
    # =====================================================
    frame_envio = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_envio.pack(fill="x", pady=10)

    etiqueta(frame_envio, "④ Envío final")

    def enviar():
        if es_modo_contrato() and not data_envio_individual:
            return

        if not es_modo_contrato() and not data_envio:
            return

        if es_modo_contrato():
            resumen = (
                "Se enviará 1 correo (primera vez con contrato)\n\n"
                f"Destinatario: {data_envio_individual['nombre']}\n"
                f"Correo: {data_envio_individual['correo']}\n"
                f"Periodo contrato: {mes_inicio_contrato_var.get()} - {mes_fin_contrato_var.get()}\n\n"
                "¿Desea continuar?"
            )
        else:
            resumen = (
                f"Se enviarán {len(data_envio)} correos\n\n"
                f"¿Desea continuar?"
            )

        if not messagebox.askyesno("Confirmar envío", resumen):
            return

        btn_enviar.configure(state="disabled")

        def tarea_envio():
            try:
                año = int(año_var.get())
                if es_modo_contrato():
                    enviado = enviar_correo_contrato_primera_vez_desde_gui(
                        datos_envio=data_envio_individual,
                        pdf_contrato_path=pdf_contrato,
                        mes=mes_var.get(),
                        mes_inicio_contrato=mes_inicio_contrato_var.get(),
                        mes_fin_contrato=mes_fin_contrato_var.get(),
                        tipo=obtener_tipo_texto(),
                        anio=año
                    )
                    if enviado:
                        messagebox.showinfo("Éxito", "Correo de primera vez enviado correctamente.")
                    else:
                        messagebox.showerror("Error", "No se pudo enviar el correo de primera vez.")
                else:
                    if tipo_var.get() == "Docente":
                        enviar_lote_desde_gui_docentes(data_envio, mes_var.get(), año)
                    else:
                        enviar_lote_desde_gui_administrativos(data_envio, mes_var.get(), año)

                    messagebox.showinfo("Éxito", "Correos enviados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Fallo en el envío:\n{e}")
            finally:
                btn_enviar.configure(state="normal")

        threading.Thread(target=tarea_envio).start()

    btn_enviar = ctk.CTkButton(
        frame_envio,
        text="Enviar correos masivos",
        width=420,
        height=45,
        state="disabled",
        command=enviar
    )
    btn_enviar.pack(padx=30, pady=15)

    actualizar_modo()

    # =========================
    # FOOTER
    # =========================
    boton_volver(root, callback_volver)
    footer(root)
    root.mainloop()