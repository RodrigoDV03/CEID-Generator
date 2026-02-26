import customtkinter as ctk
import os
import pandas as pd
import sys
import threading
from tkinter import filedialog
from datetime import datetime

from .envio_correos import (
    procesar_correos_docente_gmail,
    procesar_correos_administrativos_gmail,
    enviar_lote_desde_gui_docentes,
    enviar_lote_desde_gui_administrativos
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
    tipo_var = ctk.StringVar(value="Docente")
    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    año_var = ctk.StringVar(value=str(datetime.now().year))

    ruta_excel = ""
    pdfs = []
    data_envio = []

    # =========================
    # FUNCIONES AUXILIARES
    # =========================
    def set_estado(label, texto, color="#B0B0B0"):
        label.configure(text=texto, text_color=color)

    def validar_generar():
        valido = bool(ruta_excel and pdfs)
        btn_generar.configure(state="normal" if valido else "disabled")

    def validar_envio():
        btn_enviar.configure(
            state="normal" if data_envio else "disabled"
        )

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

    crear_option_menu(frame_conf, tipo_var, ["Docente", "Administrativo"])
    crear_option_menu(frame_conf, mes_var, meses)

    años = [str(a) for a in range(datetime.now().year - 2, datetime.now().year + 3)]
    crear_option_menu(frame_conf, año_var, años)

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
        text="Seleccionar Excel",
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

    ctk.CTkButton(
        frame_archivos,
        text="Seleccionar PDFs",
        command=seleccionar_pdfs,
        width=180
    ).pack(anchor="w", padx=30, pady=(0, 10))

    # =====================================================
    # ③ GENERAR Y VALIDAR DATOS
    # =====================================================
    frame_gen = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_gen.pack(fill="x", pady=10)

    etiqueta(frame_gen, "③ Generar y validar datos")

    lbl_estado_data = ctk.CTkLabel(
        frame_gen,
        text="⏳ Aún no se han generado datos",
        text_color="#B0B0B0"
    )
    lbl_estado_data.pack(anchor="w", padx=30)

    def generar_data():
        nonlocal data_envio
        btn_generar.configure(state="disabled")
        set_estado(lbl_estado_data, "⏳ Generando datos...", "#FFA726")

        def tarea():
            nonlocal data_envio
            try:
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
        text="Generar datos",
        state="disabled",
        command=generar_data,
        width=200
    )
    btn_generar.pack(anchor="w", padx=30, pady=10)

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
        if not data_envio:
            return

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
        text="Enviar correos",
        height=45,
        state="disabled",
        command=enviar
    )
    btn_enviar.pack(padx=30, pady=15)

    # =========================
    # FOOTER
    # =========================
    boton_volver(root, callback_volver)
    footer(root)
    root.mainloop()