import customtkinter as ctk
from tkinter import filedialog
from datetime import datetime
import os

from .generador_planilla import generar_planilla
from utils.gui_constants import *
from utils import custom_modals as messagebox


def iniciar_interfaz_planilla(callback_volver=None):
    # =========================
    # VARIABLES DE ESTADO
    # =========================
    archivo_cursos = ""
    archivo_docentes = ""
    archivo_clasif = ""
    archivo_coordinacion = ""
    carpeta_destino = ""

    # =========================
    # CONFIG APP
    # =========================
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Generador de Planillas - CEID")
    root.after(100, lambda: root.state("zoomed"))
    root.configure(fg_color=BG_COLOR)

    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    bono_var = ctk.StringVar(value="0")

    # =========================
    # FUNCIONES AUXILIARES
    # =========================
    def estado(label, texto, color=WHITE_COLOR):
        label.configure(text=texto, text_color=color)

    def validar_formulario():
        valido = (
            archivo_cursos and
            archivo_docentes and
            archivo_clasif
        )
        btn_generar.configure(state="normal" if valido else "disabled")

    # =========================
    # HEADER
    # =========================
    titulo(root, "Generador de Planillas – CEID")

    contenedor = ctk.CTkScrollableFrame(root, fg_color=BG_COLOR)
    contenedor.pack(fill="both", expand=True, padx=40, pady=20)

    # =====================================================
    # ① CONFIGURACIÓN GENERAL
    # =====================================================
    frame_conf = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_conf.pack(fill="x", pady=10)

    etiqueta(frame_conf, "① Configuración General")

    crear_option_menu(frame_conf, mes_var, meses)

    # Bono solo Enero
    frame_bono = ctk.CTkFrame(frame_conf, fg_color="transparent")
    lbl_bono = ctk.CTkLabel(frame_bono, text="Monto bono (solo enero)", text_color=WHITE_COLOR)
    entry_bono = ctk.CTkEntry(frame_bono, textvariable=bono_var, width=200)

    def toggle_bono(*_):
        if mes_var.get() == "Enero":
            frame_bono.pack(anchor="w", padx=30, pady=10)
        else:
            frame_bono.pack_forget()
            bono_var.set("0")

    lbl_bono.pack(anchor="w")
    entry_bono.pack(anchor="w")
    mes_var.trace_add("write", toggle_bono)
    toggle_bono()

    # =====================================================
    # ② ARCHIVOS OBLIGATORIOS
    # =====================================================
    frame_req = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_req.pack(fill="x", pady=10)

    etiqueta(frame_req, "② Archivos obligatorios")

    def selector_archivo(texto, extensiones, setter):
        frame = ctk.CTkFrame(frame_req, fg_color="transparent")
        frame.pack(fill="x", padx=30, pady=5)

        lbl = ctk.CTkLabel(frame, text=texto, text_color=WHITE_COLOR)
        lbl.pack(anchor="w")

        estado_lbl = ctk.CTkLabel(frame, text="📂 No seleccionado", text_color="#B0B0B0")
        estado_lbl.pack(side="left", padx=10)

        def seleccionar():
            archivo = filedialog.askopenfilename(filetypes=[extensiones])
            if archivo:
                setter(archivo)
                estado(estado_lbl, f"✅ {os.path.basename(archivo)}", "#4CAF50")
                validar_formulario()

        ctk.CTkButton(
            frame, text="Seleccionar",
            command=seleccionar,
            width=140
        ).pack(side="right")

    def set_cursos(v): nonlocal archivo_cursos; archivo_cursos = v
    def set_docentes(v): nonlocal archivo_docentes; archivo_docentes = v
    def set_clasif(v): nonlocal archivo_clasif; archivo_clasif = v

    selector_archivo("Archivo de cursos (CSV)", ("CSV", "*.csv"), set_cursos)
    selector_archivo("Lista de docentes (Excel)", ("Excel", "*.xlsx *.xls"), set_docentes)
    selector_archivo("Examen de clasificación", ("Excel", "*.xlsx *.xls"), set_clasif)

    # =====================================================
    # ③ ARCHIVOS ADICIONALES
    # =====================================================
    frame_opt = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_opt.pack(fill="x", pady=10)

    etiqueta(frame_opt, "③ Archivos adicionales")

    def set_coord(v): nonlocal archivo_coordinacion; archivo_coordinacion = v
    selector_archivo("Coordinación / Actualización", ("Excel", "*.xlsx *.xls"), set_coord)

    # =====================================================
    # ④ DESTINO Y EJECUCIÓN
    # =====================================================
    frame_exec = ctk.CTkFrame(contenedor, fg_color=SECTION_COLOR, corner_radius=15)
    frame_exec.pack(fill="x", pady=10)

    etiqueta(frame_exec, "④ Destino y ejecución")

    lbl_destino = ctk.CTkLabel(
        frame_exec,
        text="📂 Carpeta destino: automática",
        text_color="#B0B0B0"
    )
    lbl_destino.pack(anchor="w", padx=30)

    def elegir_destino():
        nonlocal carpeta_destino
        carpeta = filedialog.askdirectory()
        if carpeta:
            carpeta_destino = carpeta
            estado(lbl_destino, f"📁 {carpeta}", "#4CAF50")

    ctk.CTkButton(
        frame_exec, text="Elegir carpeta",
        command=elegir_destino
    ).pack(anchor="w", padx=30, pady=10)

    # =========================
    # BOTÓN GENERAR
    # =========================
    def procesar():
        if not messagebox.askyesno("Confirmar", "¿Generar la planilla con los datos seleccionados?"):
            return

        resultado = generar_planilla(
            archivo_cursos,
            archivo_docentes,
            archivo_clasif,
            archivo_coordinacion,
            mes_var.get(),
            float(bono_var.get() or 0),
            carpeta_destino or None
        )

        if resultado.startswith("❌"):
            messagebox.showerror("Error", resultado)
        else:
            messagebox.showinfo("Éxito", resultado)

    btn_generar = ctk.CTkButton(
        contenedor,
        text="GENERAR PLANILLA",
        height=50,
        state="disabled",
        command=procesar
    )
    btn_generar.pack(pady=20)

    boton_volver(root, callback_volver)
    footer(root)
    root.mainloop()
