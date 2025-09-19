import customtkinter as ctk
from tkinter import filedialog
import os
from .generador_planilla import generar_planilla
from utils.gui_constants import *
from utils import custom_modals as messagebox

def iniciar_interfaz_planilla(callback_volver=None):
    archivo_cursos_path = ""
    archivo_docentes_path = ""
    archivo_clasif_path = ""
    archivo_planilla_anterior_path = ""

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Generador de Planilla - CEID")
    root.after(100, lambda: root.state("zoomed"))
    root.configure(fg_color=BG_COLOR)

    mes_var = ctk.StringVar(value=datetime.now().strftime("%B").capitalize())
    carga_var = ctk.IntVar(value=1)

    # --- Header ---
    titulo(root, "Generador de Planilla - CEID")

    # --- Tarjeta principal ---
    card = ctk.CTkFrame(root, fg_color=SECTION_COLOR, corner_radius=15)
    card.pack(padx=40, pady=20, fill="x", expand=False)

    # ---------------- SECCIÓN: Mes y carga ----------------
    etiqueta(card, "Seleccione el mes de la planilla a elaborar:")

    crear_option_menu(card, variable=mes_var, opciones=meses)

    lbl_carga = ctk.CTkLabel(card, text="Número de planilla:", font=FONT_SECTION, text_color=WHITE_COLOR)
    lbl_carga.pack(anchor="w", padx=20, pady=(10, 5))
    ctk.CTkRadioButton(card, text="1 (Primera planilla)", variable=carga_var, value=1, text_color=WHITE_COLOR).pack(anchor="w", padx=35)
    ctk.CTkRadioButton(card, text="2 (Segunda planilla)", variable=carga_var, value=2, text_color=WHITE_COLOR).pack(anchor="w", padx=35, pady=(0, 10))

    # ---------------- SECCIÓN: Planilla anterior ----------------
    marco_planilla_anterior = ctk.CTkFrame(card, fg_color="transparent")
    marco_planilla_anterior.pack(fill="x", padx=30, pady=10)

    seccion_titulo = ctk.CTkLabel(marco_planilla_anterior, text="Planilla anterior", font=FONT_SECTION, text_color=WHITE_COLOR)
    seccion_titulo.pack(anchor="w", pady=(0, 5))

    label_planilla_estado = ctk.CTkLabel(marco_planilla_anterior, text="📂 No seleccionado", text_color=WHITE_COLOR)
    label_planilla_estado.pack(side="left", padx=10, pady=10)

    def seleccionar_planilla_anterior():
        archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if archivo:
            nonlocal archivo_planilla_anterior_path
            archivo_planilla_anterior_path = archivo
            label_planilla_estado.configure(text=f"📁 {os.path.basename(archivo)}", text_color=WHITE_COLOR)

    btn_planilla_anterior = ctk.CTkButton(marco_planilla_anterior, text="Seleccionar archivo", command=seleccionar_planilla_anterior, width=160,
                                            fg_color=BUTTON_BG_COLOR, hover_color=BUTTON_HOVER_BG_COLOR, text_color=WHITE_COLOR, font=FONT_BUTTON)
    btn_planilla_anterior.pack(side="right", padx=15)

    def actualizar_visibilidad_planilla_anterior(*args):
        if carga_var.get() == 2:
            marco_planilla_anterior.pack(fill="x", padx=30, pady=10)
        else:
            marco_planilla_anterior.pack_forget()
            nonlocal archivo_planilla_anterior_path
            archivo_planilla_anterior_path = ""
            label_planilla_estado.configure(text="📂 No seleccionado", text_color=WHITE_COLOR)

    carga_var.trace_add("write", actualizar_visibilidad_planilla_anterior)
    actualizar_visibilidad_planilla_anterior()

    # ---------------- SECCIÓN: Archivos ----------------
    def crear_selector_archivo(master, texto_label, extensiones, actualizar_func):
        frame = ctk.CTkFrame(master, fg_color="transparent")
        frame.pack(fill="x", padx=30, pady=10)

        lbl = ctk.CTkLabel(frame, text=texto_label, font=FONT_SECTION, text_color=WHITE_COLOR)
        lbl.pack(anchor="w", pady=(0, 5))

        container = ctk.CTkFrame(frame, fg_color="transparent")
        container.pack(fill="x")

        label_estado = ctk.CTkLabel(container, text="📂 No seleccionado", text_color=WHITE_COLOR)
        label_estado.pack(side="left", padx=10, pady=10)

        def seleccionar():
            archivo = filedialog.askopenfilename(filetypes=[extensiones])
            if archivo:
                actualizar_func(archivo)
                label_estado.configure(text=f"📁 {os.path.basename(archivo)}", text_color=WHITE_COLOR)

        ctk.CTkButton(container, text="Seleccionar archivo", command=seleccionar, width=160, fg_color=BUTTON_BG_COLOR,
                        hover_color=BUTTON_HOVER_BG_COLOR, text_color=WHITE_COLOR, font=FONT_BUTTON).pack(side="right", padx=10)

    def actualizar_cursos(path): nonlocal archivo_cursos_path; archivo_cursos_path = path
    def actualizar_docentes(path): nonlocal archivo_docentes_path; archivo_docentes_path = path
    def actualizar_clasif(path): nonlocal archivo_clasif_path; archivo_clasif_path = path

    crear_selector_archivo(card, "Adjunte el archivo de la data de cursos del mes", ("Archivos CSV", "*.csv"), actualizar_cursos)
    crear_selector_archivo(card, "Adjunte el archivo de la lista de docentes", ("Archivos Excel", "*.xlsx *.xls"), actualizar_docentes)
    crear_selector_archivo(card, "Adjunte el archivo de Examen de Clasificación", ("Archivos Excel", "*.xlsx *.xls"), actualizar_clasif)

    # ---------------- BOTÓN DE PROCESAR ----------------
    def procesar():
        if not archivo_cursos_path or not archivo_docentes_path or not archivo_clasif_path:
            messagebox.showerror("Faltan archivos", "⚠️ Debes seleccionar los tres archivos requeridos.")
            return
        resultado = generar_planilla(archivo_cursos_path, archivo_docentes_path, archivo_clasif_path, mes_var.get(), carga_var.get(), archivo_planilla_anterior_path if carga_var.get()==2 else None)
        if resultado.startswith("Error"):
            messagebox.showerror("Error", resultado)
        else:
            messagebox.showinfo("Éxito", resultado)

    boton_generador(card, "Generar Planilla", procesar)

    # ---------------- Footer y volver ----------------
    boton_volver(root, callback_volver)

    footer(root)

    root.mainloop()
