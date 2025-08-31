import customtkinter as ctk
from correos_auto.correos_gui import iniciar_interfaz_correos
from generador_planilla.planilla_gui import iniciar_interfaz_planilla
from fases.fase_inicial.fase_inicial_gui import iniciar_interfaz_fase_inicial
from fases.fase_final.fase_final_gui import iniciar_interfaz_fase_final
from control_pagos.control_pagos_gui import iniciar_interfaz_control_pagos
from utils.gui_constants import *

def iniciar_interfaz_general():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    def abrir_planilla():
        root.destroy()
        iniciar_interfaz_planilla(volver_menu)

    def abrir_fase_inicial():
        root.destroy()
        iniciar_interfaz_fase_inicial(volver_menu)

    def abrir_fase_final():
        root.destroy()
        iniciar_interfaz_fase_final(volver_menu)

    def abrir_control_pagos():
        root.destroy()
        iniciar_interfaz_control_pagos(volver_menu)

    def abrir_envio_correos():
        root.destroy()
        iniciar_interfaz_correos(volver_menu)

    def volver_menu():
        iniciar_interfaz_general()

    # --- Ventana ---
    root = ctk.CTk()
    root.title("CEID - Sistema de Generación de Documentos")
    root.configure(fg_color=BG_COLOR)
    root.after(100, lambda: root.state("zoomed"))

    # --- Header ---
    titulo(root, "CEID - Generador de Documentos")

    # --- Contenedor principal ---
    main_frame = ctk.CTkFrame(root, fg_color="transparent")
    main_frame.pack(expand=True)

    # Configuración grid
    main_frame.grid_rowconfigure((0, 1, 2), weight=1, uniform="row")
    main_frame.grid_columnconfigure((0, 1), weight=1, uniform="col")

    # --- Cards estilo botón ---
    def crear_card(row, col, texto, comando):
        card = ctk.CTkButton(
            main_frame,
            text=texto,
            font=FONT_BUTTON,
            height=120,
            width=260,
            fg_color=BUTTON_BG_COLOR,
            hover_color=BUTTON_HOVER_BG_COLOR,
            text_color=WHITE_COLOR,
            corner_radius=20,
            command=comando
        )
        card.grid(row=row, column=col, padx=30, pady=25, sticky="nsew")

    # Añadir opciones
    crear_card(0, 0, "Generar Planilla", abrir_planilla)
    crear_card(0, 1, "Generar Fase Inicial", abrir_fase_inicial)
    crear_card(1, 0, "Generar Fase Final", abrir_fase_final)
    crear_card(1, 1, "Control de Pagos Docente", abrir_control_pagos)
    crear_card(2, 0, "Envío de Correos", abrir_envio_correos)

    # --- Footer ---
    footer(root)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz_general()