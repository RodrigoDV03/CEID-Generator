import customtkinter as ctk
from generador_planilla.planilla_gui import iniciar_interfaz_planilla
from fase_inicial.fase_inicial_gui import iniciar_interfaz_fase_inicial
from fase_final.fase_final_gui import iniciar_interfaz_fase_final
from utils.constants import *

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

    def volver_menu():
        iniciar_interfaz_general()


    root = ctk.CTk()
    root.title("CEID - Sistema de Generación de Documentos")
    root.geometry("700x600")
    root.configure(fg_color=BG_COLOR)
    root.resizable(False, False)
    root.after(100, lambda: root.state("zoomed"))

    # --- TÍTULO ---
    titulo(root, "CEID - Generador de Documentos").pack(pady=(50, 20))

    # --- CONTENEDOR PRINCIPAL (FRAME) ---
    main_frame = ctk.CTkFrame(root, fg_color=BG_COLOR)
    main_frame.pack(pady=10)

    def crear_card(texto, comando):
        card = ctk.CTkFrame(main_frame, fg_color=CARD_COLOR, corner_radius=15)
        card.pack(pady=15, padx=20, fill="x", expand=True)

        boton = ctk.CTkButton(
            card, text=texto,
            font=FONT_BUTTON, height=50,
            fg_color=ACCENT_COLOR, hover_color=HOVER_COLOR,
            text_color=WHITE_COLOR, corner_radius=10,
            command=comando
        )
        boton.pack(padx=30, pady=20)

    crear_card("📋 1. Generar Planilla", abrir_planilla)
    crear_card("📄 2. Generador Fase Inicial", abrir_fase_inicial)
    crear_card("📁 3. Generador Fase Final", abrir_fase_final)

    # --- FOOTER ---
    ctk.CTkLabel(
        root, text="Centro de Idiomas - FLCH - UNMSM",
        font=FONT_FOOTER, text_color=GRAY_COLOR
    ).pack(side="bottom", pady=20)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz_general()