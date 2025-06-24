import customtkinter as ctk
from generador_planilla.interfaz_gui import iniciar_interfaz_planilla
from fase_inicial.interfaz_gui import iniciar_interfaz_fase_inicial
from fase_final.interfaz_gui import iniciar_interfaz_fase_final

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

    # --- COLORES Y FUENTES UNIFICADOS ---
    PRIMARY = "#2d415a"
    ACCENT = "#4a90e2"
    BG = "#e0e2e6"
    GRAY = "#a0a8b8"


    root = ctk.CTk()
    root.title("CEID - Sistema de Generación de Documentos")
    root.geometry("500x420")
    root.configure(fg_color=BG)
    root.resizable(True, True)

    FONT_TITLE = ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
    FONT_BUTTON = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
    FONT_FOOTER = ctk.CTkFont(family="Segoe UI", size=11)


    ctk.CTkLabel(
        root, text="CEID - Sistema de Documentos",
        font=FONT_TITLE, text_color=PRIMARY
    ).pack(pady=(30, 20))

    ctk.CTkButton(
        root, text="📋 1. Generar Planilla",
        width=300, height=40, font=FONT_BUTTON,
        fg_color=ACCENT, hover_color="#3a76c7", text_color="white",
        command=abrir_planilla
    ).pack(pady=10)

    ctk.CTkButton(
        root, text="📄 2. Generador fase inicial",
        width=300, height=40, font=FONT_BUTTON,
        fg_color=ACCENT, hover_color="#3a76c7", text_color="white",
        command=abrir_fase_inicial
    ).pack(pady=10)

    ctk.CTkButton(
        root, text="📁 3. Generador fase final",
        width=300, height=40, font=FONT_BUTTON,
        fg_color=ACCENT, hover_color="#3a76c7", text_color="white",
        command=abrir_fase_final
    ).pack(pady=10)

    ctk.CTkLabel(
        root, text="CEID Generator - Menú Principal",
        font=FONT_FOOTER, text_color=GRAY
    ).pack(side="bottom", pady=20)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz_general()