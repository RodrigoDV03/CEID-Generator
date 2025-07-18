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

    # --- PALETA DE COLORES Y FUENTES ---
    PRIMARY = "#2d415a"
    ACCENT = "#4a90e2"
    BG = "#e1e1d3"
    CARD = "#f2f2e4"
    GRAY = "#a0a8b8"
    HOVER = "#3a76c7"


    root = ctk.CTk()
    root.title("CEID - Sistema de Generación de Documentos")
    root.geometry("700x600")
    root.configure(fg_color=BG)
    root.resizable(False, False)
    root.after(100, lambda: root.state("zoomed"))


    FONT_TITLE = ctk.CTkFont("Segoe UI", 28, "bold")
    FONT_BUTTON = ctk.CTkFont("Segoe UI", 16, "bold")
    FONT_FOOTER = ctk.CTkFont("Segoe UI", weight="normal", slant="italic")

    # --- TÍTULO ---
    ctk.CTkLabel(
        root, text="📚 CEID - Generador de Documentos",
        font=FONT_TITLE, text_color=PRIMARY
    ).pack(pady=(50, 20))

    # --- CONTENEDOR PRINCIPAL (FRAME) ---
    main_frame = ctk.CTkFrame(root, fg_color=BG)
    main_frame.pack(pady=10)

    def crear_card(texto, comando):
        card = ctk.CTkFrame(main_frame, fg_color=CARD, corner_radius=15)
        card.pack(pady=15, padx=20, fill="x", expand=True)

        boton = ctk.CTkButton(
            card, text=texto,
            font=FONT_BUTTON, height=50,
            fg_color=ACCENT, hover_color=HOVER,
            text_color="white", corner_radius=10,
            command=comando
        )
        boton.pack(padx=30, pady=20)

    crear_card("📋 1. Generar Planilla", abrir_planilla)
    crear_card("📄 2. Generador Fase Inicial", abrir_fase_inicial)
    crear_card("📁 3. Generador Fase Final", abrir_fase_final)

    # --- FOOTER ---
    ctk.CTkLabel(
        root, text="CEID Generator - Menú Principal",
        font=FONT_FOOTER, text_color=GRAY
    ).pack(side="bottom", pady=20)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz_general()