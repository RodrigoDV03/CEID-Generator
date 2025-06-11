import tkinter as tk
from generador_planilla.interfaz_gui import iniciar_interfaz_planilla
from fase_inicial.interfaz_gui import iniciar_interfaz_fase_inicial
from fase_final.interfaz_gui import iniciar_interfaz_fase_final

def iniciar_interfaz_general():

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

    root = tk.Tk()
    root.title("CEID - Sistema de Generación de Documentos")
    root.geometry("500x400")
    root.configure(bg="#f4f6fa")
    root.resizable(False, False)

    PRIMARY = "#2d415a"
    ACCENT = "#4a90e2"
    BG = "#f4f6fa"
    GRAY = "#a0a8b8"
    FONT_TITLE = ("Segoe UI", 18, "bold")
    FONT_BUTTON = ("Segoe UI", 12, "bold")
    FONT_FOOTER = ("Segoe UI", 9)

    tk.Label(root, text="CEID - Sistema de Documentos", font=FONT_TITLE, bg=BG, fg=PRIMARY).pack(pady=(30, 20))

    tk.Button(root, text="📋 1. Generar Planilla", width=30, font=FONT_BUTTON, bg=ACCENT, fg="white", command=abrir_planilla).pack(pady=10)
    tk.Button(root, text="📄 2. Generador fase inicial", width=30, font=FONT_BUTTON, bg=ACCENT, fg="white", command=abrir_fase_inicial).pack(pady=10)
    tk.Button(root, text="📁 3. Generador fase final", width=30, font=FONT_BUTTON, bg=ACCENT, fg="white", command=abrir_fase_final).pack(pady=10)

    tk.Label(root, text="CEID Generator - Menú Principal", font=FONT_FOOTER, bg=BG, fg=GRAY).pack(side="bottom", pady=20)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interfaz_general()