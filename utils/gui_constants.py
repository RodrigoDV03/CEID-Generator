from datetime import datetime
import customtkinter as ctk
from .config_manager import get_gui_colors, get_gui_fonts

# Cargar configuración desde archivo externo
_colors = get_gui_colors()
_fonts = get_gui_fonts()

# Colores (con fallback a valores originales)
SECTION_COLOR = _colors.get("section_color", "#1d5fb6")
BG_COLOR = _colors.get("bg_color", "#114b98")
BUTTON_BG_COLOR = _colors.get("button_bg_color", "#f38b1a")
BUTTON_HOVER_BG_COLOR = _colors.get("button_hover_bg_color", "#e8aa69")
CONSOLE_BG = _colors.get("console_bg", "#2c2f33")
WHITE_COLOR = _colors.get("white_color", "#ffffff")

# Fuentes (con fallback a valores originales)
FONT_TITLE = tuple(_fonts.get("title", ["Segoe UI", 26, "bold"]))
FONT_SECTION = tuple(_fonts.get("section", ["Segoe UI", 16, "bold"]))
FONT_BUTTON = tuple(_fonts.get("button", ["Segoe UI", 14, "bold"]))
FONT_FOOTER = tuple(_fonts.get("footer", ["Segoe UI", 12]))

meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

años = [str(a) for a in range(datetime.now().year - 5, datetime.now().year + 6)]

def titulo(master, texto):
    return ctk.CTkLabel(master, text=texto, font=FONT_TITLE, text_color=WHITE_COLOR).pack(pady=(25, 15))

def seccion_titulo(master, texto):
    return ctk.CTkLabel(master, text=texto, font=FONT_SECTION, text_color=WHITE_COLOR)

def etiqueta(master, texto):
    return ctk.CTkLabel(master, text=texto, font=FONT_SECTION, text_color=WHITE_COLOR).pack(anchor="w", padx=20, pady=(15, 5))

def crear_option_menu(frame, variable, opciones):
    menu = ctk.CTkOptionMenu(frame, variable=variable, values=opciones, fg_color=BUTTON_BG_COLOR, button_color=BUTTON_HOVER_BG_COLOR, button_hover_color=BUTTON_HOVER_BG_COLOR, text_color=WHITE_COLOR)
    menu.pack(padx=20, fill="x", pady=(0, 10))
    return menu

def crear_boton_archivo(frame, texto_etiqueta, comando):
    contenedor = ctk.CTkFrame(frame, fg_color="transparent")
    contenedor.pack(fill="x", padx=20, pady=(0, 10))

    etiqueta = ctk.CTkLabel(contenedor, text=texto_etiqueta, text_color=WHITE_COLOR)
    etiqueta.pack(side="left", padx=(0, 10))
    
    boton = ctk.CTkButton(contenedor, text="Seleccionar archivo", command=comando, width=160, fg_color=BUTTON_BG_COLOR,
                            hover_color=BUTTON_HOVER_BG_COLOR, text_color=WHITE_COLOR, font=FONT_BUTTON)
    boton.pack(side="right", padx=10)
    
    return boton, etiqueta

def boton_generador(master, texto, comando):
    boton = ctk.CTkButton(master=master, text=texto, command=comando, state="normal", height=45,
                            fg_color=BUTTON_BG_COLOR, hover_color=BUTTON_HOVER_BG_COLOR,
                            text_color=WHITE_COLOR, font=FONT_BUTTON)
    boton.pack(pady=(15, 20))
    return boton

def boton_volver(master, callback_volver=None):
    def volver():
        if master.winfo_exists():
            master.destroy()
        if callback_volver:
            callback_volver()
    return ctk.CTkButton(master, text="⬅ Volver al menú", command=volver, fg_color=BUTTON_BG_COLOR, hover_color=BUTTON_HOVER_BG_COLOR,
                            text_color=WHITE_COLOR, font=FONT_BUTTON).pack(pady=(0, 10))

def footer(master):
    return ctk.CTkLabel(master, text="CEID Generator - FLCH - UNMSM", font=FONT_FOOTER, text_color=WHITE_COLOR).pack(pady=(0, 10))
