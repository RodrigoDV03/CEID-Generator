from datetime import datetime
import customtkinter as ctk

PRIMARY_COLOR = "#2d415a"
ACCENT_COLOR = "#4a90e2"
BG_COLOR = "#f4f6fa"
TEXT_COLOR = "#333333"
DISABLED_COLOR = "#bbbbbb"
CARD_COLOR = "#f2f2e4"
HOVER_COLOR = "#3a76c7"

BUTTON_BG_COLOR = "#f87171"
BUTTON_HOVER_BG_COLOR = "#f43f5e"


WHITE_COLOR = "#ffffff"
GRAY_COLOR = "#a0a8b8"
BLACK_COLOR = "#000000"

FONT_TITLE = ("Segoe UI", 26, "bold")
FONT_SECTION = ("Segoe UI", 17, "bold")
FONT_TEXT = ("Segoe UI", 12)
FONT_LABEL = ("Segoe UI", 11)
FONT_BUTTON = ("Segoe UI", 10, "bold")
FONT_FOOTER = ("Segoe UI", 9, "italic")

meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

años = [str(a) for a in range(datetime.now().year - 5, datetime.now().year + 6)]

def titulo(master, texto):
    return ctk.CTkLabel(master=master, text=texto, font=FONT_TITLE, text_color=PRIMARY_COLOR, fg_color=WHITE_COLOR)

def etiqueta(master, texto):
    return ctk.CTkLabel(master=master, text=texto, font=FONT_TEXT, text_color=PRIMARY_COLOR, fg_color=WHITE_COLOR)

def crear_option_menu(frame, variable, opciones):
    menu = ctk.CTkOptionMenu(frame, variable=variable, values=opciones)
    menu.pack(padx=10, fill="x", pady=(0, 7))
    return menu

def crear_boton_archivo(frame, etiqueta_archivo, comando):
    boton = ctk.CTkButton(frame, text="Seleccionar archivo", command=comando, width=160)
    boton.pack(side="right", padx=10, pady=(0, 7))
    etiqueta_archivo.pack(side="left", padx=10, pady=(0, 7))
    return boton