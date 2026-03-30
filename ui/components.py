import customtkinter as ctk
from utils.gui_constants import *

def input_archivo(master, texto, comando):

    frame = ctk.CTkFrame(master, fg_color=CARD_COLOR, corner_radius=15)
    frame.pack(fill="x", pady=10)

    label = ctk.CTkLabel(frame, text=texto, text_color=TEXT_COLOR)
    label.pack(side="left", padx=15, pady=15)

    btn = ctk.CTkButton(
        frame,
        text="Seleccionar",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=comando
    )
    btn.pack(side="right", padx=15)

    return btn

def card_opcion(master, titulo, descripcion, comando):

    card = ctk.CTkFrame(
        master,
        fg_color=CARD_COLOR,
        corner_radius=15
    )
    card.pack(side="left", expand=True, fill="x", padx=15, pady=15)
    card.configure(height=190)
    card.pack_propagate(False)

    t = ctk.CTkLabel(
        card,
        text=titulo,
        font=("Segoe UI", 16, "bold"),
        text_color=TEXT_COLOR
    )
    t.pack(pady=(20, 5))

    d = ctk.CTkLabel(
        card,
        text=descripcion,
        text_color=TEXT_LIGHT,
        wraplength=220,
        justify="center"
    )
    d.pack(pady=(0, 15), padx=10)

    btn = ctk.CTkButton(
        card,
        text="Abrir",
        fg_color=PRIMARY_COLOR,
        hover_color=ACCENT_COLOR,
        command=comando
    )
    btn.pack(pady=(0, 20), padx=30, fill="x")

    return card


def etiqueta(master, texto):
    return ctk.CTkLabel(master, text=texto, font=FONT_SECTION, text_color=WHITE_COLOR).pack(anchor="w", padx=20, pady=(15, 5))


def crear_option_menu(frame, variable, opciones):
    menu = ctk.CTkOptionMenu(
        frame,
        variable=variable,
        values=opciones,
        fg_color=BUTTON_BG_COLOR,
        button_color=BUTTON_HOVER_BG_COLOR,
        button_hover_color=BUTTON_HOVER_BG_COLOR,
        text_color=WHITE_COLOR,
    )
    menu.pack(padx=20, fill="x", pady=(0, 10))
    return menu


def crear_boton_archivo(frame, texto_etiqueta, comando):
    contenedor = ctk.CTkFrame(frame, fg_color="transparent")
    contenedor.pack(fill="x", padx=20, pady=(0, 10))

    etiqueta_archivo = ctk.CTkLabel(contenedor, text=texto_etiqueta, text_color=WHITE_COLOR)
    etiqueta_archivo.pack(side="left", padx=(0, 10))

    boton = ctk.CTkButton(
        contenedor,
        text="Seleccionar archivo",
        command=comando,
        width=160,
        fg_color=BUTTON_BG_COLOR,
        hover_color=BUTTON_HOVER_BG_COLOR,
        text_color=WHITE_COLOR,
        font=FONT_BUTTON,
    )
    boton.pack(side="right", padx=10)

    return boton, etiqueta_archivo


def boton_generador(master, texto, comando):
    boton = ctk.CTkButton(
        master=master,
        text=texto,
        command=comando,
        state="normal",
        height=45,
        fg_color=BUTTON_BG_COLOR,
        hover_color=BUTTON_HOVER_BG_COLOR,
        text_color=WHITE_COLOR,
        font=FONT_BUTTON,
    )
    boton.pack(pady=(15, 20))
    return boton


class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        if self.text_widget.winfo_exists():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", message)
            self.text_widget.see("end")
            self.text_widget.configure(state="disabled")

    def flush(self):
        return None