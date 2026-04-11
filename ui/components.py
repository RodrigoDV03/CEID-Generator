import customtkinter as ctk
from utils.gui_constants import *

def _crear_selector_archivo(
    master,
    texto,
    comando,
    *,
    frame_fg_color=CARD_COLOR,
    frame_corner_radius=15,
    frame_pack_kwargs=None,
    label_text_color=TEXT_COLOR,
    label_pack_kwargs=None,
    button_text="Seleccionar archivo",
    button_fg_color=BUTTON_BG_COLOR,
    button_hover_color=BUTTON_HOVER_BG_COLOR,
    button_text_color=WHITE_COLOR,
    button_font=FONT_BUTTON,
    button_width=160,
    button_pack_kwargs=None,
):
    frame = ctk.CTkFrame(master, fg_color=frame_fg_color, corner_radius=frame_corner_radius)
    frame.pack(**(frame_pack_kwargs or {"fill": "x", "pady": 10}))

    label = ctk.CTkLabel(frame, text=texto, text_color=label_text_color)
    label.pack(**(label_pack_kwargs or {"side": "left", "padx": 15, "pady": 15}))

    btn = ctk.CTkButton(
        frame,
        text=button_text,
        fg_color=button_fg_color,
        hover_color=button_hover_color,
        text_color=button_text_color,
        font=button_font,
        width=button_width,
        command=comando,
    )
    btn.pack(**(button_pack_kwargs or {"side": "right", "padx": 15}))

    return btn, label


def input_archivo(master, texto, comando, **kwargs):
    boton, _ = _crear_selector_archivo(
        master,
        texto,
        comando,
        button_text="Seleccionar",
        label_text_color=TEXT_COLOR,
        button_fg_color=PRIMARY_COLOR,
        button_hover_color=ACCENT_COLOR,
        button_width=kwargs.pop("button_width", 160),
        button_font=kwargs.pop("button_font", FONT_BUTTON),
        frame_fg_color=kwargs.pop("frame_fg_color", CARD_COLOR),
        frame_corner_radius=kwargs.pop("frame_corner_radius", 15),
        frame_pack_kwargs=kwargs.pop("frame_pack_kwargs", {"fill": "x", "pady": 10}),
        label_pack_kwargs=kwargs.pop("label_pack_kwargs", {"side": "left", "padx": 15, "pady": 15}),
        button_pack_kwargs=kwargs.pop("button_pack_kwargs", {"side": "right", "padx": 15}),
        **kwargs,
    )
    return boton

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


def crear_boton_archivo(frame, texto_etiqueta, comando, **kwargs):
    return _crear_selector_archivo(
        frame,
        texto_etiqueta,
        comando,
        frame_fg_color=kwargs.pop("frame_fg_color", "transparent"),
        frame_corner_radius=kwargs.pop("frame_corner_radius", 0),
        frame_pack_kwargs=kwargs.pop("frame_pack_kwargs", {"fill": "x", "padx": 20, "pady": (0, 10)}),
        label_text_color=kwargs.pop("label_text_color", WHITE_COLOR),
        label_pack_kwargs=kwargs.pop("label_pack_kwargs", {"side": "left", "padx": (0, 10)}),
        button_text=kwargs.pop("button_text", "Seleccionar archivo"),
        button_fg_color=kwargs.pop("button_fg_color", BUTTON_BG_COLOR),
        button_hover_color=kwargs.pop("button_hover_color", BUTTON_HOVER_BG_COLOR),
        button_text_color=kwargs.pop("button_text_color", WHITE_COLOR),
        button_font=kwargs.pop("button_font", FONT_BUTTON),
        button_width=kwargs.pop("button_width", 160),
        button_pack_kwargs=kwargs.pop("button_pack_kwargs", {"side": "right", "padx": 10}),
        **kwargs,
    )


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