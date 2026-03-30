import customtkinter as ctk
from utils.gui_constants import *


def titulo(master, texto):
    return ctk.CTkLabel(master, text=texto, font=FONT_TITLE, text_color=WHITE_COLOR).pack(pady=(25, 15))


def boton_volver(master, callback_volver=None):
    def volver():
        if master.winfo_exists():
            master.destroy()
        if callback_volver:
            callback_volver()

    return ctk.CTkButton(
        master,
        text="Volver al menu",
        command=volver,
        fg_color=BUTTON_BG_COLOR,
        hover_color=BUTTON_HOVER_BG_COLOR,
        text_color=WHITE_COLOR,
        font=FONT_BUTTON,
    ).pack(pady=(0, 10))


def footer(master):
    return ctk.CTkLabel(master, text="CEID Generator - FLCH - UNMSM", font=FONT_FOOTER, text_color=WHITE_COLOR).pack(pady=(0, 10))


class AppLayout:

    def __init__(self, titulo="Sistema CEID"):
        self.root = ctk.CTk()
        self.root.title(titulo)
        # Apertura maximizada consistente en Windows.
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0")
        self.root.state("zoomed")
        self.root.minsize(1280, 720)
        self.root.configure(fg_color=BG_COLOR)
        self.botones_sidebar = []

        self.crear_sidebar()
        self.crear_header()
        self.crear_contenido()

    # ---------------- SIDEBAR ----------------
    def crear_sidebar(self):
        self.sidebar = ctk.CTkFrame(
            self.root,
            width=220,
            fg_color=PRIMARY_COLOR,
            corner_radius=0
        )
        self.sidebar.pack(side="left", fill="y")

        titulo = ctk.CTkLabel(
            self.sidebar,
            text="CEID",
            font=("Segoe UI", 20, "bold"),
            text_color="white"
        )
        titulo.pack(pady=30)

    def agregar_opcion_sidebar(self, texto, comando):
        btn = ctk.CTkButton(
            self.sidebar,
            text=texto,
            fg_color="transparent",
            hover_color=ACCENT_COLOR,
            text_color="white",
            anchor="w",
            command=comando
        )
        btn.pack(fill="x", padx=10, pady=5)

    # ---------------- HEADER ----------------
    def crear_header(self):
        self.header = ctk.CTkFrame(
            self.root,
            height=60,
            fg_color="white"
        )
        self.header.pack(fill="x")

        self.titulo_header = ctk.CTkLabel(
            self.header,
            text="Panel",
            font=("Segoe UI", 18, "bold"),
            text_color=TEXT_COLOR
        )
        self.titulo_header.pack(side="left", padx=20)

    # ---------------- CONTENIDO ----------------
    def crear_contenido(self):
        self.main = ctk.CTkFrame(
            self.root,
            fg_color="transparent"
        )
        self.main.pack(expand=True, fill="both", padx=20, pady=20)

    def limpiar_contenido(self):
        for widget in self.main.winfo_children():
            widget.destroy()

    def marcar_activo(self, boton):
        for b in self.botones_sidebar:
            b.configure(fg_color="transparent")

        boton.configure(fg_color=ACCENT_COLOR)

    def ejecutar(self):
        self.root.mainloop()