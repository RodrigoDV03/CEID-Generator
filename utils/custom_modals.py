import customtkinter as ctk
import tkinter as tk
from .gui_constants import *

class BaseModal:
    accept_button_text = "Aceptar"
    accept_button_font = FONT_BUTTON
    accept_button_fg_color = BUTTON_BG_COLOR
    accept_button_hover_color = BUTTON_HOVER_BG_COLOR

    def __init__(self, parent, title, message, modal_type="info"):
        self.result = None
        self.modal_type = modal_type
        self._base_width = 560
        self._base_height = 290
        self._modal_height = self._calcular_altura_modal(message)
        
        # Crear ventana modal
        self.window = ctk.CTkToplevel(parent)
        self.window.title(title)
        self.window.geometry(f"{self._base_width}x{self._modal_height}")
        self.window.resizable(False, False)
        
        # Configurar modal
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(fg_color=BG_COLOR)
        
        # Centrar ventana
        self.center_window()
        
        # Crear contenido
        self.create_content(title, message)

    def _calcular_altura_modal(self, message):
        lineas_estimadas = 0
        for linea in message.splitlines() or [message]:
            ancho_estandar = max(1, len(linea) // 58)
            lineas_estimadas += 1 + ancho_estandar

        extra = max(0, lineas_estimadas - 6) * 14
        return min(self._base_height + extra, 520)
        
    def center_window(self):
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self._base_width // 2)
        y = (self.window.winfo_screenheight() // 2) - (self._modal_height // 2)
        self.window.geometry(f"{self._base_width}x{self._modal_height}+{x}+{y}")
    
    def create_content(self, title, message):
        # Frame principal
        main_frame = ctk.CTkFrame(self.window, fg_color=BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Área del ícono y título
        header_frame = ctk.CTkFrame(main_frame, fg_color=BG_COLOR)
        header_frame.pack(fill="x", pady=(0, 15))
        
        # Ícono según tipo
        icon_text = self.get_icon()
        icon_color = self.get_color()

        icon_label = ctk.CTkLabel(header_frame, fg_color=BG_COLOR, text=icon_text, font=("Segoe UI", 32), text_color=icon_color)
        icon_label.pack(pady=(0, 10))
        
        # Título
        title_label = ctk.CTkLabel(header_frame, fg_color=BG_COLOR, text=title, font=FONT_SECTION, text_color=WHITE_COLOR)
        title_label.pack()
        
        # Mensaje
        message_label = ctk.CTkLabel(
            main_frame,
            fg_color=BG_COLOR,
            text=message,
            font=("Segoe UI", 13),
            text_color=WHITE_COLOR,
            wraplength=500,
            justify="left"
        )
        message_label.pack(pady=(0, 20))
        
        # Frame de botones
        button_frame = ctk.CTkFrame(main_frame, fg_color=BG_COLOR)
        button_frame.pack(fill="x")
        
        self.create_buttons(button_frame)
    
    def get_icon(self):
        icons = {
            "info": "ℹ️",
            "success": "✅", 
            "error": "❌",
            "warning": "⚠️",
            "question": "❓"
        }
        return icons.get(self.modal_type, "ℹ️")
    
    def get_color(self):
        colors = {
            "info": "#2196F3",
            "success": "#4CAF50",
            "error": "#F44336", 
            "warning": "#FF9800",
            "question": "#2196F3"
        }
        return colors.get(self.modal_type, "#2196F3")
    
    def create_buttons(self, parent):
        accept_btn = ctk.CTkButton(
            parent,
            text=self.accept_button_text,
            font=self.accept_button_font,
            text_color=WHITE_COLOR,
            fg_color=self.accept_button_fg_color,
            hover_color=self.accept_button_hover_color,
            command=lambda: self.close_modal(True)
        )
        accept_btn.pack(pady=10)
    
    def close_modal(self, result=None):
        self.result = result
        self.window.grab_release()
        self.window.destroy()

class InfoModal(BaseModal):
    accept_button_fg_color = BUTTON_BG_COLOR
    accept_button_hover_color = BUTTON_HOVER_BG_COLOR

    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "success")

class ErrorModal(BaseModal):
    accept_button_fg_color = "#F44336"
    accept_button_hover_color = "#D32F2F"

    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "error")

class WarningModal(BaseModal):
    accept_button_fg_color = "#FF9800"
    accept_button_hover_color = "#F57C00"

    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "warning")

class ConfirmModal(BaseModal):
    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "question")
    
    def create_buttons(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(pady=10)
        
        no_btn = ctk.CTkButton(
            button_frame,
            text="No",
            font=("Segoe UI", 12, "bold"),
            fg_color="#757575",
            hover_color="#616161",
            width=100,
            command=lambda: self.close_modal(False)
        )
        no_btn.pack(side="left", padx=(0, 10))
        
        yes_btn = ctk.CTkButton(
            button_frame,
            text="Sí", 
            font=("Segoe UI", 12, "bold"),
            fg_color=BUTTON_BG_COLOR,
            hover_color=BUTTON_HOVER_BG_COLOR,
            width=100,
            command=lambda: self.close_modal(True)
        )
        yes_btn.pack(side="left")

def _mostrar_modal(modal_cls, title, message, parent=None):
    if parent is None:
        parent = tk._default_root
    modal = modal_cls(parent, title, message)
    parent.wait_window(modal.window)
    return modal.result


def showinfo(title, message, parent=None):
    return _mostrar_modal(InfoModal, title, message, parent)


def showerror(title, message, parent=None):
    return _mostrar_modal(ErrorModal, title, message, parent)


def showwarning(title, message, parent=None):
    return _mostrar_modal(WarningModal, title, message, parent)


def askyesno(title, message, parent=None):
    return _mostrar_modal(ConfirmModal, title, message, parent)