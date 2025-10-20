import customtkinter as ctk
import tkinter as tk
from .gui_constants import *
from .config_manager import ConfigManager

class BaseModal:
    def __init__(self, parent, title, message, modal_type="info"):
        self.result = None
        self.modal_type = modal_type
        
        # Crear ventana modal
        self.window = ctk.CTkToplevel(parent)
        self.window.title(title)
        self.window.geometry("450x250")
        self.window.resizable(False, False)
        
        # Configurar modal
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(fg_color=BG_COLOR)
        
        # Centrar ventana
        self.center_window()
        
        # Crear contenido
        self.create_content(title, message)
        
    def center_window(self):
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.window.winfo_screenheight() // 2) - (250 // 2)
        self.window.geometry(f"450x250+{x}+{y}")
    
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
        message_label = ctk.CTkLabel(main_frame, fg_color=BG_COLOR, text=message, font=("Segoe UI", 12, "bold"), text_color=WHITE_COLOR, wraplength=400)
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
        config = ConfigManager()
        modal_colors = config.get("gui.modal_colors", {})
        
        colors = {
            "info": modal_colors.get("info", "#2196F3"),
            "success": modal_colors.get("success", "#4CAF50"),
            "error": modal_colors.get("error", "#F44336"), 
            "warning": modal_colors.get("warning", "#FF9800"),
            "question": modal_colors.get("question", "#2196F3")
        }
        return colors.get(self.modal_type, "#2196F3")
    
    def get_button_colors(self, button_type):
        config = ConfigManager()
        button_colors = config.get("gui.button_colors", {})
        
        if button_type in button_colors:
            return button_colors[button_type]["bg"], button_colors[button_type]["hover"]
        
        # Valores por defecto
        defaults = {
            "error": ("#F44336", "#D32F2F"),
            "warning": ("#FF9800", "#F57C00"),
            "cancel": ("#757575", "#616161"),
            "confirm": ("#4CAF50", "#45A049")
        }
        return defaults.get(button_type, ("#2196F3", "#1976D2"))
    
    def create_buttons(self, parent):
        pass
    
    def close_modal(self, result=None):
        self.result = result
        self.window.grab_release()
        self.window.destroy()

class InfoModal(BaseModal):
    """Modal de información para mostrar mensajes informativos al usuario."""
    
    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "success")
    
    def create_buttons(self, parent):
        """Crear botón de aceptar para el modal de información."""
        accept_btn = ctk.CTkButton(
            parent,
            text="Aceptar",
            font=FONT_BUTTON,
            text_color=WHITE_COLOR,
            fg_color=BUTTON_BG_COLOR,
            hover_color=BUTTON_HOVER_BG_COLOR,
            command=lambda: self.close_modal(True)
        )
        accept_btn.pack(pady=10)

class ErrorModal(BaseModal):
    """Modal de error para mostrar mensajes de error al usuario."""
    
    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "error")
    
    def create_buttons(self, parent):
        bg_color, hover_color = self.get_button_colors("error")
        accept_btn = ctk.CTkButton(
            parent,
            text="Aceptar",
            font=("Segoe UI", 12, "bold"),
            fg_color=bg_color,
            hover_color=hover_color,
            command=lambda: self.close_modal(True)
        )
        accept_btn.pack(pady=10)

class WarningModal(BaseModal):
    """Modal de advertencia para mostrar mensajes de advertencia al usuario."""
    
    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "warning")
    
    def create_buttons(self, parent):
        bg_color, hover_color = self.get_button_colors("warning")
        accept_btn = ctk.CTkButton(
            parent,
            text="Aceptar", 
            font=("Segoe UI", 12, "bold"),
            fg_color=bg_color,
            hover_color=hover_color,
            command=lambda: self.close_modal(True)
        )
        accept_btn.pack(pady=10)

class ConfirmModal(BaseModal):
    """Modal de confirmación para solicitar confirmación del usuario (Sí/No)."""
    
    def __init__(self, parent, title, message):
        super().__init__(parent, title, message, "question")
    
    def create_buttons(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(pady=10)
        
        cancel_bg, cancel_hover = self.get_button_colors("cancel")
        no_btn = ctk.CTkButton(
            button_frame,
            text="No",
            font=("Segoe UI", 12, "bold"),
            fg_color=cancel_bg,
            hover_color=cancel_hover,
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

# Funciones wrapper para mantener compatibilidad
def showinfo(title, message, parent=None):
    if parent is None:
        parent = tk._default_root
    modal = InfoModal(parent, title, message)
    parent.wait_window(modal.window)
    return modal.result

def showerror(title, message, parent=None):
    if parent is None:
        parent = tk._default_root
    modal = ErrorModal(parent, title, message)
    parent.wait_window(modal.window)
    return modal.result

def showwarning(title, message, parent=None):
    if parent is None:
        parent = tk._default_root
    modal = WarningModal(parent, title, message)
    parent.wait_window(modal.window)
    return modal.result

def askyesno(title, message, parent=None):
    if parent is None:
        parent = tk._default_root
    modal = ConfirmModal(parent, title, message)
    parent.wait_window(modal.window)
    return modal.result