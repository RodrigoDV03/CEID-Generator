"""Modal para editar datos de correos extraídos de PDFs."""

import customtkinter as ctk
from typing import List, Dict, Callable
from utils.gui_constants import *


class EditarCorreosModal:
    """Modal para revisar y editar servicio y modalidad de correos extraídos."""
    
    def __init__(self, parent, data_envio: List[Dict], on_save: Callable[[List[Dict]], None]):
        """
        Inicializa el modal de edición.
        
        Args:
            parent: Ventana padre
            data_envio: Lista de diccionarios con datos de correos
            on_save: Callback cuando se guardan cambios
        """
        self.parent = parent
        self.data_original = [dict(d) for d in data_envio]  # Copia profunda
        self.data_editada = [dict(d) for d in data_envio]   # Copia para editar
        self.on_save = on_save
        self.resultado = None
        
        # Crear ventana modal
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Editar datos de correos")
        self.window.geometry("1200x600")
        self.window.resizable(True, True)
        
        # Configurar modal
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(fg_color=BG_COLOR)
        
        # Crear interfaz
        self._crear_ui()
    
    def _crear_ui(self):
        """Crea la interfaz del modal."""
        
        # ===== HEADER =====
        header_frame = ctk.CTkFrame(self.window, fg_color=CARD_COLOR, corner_radius=0)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        ctk.CTkLabel(
            header_frame,
            text="Revisar y editar datos extraídos de PDFs",
            font=("Helvetica", 14, "bold"),
            text_color=TEXT_COLOR
        ).pack(anchor="w", padx=20, pady=(15, 10))
        
        ctk.CTkLabel(
            header_frame,
            text=f"Total: {len(self.data_editada)} correos",
            font=("Helvetica", 10),
            text_color=TEXT_LIGHT
        ).pack(anchor="w", padx=20, pady=(0, 15))
        
        # ===== TABLA CON SCROLL =====
        tabla_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        tabla_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Crear tabla con scroll
        self.tabla_scroll = ctk.CTkScrollableFrame(tabla_frame, fg_color="transparent")
        self.tabla_scroll.pack(fill="both", expand=True)
        
        # Headers
        header_datos = ctk.CTkFrame(self.tabla_scroll, fg_color=PRIMARY_COLOR, corner_radius=5)
        header_datos.pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(header_datos, text="Nombre", font=("Helvetica", 10, "bold"), text_color="white", width=150).pack(side="left", padx=10, pady=8)
        ctk.CTkLabel(header_datos, text="Correo", font=("Helvetica", 10, "bold"), text_color="white", width=180).pack(side="left", padx=10, pady=8)
        ctk.CTkLabel(header_datos, text="RUC", font=("Helvetica", 10, "bold"), text_color="white", width=100).pack(side="left", padx=10, pady=8)
        ctk.CTkLabel(header_datos, text="Servicio", font=("Helvetica", 10, "bold"), text_color="white", width=200).pack(side="left", padx=10, pady=8)
        ctk.CTkLabel(header_datos, text="Modalidad", font=("Helvetica", 10, "bold"), text_color="white", width=100).pack(side="left", padx=10, pady=8)
        ctk.CTkLabel(header_datos, text="Acción", font=("Helvetica", 10, "bold"), text_color="white", width=100).pack(side="left", padx=10, pady=8)
        
        # Filas editables
        self.campos_editables = []
        for idx, datos in enumerate(self.data_editada):
            self._crear_fila_editable(idx, datos)
        
        # ===== FOOTER =====
        footer_frame = ctk.CTkFrame(self.window, fg_color=CARD_COLOR, corner_radius=0)
        footer_frame.pack(fill="x", padx=0, pady=0)
        
        btn_frame = ctk.CTkFrame(footer_frame, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        ctk.CTkButton(
            btn_frame,
            text="Cancelar",
            width=120,
            fg_color=BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=TEXT_LIGHT,
            border_width=1,
            hover_color="#444444",
            command=self.window.destroy
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_frame,
            text="Guardar cambios",
            width=120,
            fg_color=ACCENT_COLOR,
            hover_color=PRIMARY_COLOR,
            command=self._guardar
        ).pack(side="left", padx=10)
    
    def _crear_fila_editable(self, idx: int, datos: Dict):
        """Crea una fila editable en la tabla."""
        
        fila_frame = ctk.CTkFrame(self.tabla_scroll, fg_color=BG_COLOR, corner_radius=5)
        fila_frame.pack(fill="x", pady=3)
        
        # Nombre (Read-only)
        ctk.CTkLabel(
            fila_frame,
            text=datos.get("nombre", "N/A"),
            text_color=TEXT_COLOR,
            width=150,
            anchor="w",
            justify="left"
        ).pack(side="left", padx=10, pady=8)
        
        # Correo (Read-only)
        ctk.CTkLabel(
            fila_frame,
            text=datos.get("correo", "N/A"),
            text_color=TEXT_LIGHT,
            width=180,
            anchor="w",
            justify="left"
        ).pack(side="left", padx=10, pady=8)
        
        # RUC (Read-only)
        ctk.CTkLabel(
            fila_frame,
            text=datos.get("ruc", "N/A"),
            text_color=TEXT_LIGHT,
            width=100,
            anchor="w",
            justify="left"
        ).pack(side="left", padx=10, pady=8)
        
        # Servicio (Editable)
        servicio_var = ctk.StringVar(value=datos.get("servicio", ""))
        entry_servicio = ctk.CTkEntry(
            fila_frame,
            textvariable=servicio_var,
            width=200,
            fg_color=CARD_COLOR,
            text_color="black",
            border_color=PRIMARY_COLOR,
            border_width=1
        )
        entry_servicio.pack(side="left", padx=10, pady=8)
        
        # Modalidad (Editable)
        modalidad_var = ctk.StringVar(value=datos.get("modalidad", ""))
        entry_modalidad = ctk.CTkEntry(
            fila_frame,
            textvariable=modalidad_var,
            width=100,
            fg_color=CARD_COLOR,
            text_color="black",
            border_color=PRIMARY_COLOR,
            border_width=1
        )
        entry_modalidad.pack(side="left", padx=10, pady=8)
        
        # Botón Restaurar
        def restaurar_original():
            original = self.data_original[idx]
            servicio_var.set(original.get("servicio", ""))
            modalidad_var.set(original.get("modalidad", ""))
        
        ctk.CTkButton(
            fila_frame,
            text="Restaurar",
            width=90,
            fg_color=BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=TEXT_LIGHT,
            border_width=1,
            hover_color="#444444",
            font=("Helvetica", 9),
            command=restaurar_original
        ).pack(side="left", padx=5, pady=8)
        
        # Guardar referencias para obtener los valores después
        self.campos_editables.append({
            "idx": idx,
            "servicio_var": servicio_var,
            "modalidad_var": modalidad_var
        })
    
    def _guardar(self):
        """Guarda los cambios en data_editada."""
        # Actualizar data_editada con los valores de los campos
        for campo in self.campos_editables:
            idx = campo["idx"]
            self.data_editada[idx]["servicio"] = campo["servicio_var"].get()
            self.data_editada[idx]["modalidad"] = campo["modalidad_var"].get()
        
        # Llamar callback
        self.on_save(self.data_editada)
        
        # Cerrar modal
        self.window.destroy()


class EditarCorreoIndividualModal:
    """Modal para editar un correo individual (para modo contrato)."""
    
    def __init__(self, parent, datos: Dict, on_save: Callable[[Dict], None]):
        """
        Inicializa el modal de edición individual.
        
        Args:
            parent: Ventana padre
            datos: Diccionario con datos del correo
            on_save: Callback cuando se guardan cambios
        """
        self.parent = parent
        self.datos_original = dict(datos)
        self.datos_editado = dict(datos)
        self.on_save = on_save
        
        # Crear ventana modal
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Editar correo individual")
        self.window.geometry("600x400")
        self.window.resizable(False, False)
        
        # Configurar modal
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(fg_color=BG_COLOR)
        
        # Centrar ventana
        self.window.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 400) // 2
        self.window.geometry(f"600x400+{x}+{y}")
        
        # Crear interfaz
        self._crear_ui()
    
    def _crear_ui(self):
        """Crea la interfaz del modal."""
        
        # ===== HEADER =====
        header_frame = ctk.CTkFrame(self.window, fg_color=CARD_COLOR, corner_radius=0)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        ctk.CTkLabel(
            header_frame,
            text=f"Editar: {self.datos_original.get('nombre', 'N/A')}",
            font=("Helvetica", 14, "bold"),
            text_color=TEXT_COLOR
        ).pack(anchor="w", padx=20, pady=(15, 10))
        
        ctk.CTkLabel(
            header_frame,
            text=self.datos_original.get("correo", "N/A"),
            font=("Helvetica", 10),
            text_color=TEXT_LIGHT
        ).pack(anchor="w", padx=20, pady=(0, 15))
        
        # ===== CONTENIDO =====
        contenido_frame = ctk.CTkScrollableFrame(self.window, fg_color="transparent")
        contenido_frame.pack(fill="both", expand=True, padx=30, pady=20)
        
        # RUC (Read-only)
        ctk.CTkLabel(contenido_frame, text="RUC", text_color=TEXT_COLOR, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        ctk.CTkEntry(
            contenido_frame,
            border_width=0,
            fg_color=CARD_COLOR,
            state="disabled"
        ).insert(0, self.datos_original.get("ruc", "N/A"))
        entry_ruc = ctk.CTkEntry(contenido_frame, border_width=0, fg_color=CARD_COLOR, state="disabled")
        entry_ruc.insert(0, self.datos_original.get("ruc", "N/A"))
        entry_ruc.pack(fill="x", pady=(0, 20))
        
        # Servicio (Editable)
        ctk.CTkLabel(contenido_frame, text="Servicio", text_color=TEXT_COLOR, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        self.servicio_var = ctk.StringVar(value=self.datos_original.get("servicio", ""))
        entry_servicio = ctk.CTkEntry(
            contenido_frame,
            textvariable=self.servicio_var,
            fg_color=CARD_COLOR,
            text_color="black",
            border_color=PRIMARY_COLOR,
            border_width=2,
            height=50
        )
        entry_servicio.pack(fill="x", pady=(0, 20))
        
        # Modalidad (Editable)
        ctk.CTkLabel(contenido_frame, text="Modalidad", text_color=TEXT_COLOR, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        self.modalidad_var = ctk.StringVar(value=self.datos_original.get("modalidad", ""))
        entry_modalidad = ctk.CTkEntry(
            contenido_frame,
            textvariable=self.modalidad_var,
            fg_color=CARD_COLOR,
            text_color="black",
            border_color=PRIMARY_COLOR,
            border_width=2
        )
        entry_modalidad.pack(fill="x", pady=(0, 30))
        
        # ===== BOTONES =====
        btn_frame = ctk.CTkFrame(self.window, fg_color=CARD_COLOR, corner_radius=0)
        btn_frame.pack(fill="x", padx=0, pady=0)
        
        btn_inner = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_inner.pack(pady=15)
        
        ctk.CTkButton(
            btn_inner,
            text="Cancelar",
            width=120,
            fg_color=BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=TEXT_LIGHT,
            border_width=1,
            hover_color="#444444",
            command=self.window.destroy
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_inner,
            text="Restaurar original",
            width=150,
            fg_color=BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=TEXT_LIGHT,
            border_width=1,
            hover_color="#444444",
            command=self._restaurar
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_inner,
            text="Guardar cambios",
            width=120,
            fg_color=ACCENT_COLOR,
            hover_color=PRIMARY_COLOR,
            command=self._guardar
        ).pack(side="left", padx=10)
    
    def _restaurar(self):
        """Restaura los valores originales."""
        self.servicio_var.set(self.datos_original.get("servicio", ""))
        self.modalidad_var.set(self.datos_original.get("modalidad", ""))
    
    def _guardar(self):
        """Guarda los cambios."""
        self.datos_editado["servicio"] = self.servicio_var.get()
        self.datos_editado["modalidad"] = self.modalidad_var.get()
        self.on_save(self.datos_editado)
        self.window.destroy()
