"""
Barra de progreso "x / n" reutilizable para los procesos masivos de la
aplicación (planilla, documentos por fase, control de pagos, envío de
correos).

Igual que `ui/console_stream.py`, toda actualización se agenda con
`root.after(...)` para poder llamarse de forma segura desde el hilo de
fondo que hace el trabajo real, sin tocar el widget directamente desde ahí.
"""

import customtkinter as ctk
from utils.gui_constants import TEXT_LIGHT, TEXT_COLOR


class BarraProgreso:
    def __init__(self, master, root):
        self._root = root

        self.frame = ctk.CTkFrame(master, fg_color="transparent")

        self._label = ctk.CTkLabel(self.frame, text="", text_color=TEXT_LIGHT, font=("Segoe UI", 12))
        self._label.pack(anchor="w", pady=(0, 4))

        self._barra = ctk.CTkProgressBar(self.frame)
        self._barra.set(0)
        self._barra.pack(fill="x")

    def iniciar(self, texto="Preparando..."):
        """Reinicia la barra a 0 y muestra un texto inicial. Llamar antes de empezar la tarea."""
        def hacerlo():
            if not self._barra.winfo_exists():
                return
            self._barra.set(0)
            self._label.configure(text=texto, text_color=TEXT_LIGHT)
        self._root.after(0, hacerlo)

    def actualizar(self, actual, total, texto=""):
        """Actualiza la barra y la etiqueta a 'actual / total — texto'."""
        def hacerlo():
            if not self._barra.winfo_exists():
                return
            fraccion = (actual / total) if total else 0
            self._barra.set(min(max(fraccion, 0), 1))
            partes = [f"{actual} / {total}"] if total else []
            if texto:
                partes.append(texto)
            self._label.configure(text=" — ".join(partes), text_color=TEXT_COLOR)
        self._root.after(0, hacerlo)

    def completar(self, texto="Completado"):
        """Deja la barra llena y con un texto final."""
        def hacerlo():
            if not self._barra.winfo_exists():
                return
            self._barra.set(1)
            self._label.configure(text=texto, text_color=TEXT_COLOR)
        self._root.after(0, hacerlo)
