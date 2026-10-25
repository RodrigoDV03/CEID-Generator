"""
Consola en vivo para las pantallas de la aplicación.

Reemplaza el patrón anterior de reasignar `sys.stdout` / `sys.stderr` de forma
global y permanente en cada pantalla (ver auditoría: eso afecta a todo el
proceso, nunca se restaura al salir de la pantalla, y permite que un hilo de
fondo escriba directo sobre un widget de Tkinter, lo cual no es seguro).

En su lugar, `ConsolaEnVivo.ejecutar(...)` captura stdout/stderr de forma
*acotada*: solo mientras corre la tarea, con `contextlib.redirect_stdout`
(que garantiza restaurar el valor anterior aunque la tarea falle), y toda
escritura hacia el widget se agenda con `root.after(...)` para que ocurra en
el hilo principal de Tkinter, nunca desde el hilo de fondo.
"""

import contextlib
import io
import threading


class CallbackStream(io.TextIOBase):
    """Stream de texto que reenvía lo escrito a un callback en vez de a una consola real."""

    def __init__(self, callback):
        super().__init__()
        self._callback = callback

    def write(self, s):
        if not s:
            return 0
        self._callback(s)
        return len(s)

    def flush(self):
        return None


class ConsolaEnVivo:
    """Textbox de consola asociado a una pantalla, con captura de salida acotada por tarea."""

    def __init__(self, textbox, root):
        self._textbox = textbox
        self._root = root

    def limpiar(self):
        if not self._textbox.winfo_exists():
            return
        self._textbox.configure(state="normal")
        self._textbox.delete("1.0", "end")
        self._textbox.configure(state="disabled")

    def _escribir(self, texto):
        def hacerlo():
            if not self._textbox.winfo_exists():
                return
            self._textbox.configure(state="normal")
            self._textbox.insert("end", texto)
            self._textbox.see("end")
            self._textbox.configure(state="disabled")

        # Agendar en el hilo principal: nunca tocar el widget desde el hilo de fondo.
        self._root.after(0, hacerlo)

    def ejecutar(self, tarea, on_exito=None, on_error=None, limpiar_antes=True):
        """
        Corre `tarea` (callable sin argumentos) en un hilo de fondo, capturando
        todo lo que imprima (stdout/stderr) hacia esta consola en vivo.

        Al terminar, agenda en el hilo principal `on_exito(resultado)` si la
        tarea no lanzó excepción, o `on_error(excepcion)` si sí.
        """
        if limpiar_antes:
            self.limpiar()

        stream = CallbackStream(self._escribir)

        def hilo():
            try:
                with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
                    resultado = tarea()
                if on_exito:
                    self._root.after(0, lambda: on_exito(resultado))
            except Exception as e:
                if on_error:
                    self._root.after(0, lambda: on_error(e))

        threading.Thread(target=hilo, daemon=True).start()
