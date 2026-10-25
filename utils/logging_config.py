"""
Configuración central de logging de la aplicación.

Antes de esto, módulos como `core/correos/gmail_service.py` y
`core/correos/email_sender.py` ya llamaban a `logger.info/warning/error` de
forma consistente, pero no existía ningún `logging.basicConfig` en todo el
proyecto: sin un handler configurado, `logger.info` no se ve en ningún lado y
`logger.warning`/`error` solo llegan al "handler de último recurso" de
Python, que escribe a `sys.stderr` — inexistente en un ejecutable de
PyInstaller empaquetado en modo ventana. En la práctica, esos mensajes
(incluyendo errores de autenticación con Gmail) se perdían siempre.

`configurar_logging()` debe llamarse una única vez, al arrancar `app.py`,
antes de crear la interfaz.
"""

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

NOMBRE_ARCHIVO_LOG = "ceid.log"


def _directorio_base() -> str:
    if getattr(sys, "frozen", False):
        # Empaquetado con PyInstaller: junto al ejecutable, no en el
        # directorio temporal de extracción (sys._MEIPASS), que no es
        # persistente ni escribible de forma confiable.
        return os.path.dirname(sys.executable)
    return os.path.abspath(".")


def _directorio_logs() -> str:
    ruta = os.path.join(_directorio_base(), "logs")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def configurar_logging(nivel: int = logging.INFO) -> str:
    """
    Configura el logging de toda la aplicación con un archivo rotativo en
    `logs/ceid.log` (junto al ejecutable o al proyecto, según corresponda).

    Devuelve la ruta completa del archivo de log, por si se necesita
    mostrarla en la interfaz (por ejemplo, en un modal de error).
    """
    ruta_archivo = os.path.join(_directorio_logs(), NOMBRE_ARCHIVO_LOG)

    formato = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    handler_archivo = RotatingFileHandler(
        ruta_archivo, maxBytes=2_000_000, backupCount=5, encoding="utf-8"
    )
    handler_archivo.setFormatter(formato)
    handlers = [handler_archivo]

    # En desarrollo (python app.py, con consola real) también mostrar
    # WARNING/ERROR por stderr, usando la referencia original (sys.__stderr__)
    # para no verse afectado por una eventual captura temporal de stdout/stderr
    # (ver ui/console_stream.py). En el .exe empaquetado sin consola,
    # sys.__stderr__ es None y este handler simplemente se omite.
    if sys.__stderr__ is not None:
        handler_consola = logging.StreamHandler(sys.__stderr__)
        handler_consola.setFormatter(formato)
        handler_consola.setLevel(logging.WARNING)
        handlers.append(handler_consola) # type: ignore

    logging.basicConfig(level=nivel, handlers=handlers, force=True)

    logging.getLogger(__name__).info("Logging configurado. Archivo: %s", ruta_archivo)
    return ruta_archivo
