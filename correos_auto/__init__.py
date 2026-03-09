"""
Módulo de envío automático de correos con Gmail API.

Este módulo maneja:
- Extracción de RUC y servicios desde PDFs de órdenes de servicio
- Emparejamiento con base de datos Excel mediante RUC
- Construcción de correos HTML personalizados
- Envío automático mediante Gmail API con firma
"""

from .envio_correos import *

__version__ = "2.0.0"
__author__ = "CEID - FLCH - UNMSM"
