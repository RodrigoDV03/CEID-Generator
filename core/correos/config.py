import os
import sys
import datetime
from enum import Enum
from dataclasses import dataclass
from typing import List


# ===== CONSTANTES GLOBALES =====

AÑO_ACTUAL = datetime.datetime.now().year


# ===== CONFIGURACIÓN DE GMAIL API =====

def get_app_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


@dataclass
class GmailConfig:
    
    scopes: List[str]
    credentials_file: str
    token_file: str
    remitente: str
    
    @classmethod
    def default(cls) -> 'GmailConfig':
        """Crea una configuración por defecto."""
        app_dir = get_app_dir()
        return cls(
            scopes=[
                'https://www.googleapis.com/auth/gmail.send',
                'https://www.googleapis.com/auth/gmail.compose',
                'https://www.googleapis.com/auth/gmail.settings.basic'
            ],
            credentials_file=os.path.join(app_dir, "credentials.json"),
            token_file=os.path.join(app_dir, "token.pickle"),
            remitente="procesosadministrativosceid.flch.flch@unmsm.edu.pe"
        )


# ===== ENUMS =====

class TipoCorreo(Enum):
    DOCENTE = "docente"
    ADMINISTRATIVO = "administrativo"


# ===== CONFIGURACIÓN DE PATRONES REGEX =====

class PatronesRegex:
    """Patrones regex para extracción de datos de PDFs."""
    
    # Patrones para servicios
    SERVICIO_HORAS = r"^\d{1,2}\s+horas\s+de\s+.+"
    SERVICIO_GENERICO = r"^servicio\s+de\s+.+"


# ===== CONFIGURACIÓN DE EXTRACCIÓN DE SERVICIOS =====

class ServiciosConfig:
    
    # Marcadores de inicio
    MARCADOR_INICIO = "Código Unid. Med."
    MARCADOR_DESCRIPCION = "Descripción"
    
    # Códigos y palabras clave
    CODIGO_SERVICIO = "071100"
    TEXTO_SERVICIO = "SERVICIO DE DICTADO DE CURSO"
    
    # Palabras que indican fin de servicios
    PALABRAS_FIN = ["BAJO LA MODALIDAD", "OFICIO", "***"]


# ===== CONFIGURACIÓN DE CORREOS =====

class EmailConfig:
    
    # Estilo HTML
    FUENTE = "Verdana"
    TAMAÑO_FUENTE = "10pt"
    COLOR_DESTACADO = "#073763"
    
    # Plazos
    DIAS_VENCIMIENTO = 40
    
    # Modalidad
    MODALIDAD = "HÍBRIDA"
    
    # Formato de archivos
    FORMATO_RECIBO = "PDF"
