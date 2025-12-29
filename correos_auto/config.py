import os
import sys
import datetime
from enum import Enum
from dataclasses import dataclass
from typing import List


# ===== CONSTANTES GLOBALES =====

AÑO_ACTUAL = datetime.datetime.now().year


# ===== PALABRAS PROHIBIDAS =====

PALABRAS_PROHIBIDAS = {
    "SERVICIO", "SERVICIOSERVICIO", "CEID", "FACULTAD", "LETRAS", "CIENCIAS", 
    "HUMANAS", "UNMSM", "BAJO", "MODALIDAD", "COORDINACION", "ATENCION", 
    "DIGITACION", "CLASIFICACION", "DESCRIPCION", "EVALUACION", "SOLICITUDES", 
    "RECEPCION", "ARCHIVO", "VISITAS", "DE", "LA", "Y", "EN", "MES", "DIA", 
    "AÑO", "ADJUDICACION", "PROCESO", "SIN", "ADQUISICION", "ORDEN", "LIMA", 
    "INDUSTRIAL", "INT", "URB", "URBANIZACION", "AVENIDA", "AV", "CALLE", 
    "CAL", "JR", "JIRON", "MZ", "LOTE", "DISTRITO", "PROVINCIA", "DEPARTAMENTO", 
    "CALLAO", "ANCON", "SAN", "JUAN", "MARTIN", "PARQUE", "ALAMEDA", "PSJ", 
    "PASAJE"
}


# ===== CONFIGURACIÓN DE VALIDACIÓN =====

class ValidacionConfig:
    
    # Validación de nombres
    MIN_LETRAS_NOMBRE = 3
    MIN_LONGITUD_NOMBRE = 15
    MAX_LONGITUD_NOMBRE = 60
    MIN_PALABRAS_APELLIDOS = 1
    MIN_PALABRAS_NOMBRES = 1
    
    # Validación de servicios
    MAX_PALABRAS_NOMBRE_SIN_COMA = 6
    MIN_PALABRAS_NOMBRE_COMPLETO = 3
    
    # Matching de nombres
    UMBRAL_FUZZY_MATCHING = 80  # Porcentaje mínimo de similitud


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
            remitente="personalcontratado28.flch@unmsm.edu.pe"
        )


# ===== ENUMS =====

class TipoCorreo(Enum):
    DOCENTE = "docente"
    ADMINISTRATIVO = "administrativo"


class TipoPatron(Enum):
    CONCEPTO_UNMSM_CON_COMA = "Patrón 1 (Concepto:UNMSM con coma)"
    NUMEROS_CON_COMA = "Patrón 2 (números con coma)"
    CONCEPTO_UNMSM_SIN_COMA = "Patrón 1b (Concepto:UNMSM sin coma)"
    NUMEROS_SIN_COMA = "Patrón 2b (números sin coma)"
    DESPUES_RUC = "Patrón 2c (después de RUC)"
    CON_SALTOS = "Patrón 3 (con saltos)"
    FALLBACK = "Patrón 4 (fallback)"


# ===== CONFIGURACIÓN DE PATRONES REGEX =====

class PatronesRegex:
    
    # Patrones para nombres
    CONCEPTO_UNMSM_CON_COMA = r"Concepto:\s*UNMSM\s*\n?\s*\d+\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)"
    CONCEPTO_UNMSM_SIN_COMA = r"Concepto:\s*UNMSM\s*\n?\s*(\d{8,})\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})"
    NUMEROS_CON_COMA = r"\d{8,}\s*([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)"
    NUMEROS_SIN_COMA = r"\d{8,}\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})"
    RUC_SEGUIDO_NOMBRE = r"\bRUC:.*?(\d{11})\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,})"
    NOMBRE_CON_SALTOS = r"\b([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\s*,\s*\n?\s*([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+)*)\b"
    NOMBRE_FALLBACK = r"\b([A-ZÁÉÍÓÚÑ ]+,\s*[A-ZÁÉÍÓÚÑ ]+)\b"
    
    # Patrones para servicios
    SERVICIO_HORAS = r"^\d{1,2}\s+horas\s+de\s+.+"
    SERVICIO_GENERICO = r"^servicio\s+de\s+.+"
    
    # Limpieza
    ESPACIOS_MULTIPLES = r"\s+"


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


# ===== PRIORIDADES DE PATRONES =====

PRIORIDAD_PATRONES = {
    TipoPatron.CONCEPTO_UNMSM_CON_COMA: 1,
    TipoPatron.NUMEROS_CON_COMA: 2,
    TipoPatron.CONCEPTO_UNMSM_SIN_COMA: 3,
    TipoPatron.NUMEROS_SIN_COMA: 4,
    TipoPatron.DESPUES_RUC: 5,
    TipoPatron.CON_SALTOS: 6,
    TipoPatron.FALLBACK: 7
}
