import re
import logging
import unicodedata
from typing import Optional, List
from PyPDF2 import PdfReader

from .config import PatronesRegex, ServiciosConfig


logger = logging.getLogger(__name__)


def _normalizar_ascii(texto: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFD", texto)
        if unicodedata.category(ch) != "Mn"
    )


class RUCExtractor:
    """Extractor de RUC desde PDFs."""
    
    def __init__(self):
        self.patron_ruc_nombre = re.compile(
            r'\b(\d{11})(?=[A-ZÑ\s]{3,},\s*[A-ZÑ\s]+)',
            re.IGNORECASE,
        )

    def _buscar_ruc_cerca_de_marcador(self, texto: str) -> Optional[str]:
        """Busca el primer RUC de 11 dígitos cerca del marcador RUC."""
        for match in re.finditer(r'\bRUC\b', texto, re.IGNORECASE):
            ventana = texto[match.end():match.end() + 200]
            coincidencia = re.search(r'\b(\d{11})\b', ventana)
            if coincidencia:
                return coincidencia.group(1)
        return None
    
    def extraer(self, texto: str, debug: bool = False) -> Optional[str]:
        texto_normalizado = _normalizar_ascii(texto)

        # Priorizar el RUC del proveedor, que suele aparecer junto al nombre.
        match = self.patron_ruc_nombre.search(texto_normalizado)
        if match:
            ruc = match.group(1)
            if debug:
                logger.info(f"RUC encontrado (patrón proveedor+nombre): {ruc}")
            return ruc

        # Intentar después con el RUC cerca del marcador "RUC".
        ruc = self._buscar_ruc_cerca_de_marcador(texto_normalizado)
        if ruc:
            if debug:
                logger.info(f"RUC encontrado (cerca de marcador RUC): {ruc}")
            return ruc
        
        if debug:
            logger.warning("No se encontró RUC en el PDF")
        return None


class ServicioExtractor:
    
    def __init__(self):
        """Inicializa el extractor de servicios."""
        self.patron_horas = re.compile(PatronesRegex.SERVICIO_HORAS, re.IGNORECASE)
        self.patron_servicio = re.compile(PatronesRegex.SERVICIO_GENERICO, re.IGNORECASE)
    
    def extraer(self, texto: str, debug: bool = False) -> Optional[str]:
        lineas = texto.splitlines()
        
        # Buscar la línea que contiene el marcador de inicio
        idx_inicio = self._encontrar_inicio_servicios(lineas)
        
        if idx_inicio is None:
            if debug:
                logger.warning("No se encontró la sección de servicios")
            return None
        
        # Extraer servicios desde esa línea
        servicios = self._extraer_lineas_servicio(lineas, idx_inicio, debug)
        
        if not servicios:
            return None
        
        # Formatear servicios
        servicios_formateados = self._formatear_servicios(servicios)
        
        if debug:
            logger.info(f"Total de servicios encontrados: {len(servicios)}")
            logger.info(f"Texto formateado:\n{servicios_formateados}")
        
        return servicios_formateados
    
    def _encontrar_inicio_servicios(self, lineas: List[str]) -> Optional[int]:
        for i, linea in enumerate(lineas):
            if (ServiciosConfig.MARCADOR_INICIO in linea and 
                ServiciosConfig.MARCADOR_DESCRIPCION in linea):
                return i
        return None
    
    def _extraer_lineas_servicio(
        self, 
        lineas: List[str], 
        idx_inicio: int, 
        debug: bool
    ) -> List[str]:
        servicios = []
        capturando = False
        
        for i in range(idx_inicio + 1, len(lineas)):
            linea = lineas[i].strip()
            
            # Línea vacía, continuar
            if not linea:
                continue
            
            # Si encontramos palabra de fin, detener
            if self._es_fin_servicios(linea):
                break
            
            # Saltar códigos o encabezados
            if linea.startswith(ServiciosConfig.CODIGO_SERVICIO) or \
               ServiciosConfig.TEXTO_SERVICIO in linea.upper():
                capturando = True
                continue
            
            # Capturar líneas de servicio
            if capturando and self._es_linea_servicio(linea):
                servicios.append(linea)
                if debug:
                    logger.info(f"✓ Capturado: {linea}")
        
        return servicios
    
    def _es_fin_servicios(self, linea: str) -> bool:
        return any(palabra in linea for palabra in ServiciosConfig.PALABRAS_FIN)
    
    def _es_linea_servicio(self, linea: str) -> bool:
        return bool(self.patron_horas.match(linea) or self.patron_servicio.match(linea))
    
    def _formatear_servicios(self, servicios: List[str]) -> str:
        if len(servicios) > 1:
            return ", ".join(servicios[:-1]) + " y " + servicios[-1]
        else:
            return servicios[0]


class ModalidadExtractor:

    def __init__(self):
        self.patron_modalidad = re.compile(
            r"BAJO\s+LA\s+MODALIDAD\s*:?\s*([A-ZÁÉÍÓÚÜÑ]+)",
            re.IGNORECASE
        )

    def extraer(self, texto: str, debug: bool = False) -> Optional[str]:
        match = self.patron_modalidad.search(texto)
        if not match:
            if debug:
                logger.warning("No se encontró modalidad en el PDF")
            return None

        modalidad_cruda = match.group(1).strip().upper()
        modalidad_normalizada = _normalizar_ascii(modalidad_cruda)

        if "HIBRIDA" in modalidad_normalizada or "HIBRIDO" in modalidad_normalizada:
            return "HÍBRIDA"
        if "VIRTUAL" in modalidad_normalizada:
            return "VIRTUAL"
        if "PRESENCIAL" in modalidad_normalizada:
            return "PRESENCIAL"

        if debug:
            logger.warning(f"Modalidad detectada pero no reconocida: {modalidad_cruda}")
        return None

class ContratoExtractor:

    def __init__(self):
        self.patron_contrato = re.compile(
            r"CONTRATO\s+DE\s+LOCACION\s+DE\s+SERVICIOS",
            re.IGNORECASE
        )

    def tiene_contrato(self, texto: str) -> bool:
        texto_normalizado = _normalizar_ascii(texto).upper()
        return bool(self.patron_contrato.search(texto_normalizado))


class PDFExtractor:
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._texto: Optional[str] = None
        self.servicio_extractor = ServicioExtractor()
        self.ruc_extractor = RUCExtractor()
        self.modalidad_extractor = ModalidadExtractor()
        self.contrato_extractor = ContratoExtractor()
    
    @property
    def texto(self) -> str:
        if self._texto is None:
            self._texto = self._extraer_texto()
        return self._texto
    
    def _extraer_texto(self) -> str:
        try:
            reader = PdfReader(self.pdf_path)
            texto_completo = ""
            for pagina in reader.pages:
                texto_completo += pagina.extract_text() + "\n"
            return texto_completo
        except Exception as e:
            logger.error(f"Error extrayendo texto de {self.pdf_path}: {e}")
            raise
    
    def extraer_ruc(self, debug: bool = False) -> Optional[str]:
        """Extrae el RUC del PDF."""
        return self.ruc_extractor.extraer(self.texto, debug)
    
    def extraer_servicios(self, debug: bool = False) -> Optional[str]:
        """Extrae los servicios del PDF."""
        return self.servicio_extractor.extraer(self.texto, debug)

    def extraer_modalidad(self, debug: bool = False) -> Optional[str]:
        """Extrae la modalidad desde el PDF."""
        return self.modalidad_extractor.extraer(self.texto, debug)

    def tiene_contrato_locacion(self) -> bool:
        """Valida si la orden incluye la frase de contrato de locacion de servicios."""
        return self.contrato_extractor.tiene_contrato(self.texto)
