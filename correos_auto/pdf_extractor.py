import re
import logging
from typing import Optional, List
from PyPDF2 import PdfReader

from .config import PatronesRegex, ServiciosConfig


logger = logging.getLogger(__name__)


class RUCExtractor:
    """Extractor de RUC desde PDFs."""
    
    def __init__(self):
        """Inicializa el extractor de RUC."""
        # Patrón para RUC: puede tener saltos de línea después de "RUC:"
        self.patron_ruc = re.compile(r'\bRUC:\s*[\n\r]*\s*(\d{11})\b', re.IGNORECASE)
        # Patrón para RUC pegado al nombre (formato común en el PDF)
        self.patron_ruc_nombre = re.compile(r'\b(\d{11})([A-ZÁÉÍÓÚÑ\s]+,\s*[A-ZÁÉÍÓÚÑ\s]+)', re.IGNORECASE)
    
    def extraer(self, texto: str, debug: bool = False) -> Optional[str]:
        """
        Extrae el RUC del texto del PDF.
        
        Args:
            texto: Texto extraído del PDF
            debug: Si True, imprime información de depuración
        
        Returns:
            RUC de 11 dígitos o None si no se encuentra
        """
        # Intentar primero con patrón "RUC:XXXXXXXXXX" (puede tener saltos de línea)
        match = self.patron_ruc.search(texto)
        if match:
            ruc = match.group(1)
            if debug:
                logger.info(f"RUC encontrado (patrón RUC:): {ruc}")
            return ruc
        
        # Buscar patrón RUC pegado a nombre (11 dígitos seguidos de nombre con formato)
        match = self.patron_ruc_nombre.search(texto)
        if match:
            ruc = match.group(1)
            nombre = match.group(2).strip()
            if debug:
                logger.info(f"RUC encontrado (patrón RUC+nombre): {ruc} - {nombre}")
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


class PDFExtractor:
    """Extractor de datos desde PDFs de órdenes de servicio."""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._texto: Optional[str] = None
        self.servicio_extractor = ServicioExtractor()
        self.ruc_extractor = RUCExtractor()
    
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


# Funciones de compatibilidad con código legacy
def extraer_ruc(pdf_path: str, debug: bool = False) -> Optional[str]:
    extractor = PDFExtractor(pdf_path)
    return extractor.extraer_ruc(debug)


def extraer_servicios(pdf_path: str, debug: bool = False) -> Optional[str]:
    extractor = PDFExtractor(pdf_path)
    return extractor.extraer_servicios(debug)
