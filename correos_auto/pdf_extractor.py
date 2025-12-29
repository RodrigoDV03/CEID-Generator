import re
import logging
from abc import ABC, abstractmethod
from typing import Optional, List
from PyPDF2 import PdfReader

from .config import (
    PatronesRegex,
    TipoPatron,
    ServiciosConfig,
    ValidacionConfig,
    PRIORIDAD_PATRONES
)
from .validators import (
    NombreValidator,
    FormatoNombreConverter,
    CandidatoNombre
)


logger = logging.getLogger(__name__)


class PatternMatcher(ABC):
    """Clase base abstracta para matchers de patrones de nombres."""
    
    def __init__(self, tipo_patron: TipoPatron):
        """
        Inicializa el matcher.
        
        Args:
            tipo_patron: Tipo de patrón que este matcher busca
        """
        self.tipo_patron = tipo_patron
        self.prioridad = PRIORIDAD_PATRONES[tipo_patron]
    
    @abstractmethod
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        pass
    
    def _crear_candidato(self, nombre: str) -> CandidatoNombre:
        nombre_limpio = NombreValidator.limpiar(nombre)
        return CandidatoNombre(
            nombre=nombre_limpio,
            patron=self.tipo_patron.value,
            prioridad=self.prioridad
        )


class ConceptoUNMSMConComaMatcher(PatternMatcher):
    """Matcher para patrón: Concepto:UNMSM seguido de números y nombre CON COMA."""
    
    def __init__(self):
        super().__init__(TipoPatron.CONCEPTO_UNMSM_CON_COMA)
    
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        match = re.search(PatronesRegex.CONCEPTO_UNMSM_CON_COMA, texto, re.IGNORECASE)
        if match:
            nombre = match.group(1)
            if NombreValidator.es_valido(nombre):
                return self._crear_candidato(nombre)
        return None


class ConceptoUNMSMSinComaMatcher(PatternMatcher):
    """Matcher para patrón: Concepto:UNMSM seguido de números y nombre SIN COMA."""
    
    def __init__(self):
        super().__init__(TipoPatron.CONCEPTO_UNMSM_SIN_COMA)
    
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        match = re.search(PatronesRegex.CONCEPTO_UNMSM_SIN_COMA, texto, re.IGNORECASE)
        if match:
            nombre_sin_coma = match.group(2).strip()
            nombre_formateado = FormatoNombreConverter.convertir_a_formato_coma(nombre_sin_coma)
            if nombre_formateado and NombreValidator.es_valido(nombre_formateado, min_letras=3):
                return self._crear_candidato(nombre_formateado)
        return None


class NumerosConComaMatcher(PatternMatcher):
    """Matcher para patrón: números largos seguidos directamente de nombre CON COMA."""
    
    def __init__(self):
        super().__init__(TipoPatron.NUMEROS_CON_COMA)
    
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        match = re.search(PatronesRegex.NUMEROS_CON_COMA, texto)
        if match:
            nombre = match.group(1)
            if NombreValidator.es_valido(nombre):
                return self._crear_candidato(nombre)
        return None


class NumerosSinComaMatcher(PatternMatcher):
    """Matcher para patrón: números largos seguidos de nombre SIN COMA."""
    
    def __init__(self):
        super().__init__(TipoPatron.NUMEROS_SIN_COMA)
    
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        match = re.search(PatronesRegex.NUMEROS_SIN_COMA, texto)
        if match:
            nombre_sin_coma = match.group(1).strip()
            nombre_formateado = FormatoNombreConverter.convertir_a_formato_coma(nombre_sin_coma)
            if nombre_formateado and NombreValidator.es_valido(nombre_formateado, min_letras=3):
                return self._crear_candidato(nombre_formateado)
        return None


class DespuesRUCMatcher(PatternMatcher):
    """Matcher para patrón: nombre después de RUC (11 dígitos)."""
    
    def __init__(self):
        super().__init__(TipoPatron.DESPUES_RUC)
    
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        match = re.search(PatronesRegex.RUC_SEGUIDO_NOMBRE, texto, re.IGNORECASE | re.DOTALL)
        if match:
            nombre_sin_coma = match.group(2).strip()
            # Límite de palabras para evitar capturar demasiado texto
            if len(nombre_sin_coma.split()) <= ValidacionConfig.MAX_PALABRAS_NOMBRE_SIN_COMA:
                nombre_formateado = FormatoNombreConverter.convertir_a_formato_coma(nombre_sin_coma)
                if nombre_formateado and NombreValidator.es_valido(nombre_formateado, min_letras=3):
                    return self._crear_candidato(nombre_formateado)
        return None


class ConSaltosMatcher(PatternMatcher):
    """Matcher para patrón: nombres con saltos de línea."""
    
    def __init__(self):
        super().__init__(TipoPatron.CON_SALTOS)
    
    def buscar(self, texto: str) -> Optional[CandidatoNombre]:
        match = re.search(PatronesRegex.NOMBRE_CON_SALTOS, texto)
        if match:
            nombre_completo = f"{match.group(1).strip()}, {match.group(2).strip()}"
            if NombreValidator.es_valido(nombre_completo):
                return self._crear_candidato(nombre_completo)
        return None


class FallbackMatcher(PatternMatcher):
    """Matcher fallback: busca todos los nombres con formato 'APELLIDO(S), NOMBRE(S)'."""
    
    def __init__(self):
        super().__init__(TipoPatron.FALLBACK)
    
    def buscar(self, texto: str) -> List[CandidatoNombre]:
        """
        Busca todos los nombres en formato con coma.
        
        Returns:
            List[CandidatoNombre]: Lista de todos los candidatos encontrados
        """
        matches = re.findall(PatronesRegex.NOMBRE_FALLBACK, texto)
        candidatos = []
        for match in matches:
            if NombreValidator.es_valido(match):
                candidatos.append(self._crear_candidato(match))
        return candidatos


class NombreExtractor:
    
    def __init__(self):
        """Inicializa el extractor con todos los matchers disponibles."""
        self.matchers: List[PatternMatcher] = [
            ConceptoUNMSMConComaMatcher(),
            ConceptoUNMSMSinComaMatcher(),
            NumerosConComaMatcher(),
            NumerosSinComaMatcher(),
            DespuesRUCMatcher(),
            ConSaltosMatcher(),
            FallbackMatcher()
        ]
    
    def extraer(self, texto: str, debug: bool = False) -> Optional[str]:
        candidatos = []
        
        # Buscar con todos los matchers
        for matcher in self.matchers:
            resultado = matcher.buscar(texto)
            
            if resultado:
                # FallbackMatcher devuelve una lista
                if isinstance(resultado, list):
                    candidatos.extend(resultado)
                else:
                    candidatos.append(resultado)
        
        if debug and candidatos:
            logger.info("Candidatos encontrados:")
            for candidato in candidatos:
                logger.info(f"  - {candidato}")
        
        if not candidatos:
            if debug:
                logger.info("No se encontró ningún nombre válido")
            return None
        
        # Filtrar candidatos válidos
        candidatos_validos = [
            c for c in candidatos
            if NombreValidator.es_candidato_valido(c.nombre)
        ]
        
        # Si no hay válidos, usar todos
        if not candidatos_validos:
            candidatos_validos = candidatos
        
        # Seleccionar el mejor candidato
        mejor = min(candidatos_validos, key=lambda c: (c.prioridad, -len(c.nombre)))
        
        if debug:
            logger.info(f"Seleccionado: {mejor.nombre}")
        
        return mejor.nombre


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
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._texto: Optional[str] = None
        self.nombre_extractor = NombreExtractor()
        self.servicio_extractor = ServicioExtractor()
    
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
    
    def extraer_nombre(self, debug: bool = False) -> Optional[str]:
        return self.nombre_extractor.extraer(self.texto, debug)
    
    def extraer_servicios(self, debug: bool = False) -> Optional[str]:
        return self.servicio_extractor.extraer(self.texto, debug)


# Funciones de compatibilidad con código legacy
def extraer_nombre(pdf_path: str, debug: bool = False) -> Optional[str]:
    extractor = PDFExtractor(pdf_path)
    return extractor.extraer_nombre(debug)


def extraer_servicios(pdf_path: str, debug: bool = False) -> Optional[str]:
    extractor = PDFExtractor(pdf_path)
    return extractor.extraer_servicios(debug)
