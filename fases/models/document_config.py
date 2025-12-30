from dataclasses import dataclass
from enum import Enum


class TipoDocumento(Enum):
    OFICIO = "oficio"
    TDR = "tdr"
    COTIZACION = "cotizacion"
    CONFORMIDAD = "conformidad"
    CONTROL_AVANCE = "control_avance"


class TipoContrato(Enum):
    CONTRATO = "CONTRATO"
    TERCERO = "TERCERO"


class TipoFase(Enum):
    INICIAL = "inicial"
    FINAL = "final"


class TipoDocente(Enum):
    DOCENTE = "planilla docente"
    DOCENTE_CON_CONTRATO = "planilla docente (con contrato)"
    DOCENTE_SIN_CONTRATO = "planilla docente (sin contrato)"
    ADMINISTRATIVO = "administrativo"


class NumeroArmada(Enum):
    PRIMERA = "primera"
    SEGUNDA = "segunda"
    TERCERA = "tercera"
    SIN_ARMADA = "sin armada"


@dataclass
class DocumentConfig:
    # Información del periodo
    mes: str
    anio: int
    numero_armada: str
    
    # Tipo de proceso
    tipo_fase: str  # "inicial" o "final"
    tipo_docente: str  # "planilla docente", "administrativo", etc.
    
    # Rutas
    carpeta_destino: str
    
    @property
    def es_administrativo(self) -> bool:
        return "administrativo" in self.tipo_docente.lower()
    
    @property
    def es_docente(self) -> bool:
        return "docente" in self.tipo_docente.lower()
    
    @property
    def es_con_contrato(self) -> bool:
        return "con contrato" in self.tipo_docente.lower()
    
    @property
    def es_sin_contrato(self) -> bool:
        return "sin contrato" in self.tipo_docente.lower()
    
    @property
    def modalidad_servicio(self) -> str:
        return "presencial" if self.es_administrativo else "híbrida"
    
    def obtener_nombre_carpeta_fase(self) -> str:
        return "FASE INICIAL" if self.tipo_fase.lower() == "inicial" else "FASE FINAL"
    
    def obtener_ruta_firma(self, nombre_docente: str) -> str:
        from fases.functions import ruta_absoluta_relativa
        
        carpeta = "firmas_admin" if self.es_administrativo else "firmas_docentes"
        return ruta_absoluta_relativa(f"{carpeta}/{nombre_docente}.png")


@dataclass
class PlantillaDocumento:
    nombre: str
    ruta_template: str
    tipo_documento: TipoDocumento
    requiere_firma: bool = False
    convertir_pdf: bool = False
    
    def generar_nombre_salida(self, nombre_docente: str, mes: str, anio: int) -> str:
        prefijo_map = {
            TipoDocumento.OFICIO: "OFICIO",
            TipoDocumento.TDR: "TDR",
            TipoDocumento.COTIZACION: "COTIZACIÓN",
            TipoDocumento.CONFORMIDAD: "CONFORMIDAD",
            TipoDocumento.CONTROL_AVANCE: "CONTROL DE AVANCE"
        }
        
        prefijo = prefijo_map.get(self.tipo_documento, "DOCUMENTO")
        return f"{prefijo} - {nombre_docente} - {mes} {anio}.docx"
