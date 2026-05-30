from dataclasses import dataclass
from core.fases.utils import TextUtils


@dataclass
class DocenteData:
    # Identificación
    nombre: str
    tipo_documento: str = ""
    dni: str = ""
    ruc: str = ""
    
    # Contacto
    direccion: str = ""
    correo: str = ""
    celular: str = ""
    
    # Información académica/laboral
    curso: str = ""
    categoria_letra: str = ""
    formacion_academica: str = ""
    experiencia_laboral: str = ""
    requisitos_adicional: str = ""
    finalidad_publica: str = ""
    especialidad: str = ""
    actividades_admin: str = ""
    idioma: str = ""
    modalidad: str = ""
    
    # Tipo de contrato
    estado_docente: str = "TERCERO"  # "CONTRATO" o "TERCERO"
    numero_contrato: str = ""
    
    @property
    def es_contrato(self) -> bool:
        return self.estado_docente.upper() == "CONTRATO"
    
    @property
    def es_tercero(self) -> bool:
        return self.estado_docente.upper() == "TERCERO"
    
    @property
    def nombre_limpio(self) -> str:
        return TextUtils.limpiar_nombre_archivo(self.nombre)
    
    @property
    def dni_formateado(self) -> str:
        return TextUtils.formatear_dni(self.dni)
    
    @property
    def modalidad_texto(self) -> str:
        """Convierte el valor de la columna Modalidad al texto apropiado."""
        return TextUtils.modalidad_a_texto(self.modalidad)
    
    def __post_init__(self):
        # Limpiar espacios en blanco
        self.nombre = self.nombre.strip()
        self.dni = self.dni.strip()
        self.ruc = self.ruc.strip()
        self.estado_docente = self.estado_docente.strip().upper()
        
        # Validación básica
        if not self.nombre or self.nombre == "N/A":
            raise ValueError("El nombre del docente es requerido")
