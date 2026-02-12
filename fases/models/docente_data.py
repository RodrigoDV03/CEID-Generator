from dataclasses import dataclass


@dataclass
class DocenteData:
    # Identificación
    nombre: str
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
        import re
        return re.sub(r'[\\/*?:"<>|]', "", self.nombre)
    
    @property
    def dni_formateado(self) -> str:
        dni = self.dni.strip()
        return dni.zfill(8) if len(dni) < 8 else dni
    
    def __post_init__(self):
        # Limpiar espacios en blanco
        self.nombre = self.nombre.strip()
        self.dni = self.dni.strip()
        self.ruc = self.ruc.strip()
        self.estado_docente = self.estado_docente.strip().upper()
        
        # Validación básica
        if not self.nombre or self.nombre == "N/A":
            raise ValueError("El nombre del docente es requerido")
