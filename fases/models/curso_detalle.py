from dataclasses import dataclass


@dataclass
class CursoDetalle:
    """
    Representa un curso o servicio individual con su modalidad específica.
    Se usa para generar descripciones detalladas con modalidades en los documentos.
    """
    nombre: str
    modalidad: str  # INPERSON, VIRTUAL, INTENSIVO VIRTUAL, N/A
    tipo_servicio: str  # CURSO_DICTADO, EXAMEN_CLASIF, DISENO_EXAMENES, SERVICIO_ACTUALIZACION, BONO
    horas: int
    monto: float
    
    @property
    def modalidad_texto(self) -> str:
        """
        Convierte la modalidad del curso a texto legible.
        INTENSIVO VIRTUAL se considera como "virtual" para la redacción.
        """
        modalidad_upper = self.modalidad.strip().upper()
        
        if modalidad_upper == "INPERSON":
            return "presencial"
        elif modalidad_upper in ["VIRTUAL", "INTENSIVO VIRTUAL"]:
            return "virtual"
        elif modalidad_upper == "MIXTA":
            return "mixta"
        elif modalidad_upper == "N/A":
            return ""  # Sin modalidad (para diseño de exámenes, bono)
        else:
            return "híbrida"  # Valor por defecto
    
    @property
    def es_curso_academico(self) -> bool:
        """Indica si es un curso académico regular (no un servicio especial)"""
        return self.tipo_servicio == "CURSO_DICTADO"
    
    @property
    def es_servicio_especial(self) -> bool:
        """Indica si es un servicio especial (examen, diseño, actualización)"""
        return self.tipo_servicio in ["EXAMEN_CLASIF", "SERVICIO_ACTUALIZACION"]
    
    @property
    def es_sin_modalidad(self) -> bool:
        """Indica si el servicio no requiere especificar modalidad"""
        return self.tipo_servicio in ["DISENO_EXAMENES", "BONO"]
    
    def generar_descripcion_individual(self) -> str:
        """
        Genera la descripción textual de este curso/servicio.
        
        Returns:
            str: Descripción formateada del curso/servicio
        """
        if self.tipo_servicio == "CURSO_DICTADO":
            # "28 horas de clases de Inglés Avanzado 2 bajo la modalidad presencial"
            return f"{self.horas} horas de clases de {self.nombre} bajo la modalidad {self.modalidad_texto}"
        
        elif self.tipo_servicio == "EXAMEN_CLASIF":
            # "10 horas de examen de clasificación bajo la modalidad virtual"
            return f"{self.horas} horas de examen de clasificación bajo la modalidad {self.modalidad_texto}"
        
        elif self.tipo_servicio == "SERVICIO_ACTUALIZACION":
            # "Servicio de actualización de materiales de enseñanza bajo la modalidad virtual"
            return f"Servicio de actualización de materiales de enseñanza bajo la modalidad {self.modalidad_texto}"
        
        elif self.tipo_servicio == "DISENO_EXAMENES":
            # "8 horas de diseño de exámenes" (sin modalidad)
            return f"{self.horas} horas de diseño de exámenes"
        
        elif self.tipo_servicio == "BONO":
            # "servicio de diseño y evaluación del examen anual" (sin modalidad)
            return "servicio de diseño y evaluación del examen anual"
        
        else:
            # Fallback genérico
            return f"{self.horas} horas de {self.nombre}"
    
    def __post_init__(self):
        # Validación básica
        if not self.nombre or self.nombre.strip() == "":
            raise ValueError("El nombre del curso es requerido")
        
        if self.horas < 0:
            raise ValueError("Las horas no pueden ser negativas")
        
        if self.monto < 0:
            raise ValueError("El monto no puede ser negativo")
