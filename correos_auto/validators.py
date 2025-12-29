import re
from typing import Optional, Tuple
from .config import (
    PALABRAS_PROHIBIDAS,
    ValidacionConfig,
    PatronesRegex
)


class NombreValidator:
    
    @staticmethod
    def es_valido(nombre: str, min_letras: int = ValidacionConfig.MIN_LETRAS_NOMBRE) -> bool:
        """
        Valida que un nombre tenga el formato y contenido correcto.
        
        Args:
            nombre: Nombre en formato "APELLIDOS, NOMBRES"
            min_letras: Mínimo de letras requeridas en apellidos y nombres
        """
        if not nombre:
            return False
            
        partes = nombre.split(',')
        if len(partes) < 2:
            return False
            
        apellidos = partes[0].strip()
        nombres = partes[1].strip()
        
        # Contar solo letras (excluyendo espacios y otros caracteres)
        apellidos_letras = sum(1 for c in apellidos if c.isalpha())
        nombres_letras = sum(1 for c in nombres if c.isalpha())
        
        return apellidos_letras >= min_letras and nombres_letras >= min_letras
    
    @staticmethod
    def limpiar(nombre_raw: str) -> str:

        return re.sub(PatronesRegex.ESPACIOS_MULTIPLES, " ", nombre_raw).strip()
    
    @staticmethod
    def tiene_palabras_prohibidas(nombre: str) -> bool:

        nombre_upper = nombre.upper()
        return any(palabra in nombre_upper for palabra in PALABRAS_PROHIBIDAS)
    
    @staticmethod
    def esta_en_rango_longitud(nombre: str) -> bool:

        longitud = len(nombre)
        return ValidacionConfig.MIN_LONGITUD_NOMBRE <= longitud <= ValidacionConfig.MAX_LONGITUD_NOMBRE
    
    @staticmethod
    def es_candidato_valido(nombre: str) -> bool:

        return (
            not NombreValidator.tiene_palabras_prohibidas(nombre) and
            NombreValidator.esta_en_rango_longitud(nombre)
        )


class FormatoNombreConverter:
    
    @staticmethod
    def es_palabra_prohibida(palabra: str) -> bool:

        return palabra.upper() in PALABRAS_PROHIBIDAS
    
    @staticmethod
    def convertir_a_formato_coma(nombre_sin_coma: str) -> Optional[str]:
        palabras = nombre_sin_coma.split()
        
        if not palabras:
            return None
            
        # Validar primera palabra
        if FormatoNombreConverter.es_palabra_prohibida(palabras[0]):
            return None
        
        # Necesita al menos 3 palabras (2 apellidos + 1 nombre)
        if len(palabras) < ValidacionConfig.MIN_PALABRAS_NOMBRE_COMPLETO:
            return None
        
        # Dividir en apellidos (primeras 2 palabras) y nombres (resto)
        apellidos = " ".join(palabras[:2])
        nombres = " ".join(palabras[2:])
        
        return f"{apellidos}, {nombres}"
    
    @staticmethod
    def intentar_conversion(nombre: str) -> Tuple[str, bool]:
        # Si ya tiene coma, no convertir
        if ',' in nombre:
            return nombre, False
        
        # Intentar conversión
        convertido = FormatoNombreConverter.convertir_a_formato_coma(nombre)
        
        if convertido:
            return convertido, True
        
        return nombre, False


class CandidatoNombre:
    
    def __init__(self, nombre: str, patron: str, prioridad: int):
        self.nombre = nombre
        self.patron = patron
        self.prioridad = prioridad
    
    def __repr__(self) -> str:
        return f"CandidatoNombre(nombre='{self.nombre}', patron='{self.patron}', prioridad={self.prioridad})"
    
    def es_mejor_que(self, otro: 'CandidatoNombre') -> bool:
        if self.prioridad != otro.prioridad:
            return self.prioridad < otro.prioridad
        return len(self.nombre) > len(otro.nombre)


def seleccionar_mejor_candidato(candidatos: list) -> Optional[str]:
    if not candidatos:
        return None
    
    # Filtrar candidatos válidos (sin palabras prohibidas y longitud adecuada)
    candidatos_validos = [
        c for c in candidatos
        if NombreValidator.es_candidato_valido(c[1] if isinstance(c, tuple) else c.nombre)
    ]
    
    # Si no hay válidos, usar todos los candidatos
    if not candidatos_validos:
        candidatos_validos = candidatos
    
    # Si son tuplas, convertir a objetos CandidatoNombre
    if candidatos_validos and isinstance(candidatos_validos[0], tuple):
        from .config import PRIORIDAD_PATRONES, TipoPatron
        
        # Crear mapeo de string a prioridad
        prioridades_str = {patron.value: prio for patron, prio in PRIORIDAD_PATRONES.items()}
        
        candidatos_objetos = [
            CandidatoNombre(
                nombre=nombre,
                patron=patron,
                prioridad=prioridades_str.get(patron, 99)
            )
            for patron, nombre in candidatos_validos
        ]
    else:
        candidatos_objetos = candidatos_validos
    
    # Seleccionar el mejor
    mejor = min(candidatos_objetos, key=lambda c: (c.prioridad, -len(c.nombre)))
    
    return mejor.nombre
