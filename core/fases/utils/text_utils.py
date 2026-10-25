import re
import unicodedata
import pandas as pd
from typing import Any


class TextUtils:
    
    @staticmethod
    def limpiar_nombre_archivo(nombre: str) -> str:
        return re.sub(r'[\\/*?:"<>|]', "", nombre)
    
    @staticmethod
    def limpiar_numero(valor: Any) -> str:
        if pd.isna(valor):
            return ""
        return str(valor).split('.')[0]
    
    @staticmethod
    def limpiar_espacios(texto: str) -> str:
        return texto.strip() if isinstance(texto, str) else str(texto).strip()
    
    @staticmethod
    def formatear_dni(dni: str, longitud: int = 8) -> str:
        dni_limpio = TextUtils.limpiar_espacios(dni)
        return dni_limpio.zfill(longitud) if len(dni_limpio) < longitud else dni_limpio
    
    @staticmethod
    def es_texto_vacio(texto: Any) -> bool:
        if texto is None or pd.isna(texto):
            return True
        
        texto_str = str(texto).strip()
        return texto_str == "" or texto_str.upper() == "N/A"
    
    @staticmethod
    def quitar_acentos(texto: str) -> str:
        """
        Quita tildes/diacríticos preservando mayúsculas y minúsculas
        (p. ej. "áéí" -> "aei", "Núñez" -> "Nunez").

        Es el núcleo compartido detrás de `normalizar_texto`: antes existían
        tres implementaciones casi idénticas de esto mismo, repartidas en
        `core/planillas/transformations.py`, `core/correos/pdf_extractor.py`
        y `core/control_pagos/control_pagos.py`.
        """
        texto_descompuesto = unicodedata.normalize('NFKD', str(texto))
        return ''.join(ch for ch in texto_descompuesto if not unicodedata.combining(ch))

    @staticmethod
    def normalizar_texto(texto: str) -> str:
        """
        Normaliza texto para compararlo de forma tolerante a mayúsculas,
        espacios y tildes: es la forma canónica usada en todo el proyecto
        para decidir si dos nombres de docente (u otro texto) "son el
        mismo". Úsala en vez de reimplementar tu propia versión.
        """
        if texto is None or pd.isna(texto):
            return ""

        texto_limpio = TextUtils.limpiar_espacios(str(texto))
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio)
        texto_limpio = TextUtils.quitar_acentos(texto_limpio)
        return texto_limpio.upper()
    
    @staticmethod
    def capitalizar_primera(texto: str) -> str:
        if not texto:
            return texto
        return texto[0].upper() + texto[1:].lower()
    
    @staticmethod
    def separar_por_delimitador(texto: str, delimitador: str = "/") -> list:
        if not isinstance(texto, str):
            return []
        
        return [elem.strip() for elem in texto.split(delimitador) if elem.strip()]
    
    @staticmethod
    def unir_con_comas(elementos: list, ultimo_separador: str = "y") -> str:
        if not elementos:
            return ""
        
        if len(elementos) == 1:
            return elementos[0]
        
        if len(elementos) == 2:
            return f"{elementos[0]} {ultimo_separador} {elementos[1]}"
        
        return f"{', '.join(elementos[:-1])}, {ultimo_separador} {elementos[-1]}"

    @staticmethod
    def modalidad_a_texto(modalidad: str, permitir_na_vacio: bool = False) -> str:
        modalidad_upper = TextUtils.limpiar_espacios(modalidad).upper()

        if modalidad_upper == "INPERSON":
            return "presencial"
        if modalidad_upper in ["VIRTUAL", "SUPERINTENSIVO VIRTUAL", "INTENSIVO VIRTUAL"]:
            return "virtual"
        if modalidad_upper == "MIXTA":
            return "mixta"
        if modalidad_upper == "N/A" and permitir_na_vacio:
            return ""
        return "híbrida"
