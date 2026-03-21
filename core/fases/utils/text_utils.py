import re
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
    def normalizar_texto(texto: str) -> str:
        texto = TextUtils.limpiar_espacios(texto)
        # Eliminar espacios múltiples
        texto = re.sub(r'\s+', ' ', texto)
        return texto
    
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
