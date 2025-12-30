import os
import sys


class PathUtils:
    
    @staticmethod
    def ruta_absoluta_relativa(path_relativo: str) -> str:
        if getattr(sys, 'frozen', False):
            # Si está empaquetado con PyInstaller
            base_path = sys._MEIPASS
        else:
            # Ejecución normal
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, path_relativo)
    
    @staticmethod
    def crear_directorio(ruta: str) -> bool:
        try:
            os.makedirs(ruta, exist_ok=True)
            return True
        except Exception as e:
            print(f"Error al crear directorio {ruta}: {e}")
            return False
    
    @staticmethod
    def obtener_directorio_padre(ruta: str) -> str:
        return os.path.dirname(ruta)
    
    @staticmethod
    def combinar_rutas(*partes: str) -> str:
        return os.path.join(*partes)
    
    @staticmethod
    def archivo_existe(ruta: str) -> bool:
        return os.path.exists(ruta)
    
    @staticmethod
    def obtener_nombre_archivo(ruta: str, con_extension: bool = True) -> str:
        nombre = os.path.basename(ruta)
        if not con_extension:
            nombre = os.path.splitext(nombre)[0]
        return nombre
    
    @staticmethod
    def cambiar_extension(ruta: str, nueva_extension: str) -> str:
        if not nueva_extension.startswith('.'):
            nueva_extension = '.' + nueva_extension
        
        base = os.path.splitext(ruta)[0]
        return base + nueva_extension
