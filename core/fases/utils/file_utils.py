import os
import shutil
from typing import Optional


class FileUtils:
    
    @staticmethod
    def copiar_archivo(origen: str, destino: str) -> bool:
        try:
            shutil.copy2(origen, destino)
            return True
        except Exception as e:
            print(f"Error al copiar archivo: {e}")
            return False
    
    @staticmethod
    def mover_archivo(origen: str, destino: str) -> bool:
        try:
            shutil.move(origen, destino)
            return True
        except Exception as e:
            print(f"Error al mover archivo: {e}")
            return False
    
    @staticmethod
    def eliminar_archivo(ruta: str) -> bool:
        try:
            if os.path.exists(ruta):
                os.remove(ruta)
            return True
        except Exception as e:
            print(f"Error al eliminar archivo: {e}")
            return False
    
    @staticmethod
    def obtener_tamano_archivo(ruta: str) -> int:
        try:
            return os.path.getsize(ruta)
        except Exception:
            return -1
    
    @staticmethod
    def listar_archivos_directorio(
        directorio: str, 
        extension: Optional[str] = None,
        recursivo: bool = False
    ) -> list:
        archivos = []
        
        try:
            if recursivo:
                for root, _, files in os.walk(directorio):
                    for file in files:
                        ruta_completa = os.path.join(root, file)
                        if extension is None or file.endswith(extension):
                            archivos.append(ruta_completa)
            else:
                for item in os.listdir(directorio):
                    ruta_completa = os.path.join(directorio, item)
                    if os.path.isfile(ruta_completa):
                        if extension is None or item.endswith(extension):
                            archivos.append(ruta_completa)
        except Exception as e:
            print(f"Error al listar archivos: {e}")
        
        return archivos
    
    @staticmethod
    def crear_archivo_vacio(ruta: str) -> bool:
        try:
            with open(ruta, 'w') as f:
                pass
            return True
        except Exception as e:
            print(f"Error al crear archivo: {e}")
            return False
    
    @staticmethod
    def renombrar_archivo(ruta_antigua: str, ruta_nueva: str) -> bool:
        try:
            os.rename(ruta_antigua, ruta_nueva)
            return True
        except Exception as e:
            print(f"Error al renombrar archivo: {e}")
            return False
    
    @staticmethod
    def verificar_permisos_escritura(ruta: str) -> bool:
        return os.access(ruta, os.W_OK)
    
    @staticmethod
    def obtener_extension(ruta: str) -> str:
        return os.path.splitext(ruta)[1]
