import json
import os
from typing import Dict, Any

class ConfigManager:
    _instance = None
    _config = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(ConfigManager, cls).__new__(cls)
        return cls._instance
    
    def __init__(self):
        if self._config is None:
            self.load_config()
    
    def load_config(self, config_path: str = "config.json") -> None:
        try:
            # Obtener la ruta absoluta del archivo de configuración
            if not os.path.isabs(config_path):
                # Si es una ruta relativa, usar la carpeta del script actual
                script_dir = os.path.dirname(os.path.abspath(__file__))
                project_root = os.path.dirname(script_dir)  # Subir un nivel desde utils/
                config_path = os.path.join(project_root, config_path)
            
            with open(config_path, 'r', encoding='utf-8') as file:
                self._config = json.load(file)
                
        except FileNotFoundError:
            raise FileNotFoundError(f"Archivo de configuración no encontrado: {config_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Error al parsear el archivo de configuración: {e}")
        except Exception as e:
            raise Exception(f"Error inesperado al cargar configuración: {e}")
    
    def get(self, key_path: str, default: Any = None) -> Any:
        if self._config is None:
            return default
            
        keys = key_path.split('.')
        value = self._config
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
                
        return value
    
    def get_email_config(self) -> Dict[str, Any]:
        """Obtiene la configuración completa de email"""
        return self.get("email", {})
    
    def get_gui_config(self) -> Dict[str, Any]:
        """Obtiene la configuración completa de GUI"""
        return self.get("gui", {})
    
    def get_colors(self) -> Dict[str, str]:
        """Obtiene la configuración de colores de la GUI"""
        return self.get("gui.colors", {})
    
    def get_fonts(self) -> Dict[str, list]:
        """Obtiene la configuración de fuentes de la GUI"""
        return self.get("gui.fonts", {})
    
    def get_paths(self) -> Dict[str, str]:
        """Obtiene la configuración de rutas"""
        return self.get("paths", {})
    
    def reload_config(self, config_path: str = "config.json") -> None:
        """Recarga la configuración desde el archivo"""
        self._config = None
        self.load_config(config_path)

# Instancia singleton global
config = ConfigManager()

# Funciones helper para acceso rápido (compatibilidad con código existente)
def get_config(key_path: str, default: Any = None) -> Any:
    """Función helper para acceso rápido a configuración"""
    return config.get(key_path, default)

def get_email_config() -> Dict[str, Any]:
    """Función helper para configuración de email"""
    return config.get_email_config()

def get_gui_colors() -> Dict[str, str]:
    """Función helper para colores de GUI"""
    return config.get_colors()

def get_gui_fonts() -> Dict[str, list]:
    """Función helper para fuentes de GUI"""
    return config.get_fonts()

def get_paths_config() -> Dict[str, str]:
    """Función helper para configuración de rutas"""
    return config.get_paths()

def reload_config(config_path: str = "config.json") -> None:
    """Función helper para recargar configuración"""
    return config.reload_config(config_path)