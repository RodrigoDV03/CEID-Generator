import os
import pandas as pd
from typing import Dict, List, Optional, Any
from .base import ExcelRepository

class ExcelRepositoryImpl(ExcelRepository):
    def __init__(self):
        self._cache = {}  # Cache interno simple
        self._file_cache = {}  # Cache de archivos completos
        
    def exists(self, path: str) -> bool:
        return os.path.exists(path) and path.lower().endswith(('.xlsx', '.xls'))
    
    def get_metadata(self, path: str) -> Dict[str, Any]:
        if not self.exists(path):
            return {}
            
        stat = os.stat(path)
        return {
            'size': stat.st_size,
            'modified': stat.st_mtime,
            'extension': os.path.splitext(path)[1],
            'sheets': self.get_sheet_names(path)
        }
    
    def get_sheet_names(self, path: str) -> List[str]:
        cache_key = f"sheets_{path}_{os.path.getmtime(path)}"
        
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        try:
            excel_file = pd.ExcelFile(path)
            sheets = excel_file.sheet_names
            self._cache[cache_key] = sheets
            return sheets
        except Exception as e:
            print(f"Error leyendo hojas de {path}: {e}")
            return []
    
    def read_sheet(self, path: str, sheet_name: Optional[str] = None, **kwargs) -> pd.DataFrame:
        if not self.exists(path):
            raise FileNotFoundError(f"Archivo Excel no encontrado: {path}")
        
        # Si no se especifica sheet_name, usar la primera hoja (índice 0)
        actual_sheet_name = sheet_name if sheet_name is not None else 0
        
        # Crear clave de cache basada en archivo, hoja y parámetros
        file_mtime = os.path.getmtime(path)
        kwargs_str = "_".join(f"{k}={v}" for k, v in sorted(kwargs.items()))
        cache_key = f"excel_{path}_{actual_sheet_name}_{file_mtime}_{kwargs_str}"
        
        # Verificar cache
        if cache_key in self._cache:
            return self._cache[cache_key].copy()  # Retornar copia para evitar modificaciones
        
        try:
            # Leer Excel con sheet_name específico (nunca None)
            df = pd.read_excel(path, sheet_name=actual_sheet_name, **kwargs)
            
            # Limpieza automática de columnas
            if hasattr(df, 'columns'):
                df.columns = df.columns.str.strip()
            
            # Guardar en cache
            self._cache[cache_key] = df.copy()
            
            return df
            
        except Exception as e:
            raise Exception(f"Error leyendo Excel {path}, hoja {sheet_name}: {e}")
    
    def write_excel(self, data: Dict[str, pd.DataFrame], path: str) -> bool:
        try:
            # Crear directorio si no existe y no es la raíz
            dir_path = os.path.dirname(path)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path, exist_ok=True)
            
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                for sheet_name, df in data.items():
                    # Validar que es un DataFrame
                    if not isinstance(df, pd.DataFrame):
                        print(f"⚠️ Advertencia: {sheet_name} no es un DataFrame válido (tipo: {type(df)})")
                        continue
                        
                    if df.empty:
                        print(f"⚠️ Advertencia: {sheet_name} está vacío")
                        continue
                        
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Limpiar cache relacionado con este archivo
            self._clear_file_cache(path)
            
            return True
            
        except Exception as e:
            print(f"Error escribiendo Excel {path}: {e}")
            return False
    
    def read_multiple_sheets(self, path: str, sheet_names: List[str], **kwargs) -> Dict[str, pd.DataFrame]:
        result = {}
        
        for sheet_name in sheet_names:
            try:
                result[sheet_name] = self.read_sheet(path, sheet_name, **kwargs)
            except Exception as e:
                print(f"Error leyendo hoja {sheet_name}: {e}")
                result[sheet_name] = pd.DataFrame()  # DataFrame vacío en caso de error
                
        return result
    
    def clear_cache(self, pattern: Optional[str] = None) -> None:
        if pattern is None:
            self._cache.clear()
            self._file_cache.clear()
        else:
            # Limpiar claves que contengan el patrón
            keys_to_remove = [key for key in self._cache.keys() if pattern in key]
            for key in keys_to_remove:
                del self._cache[key]
                
            keys_to_remove = [key for key in self._file_cache.keys() if pattern in key]
            for key in keys_to_remove:
                del self._file_cache[key]
    
    def _clear_file_cache(self, file_path: str) -> None:
        self.clear_cache(file_path)
    
    def get_cache_stats(self) -> Dict[str, int]:
        return {
            'cache_size': len(self._cache),
            'file_cache_size': len(self._file_cache)
        }

# Instancia singleton
excel_repo = ExcelRepositoryImpl()