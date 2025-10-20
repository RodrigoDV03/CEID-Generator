import os
import re
from typing import Dict, List, Any, Optional
from PyPDF2 import PdfReader
from .base import PDFRepository

class PDFRepositoryImpl(PDFRepository):
    
    def __init__(self):
        self._text_cache = {}  # Cache para texto extraído
        
    def exists(self, path: str) -> bool:
        return os.path.exists(path) and path.lower().endswith('.pdf')
    
    def get_metadata(self, path: str) -> Dict[str, Any]:
        if not self.exists(path):
            return {}
            
        try:
            reader = PdfReader(path)
            stat = os.stat(path)
            
            metadata = {
                'size': stat.st_size,
                'modified': stat.st_mtime,
                'pages': len(reader.pages),
                'extension': '.pdf'
            }
            
            # Agregar metadatos del PDF si existen
            if reader.metadata:
                pdf_meta = reader.metadata
                metadata.update({
                    'title': pdf_meta.get('/Title', ''),
                    'author': pdf_meta.get('/Author', ''),
                    'creator': pdf_meta.get('/Creator', ''),
                    'producer': pdf_meta.get('/Producer', ''),
                    'creation_date': pdf_meta.get('/CreationDate', ''),
                })
            
            return metadata
            
        except Exception as e:
            print(f"Error obteniendo metadatos de {path}: {e}")
            return {}
    
    def extract_text(self, path: str) -> str:
        if not self.exists(path):
            raise FileNotFoundError(f"Archivo PDF no encontrado: {path}")
        
        # Crear clave de cache
        file_mtime = os.path.getmtime(path)
        cache_key = f"pdf_text_{path}_{file_mtime}"
        
        # Verificar cache
        if cache_key in self._text_cache:
            return self._text_cache[cache_key]
        
        try:
            reader = PdfReader(path)
            text = ""
            
            # Extraer texto de todas las páginas
            for page in reader.pages:
                text += page.extract_text() + "\n"
            
            # Guardar en cache
            self._text_cache[cache_key] = text
            
            return text
            
        except Exception as e:
            raise Exception(f"Error extrayendo texto de {path}: {e}")
    
    def extract_text_by_page(self, path: str) -> List[str]:
        if not self.exists(path):
            raise FileNotFoundError(f"Archivo PDF no encontrado: {path}")
        
        # Crear clave de cache
        file_mtime = os.path.getmtime(path)
        cache_key = f"pdf_pages_{path}_{file_mtime}"
        
        # Verificar cache
        if cache_key in self._text_cache:
            return self._text_cache[cache_key]
        
        try:
            reader = PdfReader(path)
            pages_text = []
            
            # Extraer texto página por página
            for i, page in enumerate(reader.pages):
                page_text = page.extract_text()
                pages_text.append(page_text)
            
            # Guardar en cache
            self._text_cache[cache_key] = pages_text
            
            return pages_text
            
        except Exception as e:
            raise Exception(f"Error extrayendo texto por páginas de {path}: {e}")
    
    def search_text_in_pdf(self, path: str, search_pattern: str) -> List[Dict[str, Any]]:
        pages_text = self.extract_text_by_page(path)
        matches = []
        
        for page_num, page_text in enumerate(pages_text, 1):
            # Buscar patrón (insensible a mayúsculas/minúsculas)
            pattern_matches = re.finditer(search_pattern, page_text, re.IGNORECASE)
            
            for match in pattern_matches:
                # Obtener contexto (50 caracteres antes y después)
                start = max(0, match.start() - 50)
                end = min(len(page_text), match.end() + 50)
                context = page_text[start:end].strip()
                
                matches.append({
                    'page': page_num,
                    'match': match.group(),
                    'context': context,
                    'position': match.span()
                })
        
        return matches
    
    def extract_specific_data(self, path: str, extractors: Dict[str, str]) -> Dict[str, Any]:
        full_text = self.extract_text(path)
        extracted_data = {}
        
        for field_name, pattern in extractors.items():
            try:
                match = re.search(pattern, full_text, re.IGNORECASE | re.MULTILINE)
                if match:
                    # Si el patrón tiene grupos, tomar el primer grupo
                    if match.groups():
                        extracted_data[field_name] = match.group(1).strip()
                    else:
                        extracted_data[field_name] = match.group(0).strip()
                else:
                    extracted_data[field_name] = None
                    
            except Exception as e:
                print(f"Error extrayendo campo {field_name}: {e}")
                extracted_data[field_name] = None
        
        return extracted_data
    
    def clear_cache(self, pattern: Optional[str] = None) -> None:
        if pattern is None:
            self._text_cache.clear()
        else:
            # Limpiar claves que contengan el patrón
            keys_to_remove = [key for key in self._text_cache.keys() if pattern in key]
            for key in keys_to_remove:
                del self._text_cache[key]
    
    def get_cache_stats(self) -> Dict[str, int]:
        return {
            'text_cache_size': len(self._text_cache)
        }

# Instancia singleton
pdf_repo = PDFRepositoryImpl()