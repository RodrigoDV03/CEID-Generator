from .excel_repository import excel_repo, ExcelRepositoryImpl
from .pdf_repository import pdf_repo, PDFRepositoryImpl
from .cache_repository import cache_manager, get_cache, clear_all_caches

# Exports principales
__all__ = [
    # Instancias singleton listas para usar
    'excel_repo',
    'pdf_repo',
    'cache_manager',
    
    # Clases para instanciación manual si se necesita
    'ExcelRepositoryImpl',
    'PDFRepositoryImpl',
    
    # Funciones helper
    'get_cache',
    'clear_all_caches'
]

# Información del módulo
__version__ = "1.0.0"
__author__ = "CEID Generator Team"