from abc import ABC, abstractmethod
from typing import Any, Dict, Optional, List
import pandas as pd

class BaseRepository(ABC):    
    @abstractmethod
    def exists(self, path: str) -> bool:
        pass
    
    @abstractmethod
    def get_metadata(self, path: str) -> Dict[str, Any]:
        pass

class ExcelRepository(BaseRepository):    
    @abstractmethod
    def read_sheet(self, path: str, sheet_name: Optional[str] = None, **kwargs) -> pd.DataFrame:
        pass
    
    @abstractmethod
    def get_sheet_names(self, path: str) -> List[str]:
        pass
    
    @abstractmethod
    def write_excel(self, data: Dict[str, pd.DataFrame], path: str) -> bool:
        pass

class PDFRepository(BaseRepository):    
    @abstractmethod
    def extract_text(self, path: str) -> str:
        pass
    
    @abstractmethod
    def extract_text_by_page(self, path: str) -> List[str]:
        pass

class CacheRepository(ABC):    
    @abstractmethod
    def get(self, key: str) -> Any:
        pass
    
    @abstractmethod
    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None:
        pass
    
    @abstractmethod
    def clear(self, pattern: Optional[str] = None) -> None:
        pass
    
    @abstractmethod
    def exists(self, key: str) -> bool:
        pass