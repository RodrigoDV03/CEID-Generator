import time
from typing import Any, Dict, Optional, Set
from .base import CacheRepository

class MemoryCacheRepository(CacheRepository):    
    def __init__(self):
        self._cache: Dict[str, Any] = {}
        self._timestamps: Dict[str, float] = {}
        self._ttls: Dict[str, float] = {}  # Time To Live por clave
        
    def get(self, key: str) -> Any:
        if not self.exists(key):
            return None
            
        # Verificar si ha expirado
        if self._has_expired(key):
            self._remove_key(key)
            return None
            
        return self._cache[key]
    
    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None:
        self._cache[key] = value
        self._timestamps[key] = time.time()
        
        if ttl is not None:
            self._ttls[key] = time.time() + ttl
        elif key in self._ttls:
            # Si no se especifica TTL, eliminar TTL previo
            del self._ttls[key]
    
    def exists(self, key: str) -> bool:
        return key in self._cache
    
    def clear(self, pattern: Optional[str] = None) -> None:
        if pattern is None:
            # Limpiar todo
            self._cache.clear()
            self._timestamps.clear()
            self._ttls.clear()
        else:
            # Limpiar claves que contengan el patrón
            keys_to_remove = [key for key in self._cache.keys() if pattern in key]
            for key in keys_to_remove:
                self._remove_key(key)
    
    def cleanup_expired(self) -> int:
        expired_keys = []
        current_time = time.time()
        
        for key, expiry_time in self._ttls.items():
            if current_time > expiry_time:
                expired_keys.append(key)
        
        for key in expired_keys:
            self._remove_key(key)
            
        return len(expired_keys)
    
    def get_stats(self) -> Dict[str, Any]:
        current_time = time.time()
        expired_count = sum(1 for expiry in self._ttls.values() if current_time > expiry)
        
        return {
            'total_keys': len(self._cache),
            'expired_keys': expired_count,
            'active_keys': len(self._cache) - expired_count,
            'keys_with_ttl': len(self._ttls),
            'oldest_timestamp': min(self._timestamps.values()) if self._timestamps else None,
            'newest_timestamp': max(self._timestamps.values()) if self._timestamps else None
        }
    
    def get_keys(self, pattern: Optional[str] = None) -> Set[str]:
        # Limpiar expiradas primero
        self.cleanup_expired()
        
        keys = set(self._cache.keys())
        
        if pattern:
            keys = {key for key in keys if pattern in key}
            
        return keys
    
    def _has_expired(self, key: str) -> bool:
        if key not in self._ttls:
            return False  # Sin TTL, nunca expira
            
        return time.time() > self._ttls[key]
    
    def _remove_key(self, key: str) -> None:
        self._cache.pop(key, None)
        self._timestamps.pop(key, None)
        self._ttls.pop(key, None)

class CacheManager:    
    def __init__(self):
        self.memory_cache = MemoryCacheRepository()
        # Aquí podrían agregarse otros tipos de cache (Redis, SQLite, etc.)
    
    def get_cache(self, cache_type: str = "memory") -> CacheRepository:
        if cache_type == "memory":
            return self.memory_cache
        else:
            raise ValueError(f"Tipo de cache no soportado: {cache_type}")
    
    def clear_all_caches(self) -> None:
        """Limpia todos los tipos de cache"""
        self.memory_cache.clear()
    
    def get_global_stats(self) -> Dict[str, Any]:
        """Obtiene estadísticas globales de todos los caches"""
        return {
            'memory_cache': self.memory_cache.get_stats()
        }

# Instancia singleton global
cache_manager = CacheManager()

# Funciones helper para compatibilidad
def get_cache(cache_type: str = "memory") -> CacheRepository:
    """Función helper para obtener cache"""
    return cache_manager.get_cache(cache_type)

def clear_all_caches() -> None:
    """Función helper para limpiar todos los caches"""
    cache_manager.clear_all_caches()