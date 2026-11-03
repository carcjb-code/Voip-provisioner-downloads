"""
app/monitor_cache.py - Caché en memoria con expiración
Evita saturar el Asterisk con consultas repetidas.
"""
import time
import threading
import logging
from typing import Any, Callable, Dict, Optional

logger = logging.getLogger(__name__)


class MonitorCache:
    """Caché thread-safe con TTL por clave."""
    
    def __init__(self, default_ttl: float = 5.0):
        self._cache: Dict[str, Dict[str, Any]] = {}
        self._lock = threading.Lock()
        self.default_ttl = default_ttl
    
    def get(self, key: str) -> Optional[Any]:
        """Obtiene un valor si no ha expirado."""
        with self._lock:
            entry = self._cache.get(key)
            if not entry:
                return None
            if time.time() > entry['expires']:
                del self._cache[key]
                return None
            return entry['value']
    
    def set(self, key: str, value: Any, ttl: Optional[float] = None):
        """Guarda un valor con TTL."""
        ttl = ttl if ttl is not None else self.default_ttl
        with self._lock:
            self._cache[key] = {
                'value': value,
                'expires': time.time() + ttl
            }
    
    def invalidate(self, key: str):
        """Invalida una clave específica."""
        with self._lock:
            self._cache.pop(key, None)
    
    def invalidate_prefix(self, prefix: str):
        """Invalida todas las claves que empiezan con el prefijo."""
        with self._lock:
            keys_to_delete = [k for k in self._cache if k.startswith(prefix)]
            for k in keys_to_delete:
                del self._cache[k]
    
    def clear(self):
        """Limpia toda la caché."""
        with self._lock:
            self._cache.clear()
    
    def get_or_compute(self, key: str, compute_fn: Callable, ttl: Optional[float] = None) -> Any:
        """Obtiene de caché o ejecuta la función y cachea el resultado."""
        value = self.get(key)
        if value is not None:
            return value
        
        try:
            value = compute_fn()
            self.set(key, value, ttl)
            return value
        except Exception as e:
            logger.error(f"Error computando caché para {key}: {e}")
            raise


# Instancia global compartida
cache = MonitorCache(default_ttl=5.0)