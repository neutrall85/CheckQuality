import os
import gc
from typing import Optional

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

def get_current_memory_usage() -> Optional[float]:
    """Возвращает текущее использование памяти процессом в МБ."""
    if HAS_PSUTIL:
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / 1024 / 1024
    return None

def force_garbage_collection():
    """Принудительный вызов сборщика мусора."""
    gc.collect()