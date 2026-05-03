from abc import ABC, abstractmethod
from typing import Any


class ITemplateManager(ABC):
    """Интерфейс менеджера шаблонов (например, PowerPoint)."""

    @abstractmethod
    def load_template(self, path: str) -> Any:
        """Загружает шаблон и возвращает объект темы/презентации."""
        ...