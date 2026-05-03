from abc import ABC, abstractmethod
from typing import Optional
from core.models.document import Document


class IValidator(ABC):
    """Интерфейс валидатора документа."""

    @abstractmethod
    def validate(self, document: Document) -> Optional[str]:
        """Возвращает сообщение об ошибке или None, если всё корректно."""
        ...