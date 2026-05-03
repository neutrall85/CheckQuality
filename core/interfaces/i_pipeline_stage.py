from abc import ABC, abstractmethod
from typing import Optional
from core.models.document import Document


class IPipelineStage(ABC):
    """Интерфейс этапа конвейера обработки."""

    @abstractmethod
    def process(self, document: Document) -> Optional[Document]:
        """
        Обрабатывает документ.
        Возвращает документ или None (если документ должен быть отброшен).
        """
        ...