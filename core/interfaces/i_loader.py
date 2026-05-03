from abc import ABC, abstractmethod
from typing import Iterator
from core.models.document import Document


class IDataLoader(ABC):
    """Интерфейс загрузчика данных."""

    @abstractmethod
    def load(self, file_path: str) -> Iterator[Document]:
        """Итеративно загружает документы из файла."""
        ...