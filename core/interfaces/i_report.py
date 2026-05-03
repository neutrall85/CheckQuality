from abc import ABC, abstractmethod
from core.models.statistics import Statistics
from core.models.config import ConfigModel


class IReportGenerator(ABC):
    """Интерфейс генератора отчётов."""

    @abstractmethod
    def generate(self, statistics: Statistics, config: ConfigModel) -> str:
        """Генерирует отчёт и возвращает путь к сохранённому файлу."""
        ...