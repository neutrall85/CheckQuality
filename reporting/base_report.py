from abc import ABC, abstractmethod
from core.models.statistics import Statistics
from core.models.config import ConfigModel


class BaseReport(ABC):
    """Базовый класс для всех генераторов отчётов."""

    @abstractmethod
    def generate(self, statistics: Statistics, config: ConfigModel, output_path: str) -> str:
        """Генерирует отчёт и возвращает путь к созданному файлу."""
        ...