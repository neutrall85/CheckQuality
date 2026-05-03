from core.models.document import Document


class Normalizer:
    """Дополнительная нормализация данных (при необходимости)."""

    @staticmethod
    def normalize(document: Document) -> Document:
        # Основная нормализация уже выполнена в загрузчике.
        # Здесь можно добавить дополнительные шаги.
        return document