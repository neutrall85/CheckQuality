from typing import Optional
from core.interfaces.i_validator import IValidator
from core.models.document import Document


class SchemaValidator(IValidator):
    """Валидатор типов полей."""

    def validate(self, document: Document) -> Optional[str]:
        if not isinstance(document.a4_count, int) or document.a4_count < 0:
            return "Некорректное количество форматов А4"
        if not isinstance(document.errors_cat1, int) or document.errors_cat1 < 0:
            return "Некорректное количество ошибок категории 1"
        if not isinstance(document.errors_cat2, int) or document.errors_cat2 < 0:
            return "Некорректное количество ошибок категории 2"
        return None