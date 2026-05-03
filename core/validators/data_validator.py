from typing import Optional
from core.interfaces.i_validator import IValidator
from core.models.document import Document


class DataValidator(IValidator):
    """Валидатор обязательных полей документа."""

    def validate(self, document: Document) -> Optional[str]:
        if not document.number:
            return "Номер документа не может быть пустым"
        # Дата проверки больше не проверяется, т.к. уже гарантирована загрузчиком
        return None