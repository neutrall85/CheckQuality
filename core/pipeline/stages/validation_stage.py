from typing import Optional, List
from core.interfaces.i_pipeline_stage import IPipelineStage
from core.interfaces.i_validator import IValidator
from core.models.document import Document


class ValidationStage(IPipelineStage):
    """Этап валидации документа."""

    def __init__(self, validators: Optional[List[IValidator]] = None):
        self.validators = validators or []

    def process(self, document: Document) -> Optional[Document]:
        for validator in self.validators:
            error = validator.validate(document)
            if error:
                # В реальном приложении – логирование WARNING
                return None  # отбрасываем документ
        return document