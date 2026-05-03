from typing import Optional
from datetime import date
from core.interfaces.i_pipeline_stage import IPipelineStage
from core.models.document import Document


class DateFilterStage(IPipelineStage):
    """Фильтрация по диапазону дат проверки."""

    def __init__(self, start_date: Optional[date] = None, end_date: Optional[date] = None):
        self.start_date = start_date
        self.end_date = end_date

    def process(self, document: Document) -> Optional[Document]:
        if document.check_date is None:
            return document   # не отбрасываем
        if self.start_date and document.check_date < self.start_date:
            return None
        if self.end_date and document.check_date > self.end_date:
            return None
        return document