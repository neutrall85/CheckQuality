from typing import Optional, List
from core.interfaces.i_pipeline_stage import IPipelineStage
from core.models.document import Document


class DeveloperFilterStage(IPipelineStage):
    """Отбрасывает документы, в которых нет ни одного из указанных разработчиков."""

    def __init__(self, developers: List[str]):
        self.developers = set(developers) if developers else set()

    def process(self, document: Document) -> Optional[Document]:
        # Если список разработчиков пуст – пропускаем все документы
        if not self.developers:
            return document
        if any(dev in self.developers for dev in document.developers):
            return document
        return None