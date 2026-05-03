from typing import List, Optional
from core.interfaces.i_pipeline_stage import IPipelineStage
from core.models.document import Document


class Pipeline:
    """Конвейер последовательной обработки документов."""

    def __init__(self, stages: Optional[List[IPipelineStage]] = None):
        self.stages = stages or []

    def execute(self, document: Document) -> Optional[Document]:
        for stage in self.stages:
            document = stage.process(document)
            if document is None:
                break
        return document