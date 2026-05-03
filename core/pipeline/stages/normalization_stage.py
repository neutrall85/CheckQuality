from typing import Optional
from core.interfaces.i_pipeline_stage import IPipelineStage
from core.models.document import Document


class NormalizationStage(IPipelineStage):
    """Этап дополнительной нормализации."""

    def process(self, document: Document) -> Optional[Document]:
        # Здесь может вызываться Normalizer.normalize() или другие трансформации
        return document