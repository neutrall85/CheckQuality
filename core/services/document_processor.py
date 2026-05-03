from typing import List, Optional, Callable
from core.interfaces.i_loader import IDataLoader
from core.pipeline.pipeline import Pipeline
from core.processors.deduplicator import Deduplicator
from core.processors.aggregator import Aggregator
from core.models.statistics import Statistics


class DocumentProcessor:
    """Сервисный слой: координирует загрузку, конвейер и агрегацию."""

    def __init__(self,
                 loader: IDataLoader,
                 pipeline: Pipeline,
                 deduplication_key: str = "document_number",
                 selected_developers: Optional[List[str]] = None):
        self.loader = loader
        self.pipeline = pipeline
        self.deduplicator = Deduplicator(key_mode=deduplication_key)
        self.aggregator = Aggregator()
        self.selected_developers = selected_developers

    def process_files(self,
                      file_paths: List[str],
                      progress_callback: Optional[Callable] = None,
                      log_callback: Optional[Callable] = None
                      ) -> Statistics:
        all_docs = []
        total = len(file_paths)
        for idx, path in enumerate(file_paths):
            if progress_callback:
                progress_callback(idx + 1, total, path)
            try:
                for doc in self.loader.load(path):
                    processed_doc = self.pipeline.execute(doc)
                    if processed_doc is not None:
                        all_docs.append(processed_doc)
            except Exception as e:
                if log_callback:
                    log_callback(f"Ошибка при обработке {path}: {e}", level="ERROR")
                continue

        # Дедупликация с информацией о дубликатах
        deduped, dup_info = self.deduplicator.deduplicate(all_docs)

        # Логируем информацию о дубликатах
        if log_callback:
            for info in dup_info:
                log_callback(
                    f"Дедупликация: ключ '{info['key']}' объединил {info['count']} записей. "
                    f"Типы: {', '.join(info['types'])}, "
                    f"Разработчики: {', '.join(info['developers'])}",
                    level="INFO"
                )

        # Агрегация с учётом фильтра по разработчикам
        stats = self.aggregator.aggregate(list(deduped.values()),
                                         selected_developers=self.selected_developers)
        stats.duplicates = dup_info
        return stats