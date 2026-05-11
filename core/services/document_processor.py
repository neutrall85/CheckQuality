import os
from typing import List, Optional, Callable
from core.interfaces.i_loader import IDataLoader
from core.pipeline.pipeline import Pipeline
from core.processors.aggregator import Aggregator
from core.models.statistics import Statistics
from utils.memory_utils import force_garbage_collection
import logging

logger = logging.getLogger(__name__)


class DocumentProcessor:
    """Сервисный слой: координирует загрузку, конвейер и агрегацию."""

    def __init__(self,
                 loader: IDataLoader,
                 pipeline: Pipeline,
                 selected_developers: Optional[List[str]] = None):
        self.loader = loader
        self.pipeline = pipeline
        self.aggregator = Aggregator()
        self.selected_developers = selected_developers

    def process_files(self,
                      file_paths: List[str],
                      progress_callback: Optional[Callable] = None,
                      log_callback: Optional[Callable] = None
                      ) -> Statistics:
        all_docs = []
        total = len(file_paths)
        prefix_counts = {}

        for idx, path in enumerate(file_paths):
            if progress_callback:
                progress_callback(idx + 1, total, path)

            base_name = os.path.splitext(os.path.basename(path))[0]
            if '_' in base_name:
                prefix = base_name.split('_')[0]
            else:
                prefix = base_name

            file_doc_count = 0

            try:
                for doc in self.loader.load(path):
                    processed_doc = self.pipeline.execute(doc)
                    if processed_doc is not None:
                        all_docs.append(processed_doc)
                        file_doc_count += 1
            except Exception as e:
                msg = f"Ошибка при обработке {path}: {e}"
                if log_callback:
                    log_callback(msg, level="ERROR")
                else:
                    logger.error(msg)
                continue

            prefix_counts[prefix] = prefix_counts.get(prefix, 0) + file_doc_count

            log_msg = f"Файл {path}: обработано документов — {file_doc_count} (группа '{prefix}')"
            if log_callback:
                log_callback(log_msg, level="INFO")
            else:
                logger.info(log_msg)

            if len(all_docs) % 50 == 0:
                force_garbage_collection()

        stats = self.aggregator.aggregate(all_docs,
                                         selected_developers=self.selected_developers)
        stats.duplicates = []
        stats.docs_by_file_prefix = prefix_counts
        force_garbage_collection()

        total_msg = f"Всего обработано документов из {len(file_paths)} файлов: {len(all_docs)}"
        if log_callback:
            log_callback(total_msg, level="INFO")
        else:
            logger.info(total_msg)

        return stats