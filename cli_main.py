import sys
from core.loaders.excel_loader import ExcelLoader
from core.pipeline.pipeline import Pipeline
from core.pipeline.stages.date_filter_stage import DateFilterStage
from core.pipeline.stages.validation_stage import ValidationStage
from core.pipeline.stages.normalization_stage import NormalizationStage
from core.services.document_processor import DocumentProcessor
from core.validators.data_validator import DataValidator
from core.validators.schema_validator import SchemaValidator
from core.processors.dictionary_manager import DictionaryManager
from utils.logger import setup_logger

logger = setup_logger("ExcelReporter")

def main():
    if len(sys.argv) < 2:
        print("Использование: python cli_main.py file1.xlsx [file2.xlsx ...]")
        return

    filenames = sys.argv[1:]

    # Инициализация словаря типов
    dict_mgr = DictionaryManager()

    # Загрузчик
    loader = ExcelLoader(normalize_types=dict_mgr.normalize)

    # Конвейер
    pipeline = Pipeline(stages=[
        ValidationStage(validators=[DataValidator(), SchemaValidator()]),
        NormalizationStage(),
        DateFilterStage(start_date=None, end_date=None)  # без фильтрации
    ])

    # Процессор
    processor = DocumentProcessor(
        loader=loader,
        pipeline=pipeline,
        deduplication_key="document_number"
    )

    def progress_callback(current, total, filename):
        print(f"[{current}/{total}] Обрабатывается: {filename}")

    def log_callback(message, level="INFO"):
        print(f"[{level}] {message}")
        if level == "ERROR":
            logger.error(message)
        elif level == "WARNING":
            logger.warning(message)
        else:
            logger.info(message)

    print("Начало обработки...")
    stats = processor.process_files(filenames, progress_callback, log_callback)

    print("\n=== РЕЗУЛЬТАТЫ ===")
    print(f"Всего документов (после дедупликации): {stats.total_docs}")
    print(f"Документов с ошибками: {stats.docs_with_errors}")
    print(f"Общее количество листов А4: {stats.total_a4}")
    print(f"Ошибок категории 1: {stats.total_errors_cat1}")
    print(f"Ошибок категории 2: {stats.total_errors_cat2}")

    print("\nПо типам документов:")
    for doc_type, data in sorted(stats.by_type.items()):
        print(f"  {doc_type}: кол-во={data['count']}, ошибки1={data['errors1']}, ошибки2={data['errors2']}, А4={data['a4']}")

    print("\nПо разработчикам:")
    for dev, data in sorted(stats.by_developer.items()):
        print(f"  {dev}: кол-во док.={data['count']}, ошибки1={data['errors1']}, ошибки2={data['errors2']}, А4={data['a4']}")

    print("\nПо месяцам:")
    for month, data in sorted(stats.by_month.items()):
        print(f"  {month}: кол-во док.={data['count']}, ошибки1={data['errors1']}, ошибки2={data['errors2']}")

if __name__ == "__main__":
    main()