#!/usr/bin/env python3
"""
Скрипт формирования сводных отчётов (Excel и PowerPoint).

Поведение по умолчанию (без аргументов):
  - Запрашивает даты начала и конца периода.
  - Анализирует все Excel-файлы из папки test_data и выводит список разработчиков.
  - Запрашивает номера разработчиков через пробел/запятую (Enter — все).
  - Строит отчёты только по выбранным разработчикам и сохраняет в папку output.

Примеры аргументов:
  --start-date 01.01.2026 --end-date 31.03.2026 --developers Иванов,Петров
  --excel отчет.xlsx --pptx презентация.pptx
"""

import sys
import os
import glob
from datetime import datetime, date
from typing import List, Optional, Set

from core.loaders.excel_loader import ExcelLoader
from core.processors.dictionary_manager import DictionaryManager
from core.pipeline.pipeline import Pipeline
from core.pipeline.stages.validation_stage import ValidationStage
from core.pipeline.stages.normalization_stage import NormalizationStage
from core.pipeline.stages.date_filter_stage import DateFilterStage
from core.pipeline.stages.developer_filter_stage import DeveloperFilterStage
from core.validators.data_validator import DataValidator
from core.validators.schema_validator import SchemaValidator
from core.services.document_processor import DocumentProcessor
from core.models.config import ConfigModel
from core.models.document import Document

from reporting.excel_single_sheet import ExcelSingleSheetReport
from reporting.powerpoint_report import PowerpointReport
from reporting.powerpoint_template_manager import PowerpointTemplateManager
from reporting.chart_builder import ChartBuilder

from utils.logger import setup_logger
from utils.string_utils import normalize_fio

logger = setup_logger("ReportGenerator")

DEFAULT_INPUT_DIR = "test_data"
DEFAULT_OUTPUT_DIR = "output"


def parse_date_arg(date_str: str) -> Optional[date]:
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str, '%d.%m.%Y').date()
    except ValueError:
        print(f"Предупреждение: Неверный формат даты '{date_str}'. Ожидается ДД.ММ.ГГГГ", file=sys.stderr)
        return None


def parse_developers_arg(dev_str: str) -> List[str]:
    """Разбирает строку с фамилиями через запятую и нормализует."""
    if not dev_str:
        return []
    parts = [p.strip() for p in dev_str.split(',') if p.strip()]
    normalized = []
    for p in parts:
        norm = normalize_fio(p)
        if norm:
            normalized.append(norm)
    return normalized


def collect_files(paths: List[str]) -> List[str]:
    files = []
    for p in paths:
        if os.path.isfile(p):
            if p.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                # Игнорируем временные файлы Excel (~$...)
                if os.path.basename(p).startswith('~$'):
                    print(f"Предупреждение: '{p}' является временным файлом Excel, пропускается", file=sys.stderr)
                    continue
                files.append(os.path.abspath(p))
            else:
                print(f"Предупреждение: '{p}' не является Excel-файлом, пропускается", file=sys.stderr)
        elif os.path.isdir(p):
            for ext in ('*.xlsx', '*.xls', '*.xlsm'):
                for f in glob.glob(os.path.join(p, ext)):
                    if os.path.basename(f).startswith('~$'):
                        print(f"Предупреждение: '{f}' является временным файлом Excel, пропускается", file=sys.stderr)
                        continue
                    files.append(f)
        else:
            print(f"Ошибка: '{p}' не найден", file=sys.stderr)
    return sorted(set(files))


def ensure_dir(dir_path: str):
    if dir_path and not os.path.exists(dir_path):
        os.makedirs(dir_path, exist_ok=True)


def collect_all_developers(file_paths: List[str],
                           loader: ExcelLoader,
                           pipeline: Pipeline) -> List[str]:
    """Загружает документы из файлов, пропускает через базовый pipeline
    (без фильтрации по разработчикам) и возвращает отсортированный список
    уникальных нормализованных фамилий."""
    developers: Set[str] = set()
    for path in file_paths:
        try:
            for doc in loader.load(path):
                processed = pipeline.execute(doc)
                if processed is not None:
                    for dev in processed.developers:
                        developers.add(dev)
        except Exception as e:
            logger.warning(f"Ошибка при сканировании {path}: {e}")
    return sorted(developers)


def main():
    args = sys.argv[1:]
    use_defaults = (len(args) == 0)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    input_files = []
    default_excel = f'report_{timestamp}.xlsx'
    default_pptx = f'report_{timestamp}.pptx'

    if use_defaults:
        output_excel = os.path.join(DEFAULT_OUTPUT_DIR, default_excel)
        output_pptx = os.path.join(DEFAULT_OUTPUT_DIR, default_pptx)
    else:
        output_excel = default_excel
        output_pptx = default_pptx

    start_date = None
    end_date = None
    template_path = None
    generate_excel = True
    generate_pptx = True
    selected_developers = None

    if use_defaults:
        print("Выбор периода для отчёта (ДД.ММ.ГГГГ)")
        start_input = input("Дата начала (Enter — весь период): ").strip()
        if start_input:
            start_date = parse_date_arg(start_input)
        end_input = input("Дата окончания (Enter — весь период): ").strip()
        if end_input:
            end_date = parse_date_arg(end_input)

        if not os.path.isdir(DEFAULT_INPUT_DIR):
            print(f"Папка '{DEFAULT_INPUT_DIR}' не найдена. Создайте её и поместите туда Excel-файлы.", file=sys.stderr)
            input("Нажмите Enter для выхода...")
            return
        input_files = collect_files([DEFAULT_INPUT_DIR])
        if not input_files:
            print(f"В папке '{DEFAULT_INPUT_DIR}' нет подходящих файлов .xlsx/.xls/.xlsm", file=sys.stderr)
            input("Нажмите Enter для выхода...")
            return
        ensure_dir(DEFAULT_OUTPUT_DIR)

        # --- Первый проход: сбор всех разработчиков без фильтрации ---
        print("\nАнализ файлов для получения списка разработчиков...")
        dict_mgr = DictionaryManager()
        loader = ExcelLoader(normalize_types=dict_mgr.normalize)
        base_pipeline = Pipeline(stages=[
            ValidationStage(validators=[DataValidator(), SchemaValidator()]),
            NormalizationStage(),
            DateFilterStage(start_date=start_date, end_date=end_date)
        ])
        all_devs = collect_all_developers(input_files, loader, base_pipeline)

        if not all_devs:
            print("Не найдено ни одной фамилии разработчика в документах.")
            return

        print(f"\nНайдено {len(all_devs)} разработчиков:")
        for idx, dev in enumerate(all_devs, start=1):
            print(f"  {idx:3d}. {dev}")

        choice = input("\nВведите номера разработчиков через пробел/запятую (Enter — все): ").strip()
        if choice:
            raw_nums = choice.replace(',', ' ').split()
            selected_indices = set()
            for token in raw_nums:
                try:
                    num = int(token)
                    if 1 <= num <= len(all_devs):
                        selected_indices.add(num - 1)
                    else:
                        print(f"Номер {num} вне диапазона, игнорируется.", file=sys.stderr)
                except ValueError:
                    print(f"'{token}' не похоже на номер, игнорируется.", file=sys.stderr)
            if selected_indices:
                selected_developers = [all_devs[i] for i in sorted(selected_indices)]
                print(f"Выбраны: {', '.join(selected_developers)}")
            else:
                print("Не выбрано ни одного разработчика, будут использованы все.")
                selected_developers = None
        else:
            print("Выбраны все разработчики.")
            selected_developers = None

    else:
        # Режим командной строки
        if '--help' in args or '-h' in args:
            print(__doc__)
            return

        i = 0
        while i < len(args):
            arg = args[i]
            if arg == '--excel':
                if i + 1 < len(args):
                    output_excel = args[i + 1]
                    i += 1
            elif arg == '--pptx':
                if i + 1 < len(args):
                    output_pptx = args[i + 1]
                    i += 1
            elif arg == '--start-date':
                if i + 1 < len(args):
                    start_date = parse_date_arg(args[i + 1])
                    i += 1
            elif arg == '--end-date':
                if i + 1 < len(args):
                    end_date = parse_date_arg(args[i + 1])
                    i += 1
            elif arg == '--template':
                if i + 1 < len(args):
                    template_path = args[i + 1]
                    if not os.path.exists(template_path):
                        print(f"Предупреждение: шаблон '{template_path}' не найден, используется встроенный",
                              file=sys.stderr)
                        template_path = None
                    i += 1
            elif arg == '--developers':
                if i + 1 < len(args):
                    selected_developers = parse_developers_arg(args[i + 1])
                    if not selected_developers:
                        print("Предупреждение: указаны пустые или нераспознанные фамилии разработчиков.",
                              file=sys.stderr)
                    i += 1
            elif arg == '--no-excel':
                generate_excel = False
            elif arg == '--no-pptx':
                generate_pptx = False
            elif arg.startswith('--'):
                print(f"Неизвестная опция: {arg}", file=sys.stderr)
                print(__doc__)
                return
            else:
                input_files.append(arg)
            i += 1

        if not input_files:
            print("Ошибка: не указаны входные файлы или папки", file=sys.stderr)
            return

        input_files = collect_files(input_files)
        if not input_files:
            print("Нет подходящих файлов для обработки", file=sys.stderr)
            return

    print(f"\nНайдено файлов для обработки: {len(input_files)}")
    for f in input_files:
        print(f"  - {f}")

    # Общая конфигурация
    config = ConfigModel(
        theme="light",
        anonymize_names=False,
        open_report=False,
        period_start=start_date,
        period_end=end_date,
        selected_developers=selected_developers
    )
    if template_path:
        config.use_custom_template = True
        config.custom_template_path = template_path

    dict_mgr = DictionaryManager()

    def progress_callback(current, total, filename):
        print(f"[{current}/{total}] Обработка: {os.path.basename(filename)}")

    def log_callback(message, level="INFO"):
        if level in ("ERROR", "WARNING"):
            print(f"[{level}] {message}", file=sys.stderr)
        else:
            print(f"[{level}] {message}")
        if level == "ERROR":
            logger.error(message)
        elif level == "WARNING":
            logger.warning(message)
        else:
            logger.info(message)

    loader = ExcelLoader(normalize_types=dict_mgr.normalize, log_callback=log_callback)

    # Собираем pipeline с возможной фильтрацией по разработчикам
    stages = [
        ValidationStage(validators=[DataValidator(), SchemaValidator()]),
        NormalizationStage()
    ]
    if selected_developers:
        stages.append(DeveloperFilterStage(selected_developers))
    stages.append(DateFilterStage(start_date=start_date, end_date=end_date))

    pipeline = Pipeline(stages=stages)

    processor = DocumentProcessor(
        loader=loader,
        pipeline=pipeline,
        deduplication_key=config.deduplication_key,
        selected_developers=selected_developers
    )

    print("\nНачало обработки...")
    stats = processor.process_files(input_files, progress_callback, log_callback)

    if stats.total_docs == 0:
        print("Нет данных для отчёта после фильтрации и валидации.")
        if use_defaults:
            input("Нажмите Enter для выхода...")
        return

    print(f"Обработка завершена. Всего документов после дедупликации: {stats.total_docs}")

    if stats.duplicates:
        print("\nИнформация об объединённых документах:")
        for dup in stats.duplicates:
            print(f"  Ключ: {dup['key']} – объединено {dup['count']} записей")
            print(f"    Типы: {', '.join(dup['types'])}")
            print(f"    Разработчики: {', '.join(dup['developers'])}")
            print(f"    Даты проверок (или поступления): {', '.join(dup['dates'])}")
        print()
    else:
        print("Дубликаты не обнаружены.\n")

    chart_builder = ChartBuilder()

    if generate_excel:
        print(f"\nСоздание Excel-отчёта: {output_excel}")
        excel_reporter = ExcelSingleSheetReport()
        try:
            excel_reporter.generate(stats, config, output_excel)
            print(f"Excel-отчёт успешно сохранён: {os.path.abspath(output_excel)}")
        except Exception as e:
            logger.exception("Ошибка создания Excel-отчёта")
            print(f"Ошибка создания Excel-отчёта: {e}", file=sys.stderr)

    if generate_pptx:
        print(f"\nСоздание PowerPoint-отчёта: {output_pptx}")
        pptx_reporter = PowerpointReport(
            template_manager=PowerpointTemplateManager(),
            chart_builder=chart_builder
        )
        try:
            pptx_reporter.generate(stats, config, output_pptx)
            print(f"PowerPoint-отчёт успешно сохранён: {os.path.abspath(output_pptx)}")
        except Exception as e:
            logger.exception("Ошибка создания PowerPoint-отчёта")
            print(f"Ошибка создания PowerPoint-отчёта: {e}", file=sys.stderr)

    chart_builder.cleanup()

    print("\nГотово.")
    if use_defaults:
        input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано пользователем.")