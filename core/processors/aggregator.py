from typing import Dict, List, Optional
from datetime import date
from core.models.document import Document
from core.models.statistics import Statistics


class Aggregator:
    """Агрегирует список документов в объект Statistics.

    selected_developers влияет только на группировку by_developer,
    но НЕ фильтрует общие показатели (total_docs, total_errors и т.д.).
    """

    def aggregate(self, documents: List[Document],
                  selected_developers: Optional[List[str]] = None) -> Statistics:
        stats = Statistics()
        selected_set = set(selected_developers) if selected_developers else None
        stats.total_docs = len(documents)

        # Для вычисления фактических границ дат
        all_dates = []

        for doc in documents:
            if doc.has_errors():
                stats.docs_with_errors += 1
                stats.total_a4_errors += doc.a4_count
            stats.total_a4 += doc.a4_count
            stats.total_errors_cat1 += doc.errors_cat1
            stats.total_errors_cat2 += doc.errors_cat2
            if doc.errors_cat1 > stats.max_errors_cat1:
                stats.max_errors_cat1 = doc.errors_cat1
            if doc.errors_cat2 > stats.max_errors_cat2:
                stats.max_errors_cat2 = doc.errors_cat2

            # Подсчёт для круговой диаграммы (старые поля, оставим для совместимости)
            if doc.errors_cat1 > 0:
                stats.total_docs_with_cat1 += 1
            if doc.errors_cat2 > 0:
                stats.total_docs_with_cat2 += 1

            # Новые взаимоисключающие категории для круговой диаграммы
            if doc.errors_cat1 > 0 and doc.errors_cat2 > 0:
                stats.total_docs_with_both += 1
            elif doc.errors_cat1 > 0:
                stats.total_docs_with_only_cat1 += 1
            elif doc.errors_cat2 > 0:
                stats.total_docs_with_only_cat2 += 1

            # По типам документов
            type_stats = stats.by_type.setdefault(doc.doc_type, {
                "count": 0, "errors1": 0, "errors2": 0, "a4": 0
            })
            type_stats["count"] += 1
            type_stats["errors1"] += doc.errors_cat1
            type_stats["errors2"] += doc.errors_cat2
            type_stats["a4"] += doc.a4_count

            # По разработчикам (с учётом фильтра, если задан)
            for dev in doc.developers:
                if selected_set is not None and dev not in selected_set:
                    continue
                dev_stats = stats.by_developer.setdefault(dev, {
                    "count": 0, "errors1": 0, "errors2": 0, "a4": 0
                })
                dev_stats["count"] += 1
                dev_stats["errors1"] += doc.errors_cat1
                dev_stats["errors2"] += doc.errors_cat2
                dev_stats["a4"] += doc.a4_count

            # По месяцам (дата проверки)
            if doc.check_date:
                month_key = doc.check_date.strftime('%Y-%m')
                month_stats = stats.by_month.setdefault(month_key, {
                    "count": 0, "errors1": 0, "errors2": 0, "a4": 0
                })
                month_stats["count"] += 1
                month_stats["errors1"] += doc.errors_cat1
                month_stats["errors2"] += doc.errors_cat2
                month_stats["a4"] += doc.a4_count

            # Собираем все даты документа для вычисления фактических границ
            if doc.check_date:
                all_dates.append(doc.check_date)
            elif doc.receipt_date:
                all_dates.append(doc.receipt_date)

        # Вычисляем фактические границы
        if all_dates:
            stats.actual_start_date = min(all_dates)
            stats.actual_end_date = max(all_dates)

        return stats