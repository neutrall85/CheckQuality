from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
from datetime import date

@dataclass
class Statistics:
    total_docs: int = 0
    docs_with_errors: int = 0
    total_a4: int = 0
    total_a4_errors: int = 0
    total_errors_cat1: int = 0
    total_errors_cat2: int = 0
    max_errors_cat1: int = 0
    max_errors_cat2: int = 0

    # Новые поля для круговой диаграммы
    total_docs_with_cat1: int = 0
    total_docs_with_cat2: int = 0

    by_type: Dict[str, Dict[str, int]] = field(default_factory=dict)
    by_developer: Dict[str, Dict[str, int]] = field(default_factory=dict)
    by_month: Dict[str, Dict[str, int]] = field(default_factory=dict)

    duplicates: List[Dict] = field(default_factory=list)

    # Фактические границы дат, вычисляются агрегатором
    actual_start_date: Optional[date] = None
    actual_end_date: Optional[date] = None