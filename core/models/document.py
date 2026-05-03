from dataclasses import dataclass, field
from typing import List, Optional
from datetime import date


@dataclass
class Document:
    doc_type: str
    number: str
    developers: List[str] = field(default_factory=list)
    receipt_date: Optional[date] = None
    a4_count: int = 0
    errors_cat1: int = 0
    errors_cat2: int = 0
    check_date: Optional[date] = None

    def has_errors(self) -> bool:
        return self.errors_cat1 > 0 or self.errors_cat2 > 0

    @property
    def total_errors(self) -> int:
        return self.errors_cat1 + self.errors_cat2