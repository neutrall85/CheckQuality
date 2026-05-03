from dataclasses import dataclass, field
from typing import List, Optional
from datetime import date


@dataclass
class ConfigModel:
    """Модель конфигурации приложения."""
    theme: str = "light"
    last_output_path: str = ""
    deduplication_key: str = "document_number"
    developer_separators: List[str] = field(default_factory=lambda: [",", ";", "и", "\n", "\\s{2,}"])
    memory_warning_threshold: int = 500
    use_custom_template: bool = False
    custom_template_path: Optional[str] = None
    anonymize_names: bool = False
    open_report: bool = True
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    selected_developers: Optional[List[str]] = None