import json
import os
import re
from typing import Dict, Optional, List


class DictionaryManager:
    """Управление словарём типов документов."""

    DEFAULT_DICT_PATH = os.path.join(
        os.path.expanduser("~"), "AppData", "Roaming", "ExcelReporter", "dictionary.json"
    )

    def __init__(self, dict_path: Optional[str] = None):
        self.dict_path = dict_path or self.DEFAULT_DICT_PATH
        self.rules: List[Dict] = []
        self.unknown_handling = "as_is"
        self.auto_add_unknown = True
        self._dirty = False
        self.load()

    def load(self):
        if not os.path.exists(self.dict_path):
            self._save_default()
        try:
            with open(self.dict_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.rules = data.get("rules", [])
            self.unknown_handling = data.get("unknown_handling", "as_is")
            self.auto_add_unknown = data.get("auto_add_unknown", True)
            self.rules.sort(key=lambda r: r.get("priority", 0), reverse=True)
        except Exception:
            self.rules = []

    def _save_default(self):
        default = {
            "version": "1.0",
            "last_updated": "2026-01-01T00:00:00",
            "rules": [],
            "unknown_handling": "as_is",
            "auto_add_unknown": True
        }
        os.makedirs(os.path.dirname(self.dict_path), exist_ok=True)
        with open(self.dict_path, 'w', encoding='utf-8') as f:
            json.dump(default, f, indent=4, ensure_ascii=False)

    def normalize(self, source_type: str) -> str:
        """Применяет словарь и возвращает каноническое имя типа."""
        source_type = source_type.strip()
        if not source_type:
            return source_type

        for rule in self.rules:
            pattern = rule["source_pattern"]
            is_regex = rule.get("is_regex", False)
            case_sensitive = rule.get("case_sensitive", False)
            flags = 0 if case_sensitive else re.IGNORECASE

            if is_regex:
                if re.match(pattern, source_type, flags):
                    return rule["canonical_name"]
            else:
                if case_sensitive:
                    if source_type == pattern:
                        return rule["canonical_name"]
                else:
                    if source_type.lower() == pattern.lower():
                        return rule["canonical_name"]

        # Неизвестный тип
        if self.auto_add_unknown:
            self.rules.append({
                "source_pattern": source_type,
                "canonical_name": source_type,
                "is_regex": False,
                "case_sensitive": False,
                "priority": 0,
                "auto_added": True
            })
            self._dirty = True

        if self.unknown_handling == "as_is":
            return source_type
        elif self.unknown_handling == "skip":
            return ""
        else:
            return source_type

    def save_if_needed(self):
        """Сохраняет правила, если были изменения."""
        if self._dirty:
            self._save_rules()
            self._dirty = False

    def _save_rules(self):
        data = {
            "version": "1.0",
            "last_updated": "2026-01-01T00:00:00",
            "rules": self.rules,
            "unknown_handling": self.unknown_handling,
            "auto_add_unknown": self.auto_add_unknown
        }
        with open(self.dict_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)