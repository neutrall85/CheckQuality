from typing import Optional, Union
from datetime import date, datetime, timedelta
import re

def parse_date(value: Union[str, float, int, datetime, date, None]) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        if value < 1:   # Значения менее единицы (например, чистое время) – не дата
            return None
        try:
            base = date(1899, 12, 30)
            return base + timedelta(days=int(value))
        except (ValueError, OverflowError):
            return None
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        # 1. YYYY-MM-DD HH:MM:SS или YYYY-MM-DD
        m = re.match(r'(\d{4})-(\d{2})-(\d{2})( \d{2}:\d{2}:\d{2})?', value)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        # 2. DD.MM.YYYY с опциональной 'г' и возможным временем
        m = re.match(r'(\d{1,2})\.(\d{1,2})\.(\d{4})г?( \d{2}:\d{2}:\d{2})?', value)
        if m:
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1900 <= year <= 2100:
                return date(year, month, day)
        # 3. DD/MM/YYYY
        m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', value)
        if m:
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1900 <= year <= 2100:
                return date(year, month, day)
        # 4. Пробуем dateutil (если установлен)
        try:
            from dateutil.parser import parse
            dt = parse(value, dayfirst=True, yearfirst=False)
            return dt.date()
        except (ImportError, ValueError):
            pass
    return None