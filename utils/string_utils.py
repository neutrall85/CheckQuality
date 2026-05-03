import re


def normalize_fio(raw: str) -> str:
    """Очищает ФИО, оставляя только фамилию (первое слово), приводит к Title Case."""
    if not raw:
        return ''
    raw = raw.strip()
    # Удаляем возможные инициалы (с точками)
    raw = re.sub(r'\s+[А-ЯЁ]\s*\.\s*[А-ЯЁ]?\.?', '', raw, flags=re.IGNORECASE)
    parts = raw.split()
    surname = parts[0] if parts else raw
    # Приведение к Title Case (только первая заглавная)
    surname = surname[0].upper() + surname[1:].lower() if surname else ''
    return surname