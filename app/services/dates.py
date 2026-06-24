# Модуль работы с датами
from datetime import date, datetime, timedelta


def normalize_date(value) -> date:
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, str):
        return datetime.fromisoformat(value).date()
    raise ValueError(f"Unsupported date value: {value!r}")


def subtract_calendar_days(value: date, days: int) -> date:
    return value - timedelta(days=days)


def add_calendar_days(value: date, days: int) -> date:
    return value + timedelta(days=days)
