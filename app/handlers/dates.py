# Модуль работы с датами
from datetime import date, datetime, timedelta





def subtract_calendar_days(value: date, days: int) -> date:
    return value - timedelta(days=days)


def add_calendar_days(value: date, days: int) -> date:
    return value + timedelta(days=days)
