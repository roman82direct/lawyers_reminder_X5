from dataclasses import dataclass
from datetime import date


@dataclass
class LicenseModel:
    """Класс объекта строки в таблице."""

    name: str
    license_end_date: date
    ability_date: date | None = None
    renewal_date: date | None = None
    reissue_date: date | None = None
