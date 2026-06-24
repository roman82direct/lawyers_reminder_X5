from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True)
class LicenseObject:
    """Класс объекта строки в таблице."""

    row_number: int
    object_name: str
    object_address: str
    license_end_date: date
    submission_date: date | None = None
