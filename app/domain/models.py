from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True)
class LicenseObject:
    """Класс объекта строки в таблице."""

    object_name: str
    license_end_date: date
    ability_to_license_date: date | None = None
    renewal_date: date | None = None
    reissue_date: date | None = None
