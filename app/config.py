from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    """Класс настроек приложения."""

    excel_path: Path = Path('data/submissions.xlsx')
    sheet_name: str | None = None
    license_end_column: str = 'Дата окончания лицензии'
    submission_date_column: str = 'Дата подачи заявления на продление'
    days_before_submission: int = 15
    notify_before_days: int = 10
    run_hour: int = 0
    run_minute: int = 0
