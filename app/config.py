import logging
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ExcelConfig:
    """Класс настроек для обработки xlsx-файла."""

    excel_path: Path = Path('data/submissions.xlsx')
    sheet_name: str | None = None
    license_end_column: str = 'Дата окончания лицензии'
    ability_to_license_date: str = ('Дата когда возможно подавать '
                                    'на лицензирование')
    renewal_date_column: str = 'Дата подачи заявления на продление'
    reissue_date_column: str = 'Крайний день подачи на переоформление'
    days_before_submission: int = 90
    deadline_for_submission: int = 16
    notify_before_days: int = 3


def setup_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s %(levelname)s %(name)s %(message)s',
        handlers=(
            logging.FileHandler('logs/app.log', encoding='utf-8'),
            logging.StreamHandler(),
        ),
    )
