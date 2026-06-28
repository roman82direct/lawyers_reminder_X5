import logging
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class ExcelConfig:
    """Класс настроек для обработки xlsx-файла."""

    excel_path: Path = field(default_factory=lambda: Path(
        os.getenv('EXCEL_PATH', 'data/submissions.xlsx')
    ))
    sheet_name: str | None = field(
        default_factory=lambda: os.getenv('SHEET_NAME') or None)
    header_position: int = 1
    license_name_column: str = 'Объект'
    license_end_column: str = 'Дата окончания лицензии'
    ability_date_column: str = 'Дата когда возможно подавать на лицензирование'
    renewal_date_column: str = 'Дата подачи заявления на продление'
    reissue_date_column: str = 'Крайний день подачи на переоформление'
    days_before_submission: int = 90
    deadline_for_submission: int = 16
    notify_before_days: int = 3
    date_format: str = 'DD.MM.YYYY'
    column_width: int = 30


def setup_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s %(levelname)s %(name)s %(message)s',
        handlers=(
            logging.FileHandler('logs/app.log', encoding='utf-8'),
            logging.StreamHandler(sys.stdout),
        ),
    )
