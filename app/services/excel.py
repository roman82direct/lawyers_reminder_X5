import logging
from datetime import date, datetime, timedelta

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from app.config import ExcelConfig
from app.models import LicenseModel


class ExcelService:
    """Класс обработки xlsx-файла."""

    def __init__(self, xls_config: ExcelConfig):
        self.config = xls_config
        self.logger = logging.getLogger(__name__)

    def process_file(self):
        """Обрабатывает файл по тех. заданию."""
        wb = load_workbook(self.config.excel_path)

        if self.config.sheet_name:
            ws = wb[self.config.sheet_name]
        else:
            ws = wb.active

        headers = self._get_headers(ws)

        for column in (
            self.config.ability_date_column,
            self.config.renewal_date_column,
            self.config.reissue_date_column
        ):
            self._ensure_column(ws, headers, column_name=column)

        headers = self._get_headers(ws)
        self.logger.debug('NEW HEADERS: %s', headers)

        # получаем колонки
        license_name_col = headers[self.config.license_name_column]
        license_end_col = headers[self.config.license_end_column]
        ability_col = headers[self.config.ability_date_column]
        renewal_col = headers[self.config.renewal_date_column]
        reissue_col = headers[self.config.reissue_date_column]

        for row in range(2, ws.max_row + 1):
            license_name_cell = ws.cell(row, license_name_col)
            license_end_cell = ws.cell(row, license_end_col)
            ability_date_cell = ws.cell(row, ability_col)
            renewal_date_cell = ws.cell(row, renewal_col)
            reissue_date_cell = ws.cell(row, reissue_col)

            license_name = license_name_cell.value
            license_end = license_end_cell.value

            if not isinstance(license_end, datetime):
                self.logger.error('В строке %d, колонке %d '
                                  'неверный формат даты', row, license_end_col)
                continue

            ability_date = ability_date_cell.value

            if ability_date is not None:
                continue

            ability_date = license_end - timedelta(
                    days=self.config.days_before_submission
                    )

            renewal_date = (
                renewal_date_cell.value or
                license_end - timedelta(
                    days=self.config.deadline_for_submission
                )
            )
            reissue_date = (
                reissue_date_cell.value or
                renewal_date - timedelta(
                    days=self.config.deadline_for_submission
                )
            )
            # перезаписываем
            ability_date_cell.value = ability_date
            renewal_date_cell.value = renewal_date
            reissue_date_cell.value = reissue_date
            for cell in (ability_date_cell, renewal_date_cell,
                         reissue_date_cell):
                cell.number_format = self.config.date_format

        wb.save('data/copy.xlsx')

    def _get_headers(self, ws) -> dict[str, int]:
        """Извлекает названия колонок."""
        headers = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(1, col).value
            if value:
                headers[str(value).strip()] = col
        return headers

    def _ensure_column(
            self, ws, headers: dict[str, int], column_name: str
    ) -> None:
        """Добавляет новую колонку, если её нет."""
        if column_name not in headers:
            new_col = ws.max_column + 1
            ws.cell(1, new_col).value = column_name
            ws.cell(1, new_col).font = Font(bold=True)
            ws.cell(1, new_col).alignment = Alignment(horizontal='center',
                                                      vertical='center',
                                                      wrap_text=True,)
            col_letter = get_column_letter(new_col)
            ws.column_dimensions[col_letter].width = self.config.column_width

            self.logger.debug('Add column: %s', column_name)
