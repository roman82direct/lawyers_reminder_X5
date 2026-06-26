import logging
from datetime import date, datetime, timedelta

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from workalendar.europe import Russia

from app.config import ExcelConfig


class ExcelService:
    """Класс обработки xlsx-файла."""

    def __init__(self, xls_config: ExcelConfig):
        self.config = xls_config
        thin = Side(border_style='thin', color='000000')
        self.cell_border = Border(left=thin, right=thin,
                                  top=thin, bottom=thin,)
        self.calendar = Russia()
        self.logger = logging.getLogger(__name__)

    def process_file(self):
        """Обрабатывает файл по тех. заданию."""
        wb = load_workbook(self.config.excel_path)
        ws = (wb[self.config.sheet_name] if self.config.sheet_name
              else wb.active)

        for column in (
            self.config.ability_date_column,
            self.config.renewal_date_column,
            self.config.reissue_date_column
        ):
            self._ensure_column(ws, self._get_headers(ws), column_name=column)

        headers = self._get_headers(ws)

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
            license_end = self._normalize_date(license_end_cell.value,
                                               row, license_end_col)
            if license_end is None:
                continue

            ability_date = self._normalize_date(ability_date_cell.value,
                                                row, ability_col)
            renewal_date = self._normalize_date(renewal_date_cell.value,
                                                row, renewal_col)
            reissue_date = self._normalize_date(reissue_date_cell.value,
                                                row, reissue_col)

            if ability_date is None:
                ability_date = license_end - timedelta(
                    days=self.config.days_before_submission
                )

            if renewal_date is None:
                renewal_date = self._subtract_working_days(
                    license_end, self.config.deadline_for_submission
                )

            if reissue_date is None:
                reissue_date = self._subtract_working_days(
                    renewal_date, self.config.deadline_for_submission
                )

            ability_date_cell.value = ability_date
            renewal_date_cell.value = renewal_date
            reissue_date_cell.value = reissue_date

            for cell in (ability_date_cell, renewal_date_cell,
                         reissue_date_cell):
                cell.number_format = self.config.date_format
                cell.border = self.cell_border

        # filename = datetime.now().strftime('%Y-%m-%d_%H-%M')
        wb.save('data/_submissions.xlsx')

    def _get_headers(self, ws) -> dict[str, int]:
        """Извлекает названия колонок."""
        headers = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(self.config.header_position, col).value
            if value:
                headers[str(value).strip()] = col
        return headers

    def _ensure_column(
        self, ws, headers: dict[str, int], column_name: str
    ) -> None:
        """Добавляет новую колонку, если её нет."""
        if column_name not in headers:
            new_col = ws.max_column + 1
            cell = ws.cell(self.config.header_position, new_col)
            cell.value = column_name
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center',
                                       wrap_text=True,)
            cell.border = self.cell_border
            col_letter = get_column_letter(new_col)
            ws.column_dimensions[col_letter].width = self.config.column_width

            self.logger.debug('Add column: %s', column_name)

    def _normalize_date(self, date_value, row: int, col: int) -> date | None:
        if (isinstance(date_value, date)
                and not isinstance(date_value, datetime)):
            return date_value

        if isinstance(date_value, datetime):
            return date_value.date()

        message = (f'Неподдерживаемый формат даты в строке {row}, '
                   f'колонке {col}: {date_value!r}')

        if isinstance(date_value, str):
            try:
                return datetime.strptime(date_value, '%d.%m.%Y').date()
            except ValueError:
                self.logger.error(message)
                return None

        # self.logger.error(message)
        return None

    def _subtract_working_days(self, start_date: date, days: int) -> date:
        """При расчете новой даты учитывает рабочие и праздничные дни."""
        current = start_date
        remaining = days
        while remaining > 0:
            current -= timedelta(days=1)
            if self.calendar.is_working_day(current):
                remaining -= 1
        return current
