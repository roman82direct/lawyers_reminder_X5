# Работает с Excel:
# открывает файл;
# ищет/создает колонку;
# заполняет пустые значения;
# сохраняет файл;
# возвращает список строк, которые нужно уведомить.

from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from app.services.dates import (
    add_calendar_days, normalize_date, subtract_calendar_days
)


class ExcelService:
    def __init__(self, excel_path: Path, sheet_name: str | None = None):
        self.excel_path = excel_path
        self.sheet_name = sheet_name

    def process_file(self, license_end_column: str, submission_date_column: str,
                     days_before_submission: int, notify_before_days: int):
        wb = load_workbook(self.excel_path)
        ws = wb[self.sheet_name] if self.sheet_name else wb.active

        header_row = 1
        headers = {ws.cell(header_row, col).value: col for col in range(1, ws.max_column + 1)}

        if submission_date_column not in headers:
            new_col = ws.max_column + 1
            ws.cell(header_row, new_col).value = submission_date_column
            headers[submission_date_column] = new_col

        if license_end_column not in headers:
            raise ValueError(f"Missing required column: {license_end_column}")

        name_col = headers.get("Название объекта")
        address_col = headers.get("Адрес объекта")
        end_col = headers[license_end_column]
        sub_col = headers[submission_date_column]

        today = date.today()
        notify_date = today + __import__("datetime").timedelta(
            days=notify_before_days
        )

        reminders = []

        for row in range(2, ws.max_row + 1):
            end_value = ws.cell(row, end_col).value
            if not end_value:
                continue

            end_date = normalize_date(end_value)

            submission_cell = ws.cell(row, sub_col)
            if not submission_cell.value:
                submission_date = subtract_calendar_days(
                    end_date, days_before_submission
                )
                submission_cell.value = submission_date
            else:
                submission_date = normalize_date(submission_cell.value)

            if submission_date == notify_date:
                reminders.append({
                    'row': row,
                    'name': ws.cell(row, name_col).value if name_col else '',
                    'address': ws.cell(row, address_col).value if address_col else '',
                    'license_end_date': end_date,
                    'submission_date': submission_date,
                })

        wb.save(self.excel_path)
        return reminders
