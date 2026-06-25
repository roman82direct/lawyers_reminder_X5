import logging

from openpyxl import load_workbook

from app.config import ExcelConfig


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
        self.logger.debug(f'HEADERS: {headers}')
        

    def _get_headers(self, ws) -> dict[str, int]:
        """Извлекает названия колонок."""
        headers = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(1, col).value
            if value:
                headers[str(value).strip()] = col
        return headers
