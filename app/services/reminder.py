# Содержит бизнес-логику:
# определяет объекты, у которых эта дата наступает через 3 дня;
# формирует текст для уведомления.

from app.services.excel import ExcelService


class ReminderService:
    """Класс формирования уведомлений."""

    def __init__(self, excel_service: ExcelService):
        self.excel_service = excel_service

    def run(self):
        return self.excel_service.process_file()
