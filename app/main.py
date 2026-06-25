from app.config import ExcelConfig, setup_logging
from app.services.excel import ExcelService
from app.services.reminder import ReminderService


def main():
    setup_logging()
    xls_config = ExcelConfig()
    reminder = ReminderService(
        ExcelService(xls_config)
    )
    reminder.run()


if __name__ == '__main__':
    main()
