import asyncio

from app.config import ExcelConfig, setup_logging
from app.services.excel import ExcelService
from app.services.notification import NotificationService
from app.services.reminder import ReminderService


async def main():
    setup_logging()
    xls_config = ExcelConfig()
    reminder = ReminderService(ExcelService(xls_config),
                               NotificationService())
    await reminder.run()


if __name__ == '__main__':
    asyncio.run(main())
