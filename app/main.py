import logging

from app.config import AppConfig
from app.scheduler import SchedulerService
from app.services.excel import ExcelService
from app.services.notification import NotificationService
from app.services.reminder import ReminderService


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(name)s %(message)s',
        handlers=[
            logging.FileHandler('logs/app.log', encoding='utf-8'),
            logging.StreamHandler(),
        ],
    )


def build_app():
    config = AppConfig()
    excel_service = ExcelService(config.excel_path, config.sheet_name)
    notification_service = NotificationService()
    reminder_service = ReminderService(excel_service, notification_service)

    def job():
        reminder_service.run(config)

    scheduler = SchedulerService(job, config.run_hour, config.run_minute)
    return scheduler


def main():
    setup_logging()
    scheduler = build_app()
    scheduler.start()

    try:
        import time
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        scheduler.shutdown()


if __name__ == '__main__':
    main()
