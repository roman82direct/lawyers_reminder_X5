# Содержит бизнес-логику:
# рассчитав дату подачи как license_end - 15 дней;
# определяет объекты, у которых эта дата наступает через 10 дней;
# формирует текст для уведомления.

from app.services.excel import ExcelService
from app.services.notification import NotificationService


class ReminderService:
    def __init__(
            self, excel_service: ExcelService,
            notification_service: NotificationService
    ):
        self.excel_service = excel_service
        self.notification_service = notification_service

    def run(self, config):
        reminders = self.excel_service.process_file(
            license_end_column=config.license_end_column,
            submission_date_column=config.submission_date_column,
            days_before_submission=config.days_before_submission,
            notify_before_days=config.notify_before_days,
        )

        if reminders:
            lines = [
                f'{item['name']} | {item['address']} | '
                f'срок подачи: {item['submission_date']}'
                for item in reminders
            ]
            self.notification_service.notify(
                title='Срок продления лицензии через 10 дней',
                message='\n'.join(lines),
            )
        return reminders
