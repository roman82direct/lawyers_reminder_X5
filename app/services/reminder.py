from app.config import ExcelConfig
from app.services.excel import ExcelService, ProcessResult
from app.services.notification import NotificationService

config = ExcelConfig()


def create_notification_text(result: ProcessResult) -> str:
    """Собирает текст уведомления с данными полученными при обработке xlsx."""
    parts: list[str] = []

    if result.objects_to_notify:
        parts.append(
            f'Осталось {config.notify_before_days} рабочих дня '
            f'до подачи заявления на переоформление:\n' + '\n'.join(
                f'- {obj}' for obj in result.objects_to_notify
            )
        )
    else:
        parts.append('Срочных переоформлений сегодня нет.')

    if result.errors:
        parts.append(
            '\nОшибки:\n' + '\n'.join(f'- {err}' for err in result.errors)
        )

    return '\n'.join(parts)


class ReminderService:
    """
    Класс бизнес-логики приложения.

    excel_service - обрабатывает xlsx-файл.
    notification_service - отправляет уведомление.
    """

    def __init__(
        self,
        excel_service: ExcelService,
        notification_service: NotificationService,
    ):
        self.excel_service = excel_service
        self.notification_service = notification_service

    async def run(self):
        await self.notification_service.send(
            'Проверка лицензий',
            create_notification_text(self.excel_service.process_file())
        )
