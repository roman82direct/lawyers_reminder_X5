import logging

from desktop_notifier import DesktopNotifier


class NotificationService:
    """Сервис отправки уведомлений."""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.notifier = DesktopNotifier()

    async def send(self, title: str, message: str) -> None:
        try:
            await self.notifier.send(title=title, message=message)
        except Exception:
            self.logger.exception('desktop-notifier error')
