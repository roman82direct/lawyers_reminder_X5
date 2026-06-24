# Показывает системное уведомление в Windows.

from dataclasses import dataclass


@dataclass
class NotificationService:
    def notify(self, title: str, message: str) -> None:
        try:
            from plyer import notification
            notification.notify(title=title, message=message, timeout=15)
        except Exception:
            pass
