# Отвечает за расписание.
# Если процесс живет постоянно — запускает задачу каждый день в 00:00.
# Проще и надежнее — этот модуль можно почти не использовать,
# а запускать exe через Windows Task Scheduler.

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger


class SchedulerService:
    def __init__(self, job_func, hour: int, minute: int):
        self.job_func = job_func
        self.scheduler = BackgroundScheduler()

        self.scheduler.add_job(
            self.job_func,
            CronTrigger(hour=hour, minute=minute),
            id='daily_reminder_job',
            replace_existing=True,
            max_instances=1,
            coalesce=True,
        )

    def start(self):
        self.scheduler.start()

    def shutdown(self):
        self.scheduler.shutdown(wait=False)
