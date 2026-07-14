from __future__ import annotations

import threading
from datetime import datetime, timedelta
from typing import Any, Callable

from logger import setup_logger
from sender import send_mail

logger = setup_logger()


class MailScheduler:
    def __init__(self, get_config: Callable[[], dict[str, Any]], interval_seconds: int = 30) -> None:
        self.get_config = get_config
        self.interval_seconds = interval_seconds
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._last_key = ""
        self._stopping = False

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stopping = False
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        logger.info("Scheduler started")

    def stop(self) -> None:
        self._stopping = True
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=2)
        self._stopping = False
        logger.info("Scheduler stopped")

    def status(self) -> str:
        if self._stopping:
            return "stopping"
        if self._thread and self._thread.is_alive():
            return "running"
        return "stopped"

    def _run(self) -> None:
        while not self._stop.is_set():
            config = self.get_config()
            schedule = config.get("schedule", {})
            date_item = matching_date_schedule(schedule, self._last_key)
            if date_item or should_send_now(schedule, self._last_key):
                self._last_key = datetime.now().strftime("%Y%m%d%H%M")
                mail = dict(config.get("mail", {}))
                if date_item:
                    for field in ("account", "to", "cc", "bcc", "subject"):
                        if date_item.get(field):
                            mail[field] = date_item[field]
                    if date_item.get("template"):
                        mail.pop("body", None)
                        mail["template"] = date_item["template"]
                        mail["force_template"] = True
                send_mail(mail)
            self._stop.wait(self.interval_seconds)


def matching_date_schedule(schedule: dict[str, Any], last_key: str) -> dict[str, Any] | None:
    if not schedule.get("enabled"):
        return None
    now = datetime.now()
    if now.strftime("%Y%m%d%H%M") == last_key:
        return None
    for item in schedule.get("date_schedules", []):
        if _schedule_item_matches(item, now):
            return item
    return None


def _schedule_item_matches(item: dict[str, Any], now: datetime) -> bool:
    if item.get("time", "08:00") != now.strftime("%H:%M"):
        return False
    repeat = str(item.get("repeat", "once")).lower()
    item_date = str(item.get("date", ""))
    if repeat == "daily":
        return True
    if repeat == "weekly":
        try:
            return datetime.fromisoformat(item_date).weekday() == now.weekday()
        except ValueError:
            return False
    if repeat == "monthly":
        try:
            return datetime.fromisoformat(item_date).day == now.day
        except ValueError:
            return False
    return item_date == now.strftime("%Y-%m-%d")


def should_send_now(schedule: dict[str, Any], last_key: str) -> bool:
    if not schedule.get("enabled"):
        return False
    now = datetime.now()
    key = now.strftime("%Y%m%d%H%M")
    if key == last_key:
        return False
    if matching_date_schedule(schedule, last_key):
        return True
    typ = schedule.get("type", "daily")
    if typ == "interval":
        minutes = int(schedule.get("interval_minutes", 60))
        anchor = datetime(now.year, now.month, now.day)
        return (now - anchor) >= timedelta() and int((now - anchor).total_seconds() // 60) % minutes == 0
    if typ == "weekly" and now.weekday() not in schedule.get("weekdays", [0,1,2,3,4]):
        return False
    return now.strftime("%H:%M") == schedule.get("time", "08:00")
