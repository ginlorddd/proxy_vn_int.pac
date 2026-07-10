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

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        logger.info("Scheduler started")

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=2)
        logger.info("Scheduler stopped")

    def _run(self) -> None:
        while not self._stop.is_set():
            config = self.get_config()
            if should_send_now(config.get("schedule", {}), self._last_key):
                self._last_key = datetime.now().strftime("%Y%m%d%H%M")
                send_mail(config.get("mail", {}))
            self._stop.wait(self.interval_seconds)


def should_send_now(schedule: dict[str, Any], last_key: str) -> bool:
    if not schedule.get("enabled"):
        return False
    now = datetime.now()
    key = now.strftime("%Y%m%d%H%M")
    if key == last_key:
        return False
    typ = schedule.get("type", "daily")
    if typ == "interval":
        minutes = int(schedule.get("interval_minutes", 60))
        anchor = datetime(now.year, now.month, now.day)
        return (now - anchor) >= timedelta() and int((now - anchor).total_seconds() // 60) % minutes == 0
    if typ == "weekly" and now.weekday() not in schedule.get("weekdays", [0,1,2,3,4]):
        return False
    return now.strftime("%H:%M") == schedule.get("time", "08:00")
