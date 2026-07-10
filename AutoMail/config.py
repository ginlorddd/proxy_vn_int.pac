from __future__ import annotations

import json
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
TEMPLATE_DIR = BASE_DIR / "templates"

DEFAULT_CONFIG: dict[str, Any] = {
    "api": {"host": "127.0.0.1", "port": 8765, "token": "change-me"},
    "mail": {
        "account": "",
        "to": [],
        "cc": [],
        "bcc": [],
        "subject": "AutoMail",
        "body": "<p>Xin chào,</p><p>Đây là email tự động.</p>",
        "body_format": "html",
        "font_family": "Calibri",
        "font_size": 11,
        "importance": "normal",
        "attachments": [],
        "template": ""
    },
    "templates": [],
    "schedule": {"enabled": False, "type": "daily", "time": "08:00", "weekdays": [0,1,2,3,4], "interval_minutes": 60, "date_schedules": []},
    "send_on_startup": False
}


def load_config() -> dict[str, Any]:
    if not CONFIG_FILE.exists():
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    with CONFIG_FILE.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return _merge(DEFAULT_CONFIG.copy(), data)


def save_config(config: dict[str, Any]) -> None:
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    with CONFIG_FILE.open("w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def _merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            base[k] = _merge(base[k], v)
        else:
            base[k] = v
    return base
