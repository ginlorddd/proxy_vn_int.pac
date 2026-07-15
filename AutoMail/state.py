from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
STATE_FILE = BASE_DIR / "state.json"
DEFAULT_STATE = {"last_sent_at": None, "sent_count": 0, "last_error": None, "history": []}


def load_state() -> dict[str, Any]:
    if not STATE_FILE.exists():
        save_state(DEFAULT_STATE.copy())
        return DEFAULT_STATE.copy()
    with STATE_FILE.open("r", encoding="utf-8") as f:
        return {**DEFAULT_STATE, **json.load(f)}


def save_state(state: dict[str, Any]) -> None:
    with STATE_FILE.open("w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def record_success(summary: str) -> None:
    state = load_state()
    now = datetime.now().isoformat(timespec="seconds")
    state["last_sent_at"] = now
    state["sent_count"] = int(state.get("sent_count", 0)) + 1
    state["last_error"] = None
    state.setdefault("history", []).append({"time": now, "status": "sent", "summary": summary})
    state["history"] = state["history"][-100:]
    save_state(state)


def record_error(error: str) -> None:
    state = load_state()
    now = datetime.now().isoformat(timespec="seconds")
    state["last_error"] = error
    state.setdefault("history", []).append({"time": now, "status": "error", "summary": error})
    state["history"] = state["history"][-100:]
    save_state(state)
