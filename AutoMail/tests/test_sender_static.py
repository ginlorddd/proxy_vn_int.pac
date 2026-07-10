from __future__ import annotations

from pathlib import Path

SENDER_PATH = Path(__file__).resolve().parents[1] / "sender.py"
SCHEDULER_PATH = Path(__file__).resolve().parents[1] / "scheduler.py"


def test_sender_body_takes_precedence_over_template() -> None:
    source = SENDER_PATH.read_text(encoding="utf-8")
    assert 'if template and "body" not in cfg:' in source
    assert 'cfg["body"] = data["body"]' in source


def test_date_schedule_template_explicitly_replaces_body() -> None:
    source = SCHEDULER_PATH.read_text(encoding="utf-8")
    assert 'mail.pop("body", None)' in source
    assert 'mail["template"] = date_item["template"]' in source


def test_scheduler_date_item_can_override_recipients_and_repeat() -> None:
    source = SCHEDULER_PATH.read_text(encoding="utf-8")
    assert 'for field in ("to", "cc", "bcc")' in source
    assert 'mail[field] = date_item[field]' in source
    assert 'def _schedule_item_matches' in source
    assert 'repeat == "daily"' in source
    assert 'repeat == "weekly"' in source
    assert 'repeat == "monthly"' in source
