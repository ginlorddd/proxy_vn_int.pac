from __future__ import annotations

from pathlib import Path

SENDER_PATH = Path(__file__).resolve().parents[1] / "sender.py"
SCHEDULER_PATH = Path(__file__).resolve().parents[1] / "scheduler.py"


def test_sender_body_takes_precedence_over_template() -> None:
    source = SENDER_PATH.read_text(encoding="utf-8")
    assert 'if template and (force_template or "body" not in cfg):' in source
    assert 'cfg["body"] = data["body"]' in source


def test_date_schedule_template_explicitly_replaces_body() -> None:
    source = SCHEDULER_PATH.read_text(encoding="utf-8")
    assert 'mail.pop("body", None)' in source
    assert 'mail["template"] = date_item["template"]' in source


def test_scheduler_date_item_can_override_recipients_and_repeat() -> None:
    source = SCHEDULER_PATH.read_text(encoding="utf-8")
    assert 'for field in ("account", "to", "cc", "bcc", "subject")' in source
    assert 'mail[field] = date_item[field]' in source
    assert 'def _schedule_item_matches' in source
    assert 'repeat == "daily"' in source
    assert 'repeat == "weekly"' in source
    assert 'repeat == "monthly"' in source


def test_sender_requires_selected_account_match() -> None:
    source = SENDER_PATH.read_text(encoding="utf-8")
    assert "def _iter_com_collection" in source
    assert "def _find_outlook_account" in source
    assert "raise RuntimeError" in source
    assert "def _apply_send_account" in source
    assert "mail.SendUsingAccount = account" in source
    assert "SentOnBehalfOfName" in source
    assert "Không set SentOnBehalfOfName" in source


def test_scheduler_forces_template_and_account_for_date_items() -> None:
    source = SCHEDULER_PATH.read_text(encoding="utf-8")
    assert 'for field in ("account", "to", "cc", "bcc", "subject")' in source
    assert 'mail["force_template"] = True' in source
    sender = SENDER_PATH.read_text(encoding="utf-8")
    assert 'force_template = bool(cfg.pop("force_template", False))' in sender
    assert 'if template and (force_template or "body" not in cfg):' in sender


def test_scheduler_exposes_status() -> None:
    source = SCHEDULER_PATH.read_text(encoding="utf-8")
    assert "self._stopping = False" in source
    assert "def status" in source
    assert 'return "running"' in source
    assert 'return "stopped"' in source
