from __future__ import annotations

from pathlib import Path
from typing import Any

from eml_reader import read_eml
from logger import setup_logger
from state import record_error, record_success

logger = setup_logger()


def _list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [x.strip() for x in value.replace(";", ",").split(",") if x.strip()]
    return [str(x).strip() for x in value if str(x).strip()]


def _extract_email(value: str) -> str:
    text = value.strip()
    if "<" in text and ">" in text:
        return text.split("<", 1)[1].split(">", 1)[0].strip()
    return text


def _account_smtp(account: Any) -> str:
    smtp = str(getattr(account, "SmtpAddress", "") or "").strip()
    if smtp:
        return smtp
    try:
        user = account.CurrentUser
        entry = user.AddressEntry
        if str(getattr(entry, "Type", "")).upper() == "EX":
            exchange_user = entry.GetExchangeUser()
            smtp = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
            if smtp:
                return smtp
        return str(getattr(entry, "Address", "") or "").strip()
    except Exception:
        return ""


def send_mail(mail_config: dict[str, Any]) -> dict[str, Any]:
    """Gửi email qua Microsoft Outlook desktop bằng COM (Windows + pywin32)."""
    try:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError("Cần cài pywin32 và chạy trên Windows có Microsoft Outlook.") from exc

        cfg = dict(mail_config)
        template = cfg.get("template")
        if template and "body" not in cfg:
            data = read_eml(template)
            cfg["subject"] = cfg.get("subject") or data["subject"]
            cfg["body"] = data["body"]
            cfg["body_format"] = data["body_format"]

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        if cfg.get("account"):
            wanted_account = _extract_email(str(cfg["account"])).lower()
            for account in outlook.Session.Accounts:
                smtp = _account_smtp(account).lower()
                display = str(getattr(account, "DisplayName", "") or "").lower()
                if wanted_account in {smtp, display} or wanted_account in f"{display} <{smtp}>":
                    mail.SendUsingAccount = account
                    break
        mail.To = ";".join(_list(cfg.get("to")))
        mail.CC = ";".join(_list(cfg.get("cc")))
        mail.BCC = ";".join(_list(cfg.get("bcc")))
        mail.Subject = cfg.get("subject", "")
        body = cfg.get("body", "")
        if cfg.get("body_format", "html").lower() == "html":
            font = cfg.get("font_family", "Calibri")
            size = cfg.get("font_size", 11)
            mail.HTMLBody = f'<div style="font-family:{font};font-size:{size}pt">{body}</div>'
        else:
            mail.Body = body
        if cfg.get("importance") == "high":
            mail.Importance = 2
        elif cfg.get("importance") == "low":
            mail.Importance = 0
        for attachment in _list(cfg.get("attachments")):
            path = Path(attachment)
            if path.exists():
                mail.Attachments.Add(str(path.resolve()))
        mail.Send()
        summary = f"Đã gửi mail tới {mail.To} - {mail.Subject}"
        logger.info(summary)
        record_success(summary)
        return {"ok": True, "message": summary}
    except Exception as exc:  # noqa: BLE001 - lưu lỗi gửi mail để UI/API hiển thị
        logger.exception("Gửi mail thất bại")
        record_error(str(exc))
        return {"ok": False, "message": str(exc)}
