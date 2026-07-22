from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from eml_reader import read_eml
from logger import setup_logger
from state import record_error, record_success

logger = setup_logger()

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


def _list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [x.strip() for x in value.replace(";", ",").split(",") if x.strip()]
    return [str(x).strip() for x in value if str(x).strip()]


def _body_text(value: Any) -> str:
    text = re.sub(r"<[^>]+>", " ", str(value or ""))
    text = text.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">")
    return " ".join(text.split())


def _validate_mail_config(cfg: dict[str, Any]) -> None:
    recipients = [*_list(cfg.get("to")), *_list(cfg.get("cc")), *_list(cfg.get("bcc"))]
    if not recipients:
        raise RuntimeError("Không gửi mail vì chưa có người nhận To/Cc/Bcc.")
    if not cfg.get("template") and not str(cfg.get("subject", "")).strip() and not _body_text(cfg.get("body", "")):
        raise RuntimeError("Không gửi mail rỗng: cần có subject, body hoặc template.")


def _extract_email(value: str) -> str:
    text = value.strip()
    if "<" in text and ">" in text:
        return text.split("<", 1)[1].split(">", 1)[0].strip()
    return text


def _iter_com_collection(collection: Any) -> list[Any]:
    try:
        return [collection.Item(index) for index in range(1, int(collection.Count) + 1)]
    except Exception:
        try:
            return list(collection)
        except Exception:
            return []


def _property_accessor_smtp(obj: Any) -> str:
    try:
        return str(obj.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) or "").strip()
    except Exception:
        return ""


def _address_entry_smtp(entry: Any) -> str:
    try:
        if str(getattr(entry, "Type", "")).upper() == "EX":
            exchange_user = entry.GetExchangeUser()
            smtp = str(getattr(exchange_user, "PrimarySmtpAddress", "") or "").strip()
            if smtp:
                return smtp
        return str(getattr(entry, "Address", "") or "").strip()
    except Exception:
        return ""


def _account_smtp(account: Any) -> str:
    for candidate in (
        str(getattr(account, "SmtpAddress", "") or "").strip(),
        str(getattr(account, "UserName", "") or "").strip(),
    ):
        if "@" in candidate:
            return candidate
    try:
        smtp = _address_entry_smtp(account.CurrentUser.AddressEntry)
        if smtp:
            return smtp
    except Exception:
        pass
    try:
        smtp = _property_accessor_smtp(account.DeliveryStore)
        if smtp:
            return smtp
    except Exception:
        pass
    return ""


def _find_outlook_account(session: Any, wanted: str) -> Any:
    wanted_account = _extract_email(wanted).lower()
    for account in _iter_com_collection(session.Accounts):
        smtp = _account_smtp(account).lower()
        display = str(getattr(account, "DisplayName", "") or "").lower()
        user_name = str(getattr(account, "UserName", "") or "").lower()
        candidates = {smtp, display, user_name, f"{display} <{smtp}>"}
        if wanted_account in candidates or wanted_account == smtp:
            return account
    raise RuntimeError(f"Không tìm thấy Outlook account đã chọn: {wanted}")


def _apply_send_account(mail: Any, account: Any) -> None:
    # Chỉ dùng SendUsingAccount cho account Outlook đã đăng nhập.
    # Không set SentOnBehalfOfName vì sẽ biến mail thành dạng "gửi thay mặt" và dễ hiển thị sai From.
    mail.SendUsingAccount = account


def send_mail(mail_config: dict[str, Any]) -> dict[str, Any]:
    """Gửi email qua Microsoft Outlook desktop bằng COM (Windows + pywin32)."""
    try:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError("Cần cài pywin32 và chạy trên Windows có Microsoft Outlook.") from exc

        cfg = dict(mail_config)
        template = cfg.get("template")
        force_template = bool(cfg.pop("force_template", False))
        if template and (force_template or "body" not in cfg):
            data = read_eml(template)
            cfg["subject"] = cfg.get("subject") or data["subject"]
            cfg["body"] = data["body"]
            cfg["body_format"] = data["body_format"]

        _validate_mail_config(cfg)

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        selected_account = _find_outlook_account(outlook.Session, str(cfg["account"])) if cfg.get("account") else None
        if selected_account is not None:
            _apply_send_account(mail, selected_account)
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
        if selected_account is not None:
            _apply_send_account(mail, selected_account)
        mail.Send()
        summary = f"Đã gửi mail tới {mail.To} - {mail.Subject}"
        logger.info(summary)
        record_success(summary)
        return {"ok": True, "message": summary}
    except Exception as exc:  # noqa: BLE001 - lưu lỗi gửi mail để UI/API hiển thị
        logger.exception("Gửi mail thất bại")
        record_error(str(exc))
        return {"ok": False, "message": str(exc)}
