from __future__ import annotations

from email import policy
from email.parser import BytesParser
from pathlib import Path


def read_eml(path: str | Path) -> dict[str, str]:
    """Đọc .eml và trả về subject/body ưu tiên HTML, fallback plain text."""
    with Path(path).open("rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    html = ""
    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/html" and not html:
                html = part.get_content()
            elif ctype == "text/plain" and not text:
                text = part.get_content()
    else:
        if msg.get_content_type() == "text/html":
            html = msg.get_content()
        else:
            text = msg.get_content()
    return {"subject": str(msg.get("subject", "")), "body": html or text, "body_format": "html" if html else "plain"}
