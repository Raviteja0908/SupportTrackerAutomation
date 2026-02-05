from datetime import datetime
from zoneinfo import ZoneInfo
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
import html as _html
import re as _re

from .models import EmailRecord


def read_eml_directory(root: Path, logger):
    root = Path(root)
    emails = []
    if not root.exists():
        logger.log(f"[WARNING] EML dir missing: {root}")
        return emails

    for path in root.rglob("*.eml"):
        try:
            with path.open("rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
        except Exception:
            continue

        subject = (msg.get("subject") or "").strip()
        from_header = msg.get("from") or ""
        sender_name, sender_email = parseaddr(from_header)
        sender_email = (sender_email or "").strip().lower()

        sent_time = _parse_date(msg.get("date"))
        body = _extract_body(msg)

        emails.append(
            EmailRecord(
                subject=subject,
                sender_email=sender_email,
                sender_name=sender_name or sender_email,
                sent_time=sent_time,
                body=body,
            )
        )

    logger.log(f"[INFO] EML loaded: {len(emails)} from {root}")
    return emails


IST = ZoneInfo("Asia/Kolkata")


def _parse_date(value):
    if not value:
        return datetime.min
    try:
        dt = parsedate_to_datetime(value)
        if not isinstance(dt, datetime):
            return datetime.min
        if dt.tzinfo is None:
            return dt.replace(tzinfo=IST)
        return dt
    except Exception:
        return datetime.min


def _extract_body(msg):
    try:
        if msg.is_multipart():
            plain_parts = []
            html_parts = []
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == "text/plain":
                    content = part.get_content()
                    if isinstance(content, str) and content.strip():
                        plain_parts.append(content.strip())
                elif ctype == "text/html":
                    html_parts.append(part.get_content())

            if plain_parts or html_parts:
                plain_text = "\n".join(plain_parts).strip()
                html_text = "\n".join(_html_to_text(h) for h in html_parts).strip()
                if plain_text and _has_header_block(plain_text):
                    return plain_text
                if html_text:
                    return html_text
                return plain_text
        else:
            content = msg.get_content()
            if isinstance(content, str):
                return content.strip()
            return ""
    except Exception:
        return ""
    return ""


def _html_to_text(value: str) -> str:
    if not value:
        return ""
    text = _html.unescape(value)
    text = _re.sub(r"(?i)<br\\s*/?>", "\n", text)
    text = _re.sub(r"(?i)</p>", "\n", text)
    text = _re.sub(r"<[^>]+>", "", text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


def _has_header_block(text: str) -> bool:
    if not text:
        return False
    lower = text.lower()
    return "from:" in lower and "sent:" in lower and "subject:" in lower
