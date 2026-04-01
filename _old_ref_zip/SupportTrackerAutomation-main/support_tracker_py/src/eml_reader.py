from datetime import datetime, timedelta
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
        received_time = _parse_received(msg.get_all("received"))
        # Prefer Received time when Date is missing or significantly off.
        if received_time:
            if sent_time == datetime.min:
                sent_time = received_time
            else:
                try:
                    if abs(received_time - sent_time) > timedelta(hours=6):
                        sent_time = received_time
                except Exception:
                    pass
        body, body_html, body_html_raw = _extract_body_parts(msg)
        body = _select_body(body, body_html)

        emails.append(
            EmailRecord(
                path=str(path),
                subject=subject,
                sender_email=sender_email,
                sender_name=sender_name or sender_email,
                sent_time=sent_time,
                body=body,
                body_html=body_html,
                body_html_raw=body_html_raw,
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


def _parse_received(values):
    if not values:
        return None
    for raw in values:
        if not raw:
            continue
        candidate = ""
        if ";" in raw:
            candidate = raw.split(";")[-1].strip()
        else:
            candidate = raw.strip()
        try:
            dt = parsedate_to_datetime(candidate)
        except Exception:
            dt = None
        if isinstance(dt, datetime):
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=IST)
            return dt
    return None


def _extract_body_parts(msg):
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
                html_raw = "\n".join(h for h in html_parts if isinstance(h, str)).strip()
                html_text = "\n".join(_html_to_text(h) for h in html_parts).strip()
                return plain_text, html_text, html_raw
        else:
            content = msg.get_content()
            if isinstance(content, str):
                return content.strip(), "", ""
            return "", "", ""
    except Exception:
        return "", "", ""
    return "", "", ""


def _select_body(plain_text: str, html_text: str) -> str:
    plain_has_headers = _has_header_block(plain_text)
    html_has_headers = _has_header_block(html_text)
    # Prefer the body that contains quoted header blocks.
    if plain_has_headers and html_has_headers:
        # Prefer plain when both have headers to avoid duplication noise.
        return plain_text
    if plain_has_headers:
        return plain_text
    if html_has_headers:
        # Preserve any plain top-level content but include HTML headers.
        if plain_text:
            return f"{plain_text}\n{html_text}"
        return html_text
    if plain_text:
        return plain_text
    if html_text:
        return html_text
    return ""


def _html_to_text(value: str) -> str:
    if not value:
        return ""
    text = _html.unescape(value)
    # Preserve structure for quoted headers by inserting line breaks.
    text = _re.sub(r"(?i)<br\\s*/?>", "\n", text)
    text = _re.sub(r"(?i)</p>", "\n", text)
    text = _re.sub(r"(?i)</\\s*(div|tr|td|th|li|h[1-6])\\s*>", "\n", text)
    text = _re.sub(r"<[^>]+>", "", text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


def _has_header_block(text: str) -> bool:
    if not text:
        return False
    lower = text.lower()
    has_from = ("from:" in lower) or ("de:" in lower)
    has_sent = ("sent:" in lower) or ("envoy" in lower)
    has_subject = ("subject:" in lower) or ("objet" in lower)
    return has_from and has_sent and has_subject
