from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime, getaddresses
from pathlib import Path
import os
import html as _html
import re as _re
from concurrent.futures import ThreadPoolExecutor, as_completed
try:
    from bs4 import BeautifulSoup as _BeautifulSoup
except Exception:
    _BeautifulSoup = None

from .models import EmailRecord


def read_eml_directory(root: Path, logger):
    root = Path(root)
    emails = []
    if not root.exists():
        logger.log(f"[WARNING] EML dir missing: {root}")
        return emails

    paths = list(root.rglob("*.eml"))
    if not paths:
        logger.log(f"[INFO] EML loaded: 0 from {root}")
        return emails

    failed_paths = []

    def _parse_one(path):
        try:
            with path.open("rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
        except Exception as exc:
            logger.log(f"[WARNING] Failed to parse EML file {path}: {exc}")
            return None

        subject = (msg.get("subject") or "").strip()
        from_header = msg.get("from") or ""
        sender_name, sender_email = parseaddr(from_header)
        sender_email = (sender_email or "").strip().lower()
        to_header = msg.get_all("to") or []
        cc_header = msg.get_all("cc") or []
        to_recipients = tuple(
            sorted({addr.strip().lower() for _name, addr in getaddresses(to_header) if addr})
        )
        cc_recipients = tuple(
            sorted({addr.strip().lower() for _name, addr in getaddresses(cc_header) if addr})
        )

        sent_time = _parse_date(msg.get("date"))
        received_time = _parse_received(msg.get_all("received"))
        # Prefer Received time when Date is missing or significantly off.
        if received_time:
            if sent_time is None:
                sent_time = received_time
            elif sent_time is not None:
                try:
                    if abs(received_time - sent_time) > timedelta(hours=6):
                        sent_time = received_time
                except Exception:
                    pass
        # Default to received_time if sent_time is still None, otherwise use sent_time
        if sent_time is None:
            sent_time = received_time
        # If both are None, use current time as fallback
        if sent_time is None:
            sent_time = datetime.now(IST)
        body, body_html, body_html_raw = _extract_body_parts(msg)
        body = _select_body(body, body_html)

        return EmailRecord(
            path=str(path),
            subject=subject,
            sender_email=sender_email,
            sender_name=sender_name or sender_email,
            sent_time=sent_time,
            body=body,
            body_html=body_html,
            body_html_raw=body_html_raw,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
        )

    max_workers = min(8, os.cpu_count() or 4)
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_parse_one, p): p for p in paths}
        for future in as_completed(futures):
            result = future.result()
            path = futures[future]
            if result is not None:
                emails.append(result)
            else:
                failed_paths.append(str(path))

    logger.log(f"[INFO] EML loaded: {len(emails)} from {root}")
    if failed_paths:
        logger.log(f"[WARNING] Failed to parse {len(failed_paths)} EML files (see details above)")

    return emails


IST = ZoneInfo("Asia/Kolkata")


def _parse_date(value):
    if not value:
        return None
    try:
        dt = parsedate_to_datetime(value)
        if not isinstance(dt, datetime):
            return None
        if dt.tzinfo is None:
            # Note: Assuming IST for timezone-naive dates. This may be incorrect for non-India emails.
            return dt.replace(tzinfo=IST)
        return dt
    except Exception:
        return None


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
        return html_text
    if plain_text:
        return plain_text
    if html_text:
        return html_text
    return ""


def _html_to_text(value: str) -> str:
    if not value:
        return ""
    text = ""
    if _BeautifulSoup is not None:
        try:
            soup = _BeautifulSoup(value, "html.parser")
            for tag in soup.find_all(["br", "p", "div", "blockquote", "td", "tr", "li"]):
                tag.insert_after("\n")
            text = _html.unescape(soup.get_text(" "))
        except Exception as exc:
            text = ""
    if not text:
        # Fallback when BS4 parsing is unavailable or errors
        try:
            text = _html.unescape(value)
            text = _re.sub(r"(?i)<br\s*/>", "\n", text)
            text = _re.sub(r"(?i)</p>", "\n", text)
            text = _re.sub(r"(?i)</\s*(div|tr|td|th|li|h[1-6])\s*>", "\n", text)
            text = _re.sub(r"<[^>]+>", "", text)
        except Exception as exc:
            text = ""  # Return empty string if fallback fails
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    result = "\n".join(lines)
    # Validate output is not garbage (at least some printable content)
    if result and len(result) > 10:
        return result
    return ""


def _has_header_block(text: str) -> bool:
    if not text:
        return False
    lower = text.lower()
    has_from = ("from:" in lower) or ("de:" in lower)
    has_sent = ("sent:" in lower) or ("envoy" in lower)
    has_subject = ("subject:" in lower) or ("objet" in lower)
    return has_from and has_sent and has_subject
