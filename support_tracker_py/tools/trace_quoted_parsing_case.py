import argparse
import csv
import html
import re
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path

from src.rules.subject_normalizer import extract_subject_from_description, normalize_subject
from src.rules.time_resolver import _match_requester, _to_ist

try:
    from bs4 import BeautifulSoup  # type: ignore
except Exception:  # pragma: no cover
    BeautifulSoup = None


@dataclass
class EmailRec:
    subject: str
    sender_name: str
    sender_email: str
    sent_time: datetime | None
    body: str
    body_html: str
    path: Path


def _load_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = []
        for idx, row in enumerate(csv.DictReader(handle), start=2):
            row["_line"] = idx
            rows.append(row)
        return rows


def _get(row: dict, *names: str) -> str:
    lowered = {str(k).strip().lower(): v for k, v in row.items()}
    for name in names:
        if name in row and row[name] not in (None, ""):
            return str(row[name]).strip()
        value = lowered.get(name.lower())
        if value not in (None, ""):
            return str(value).strip()
    return ""


def _fmt(dt: datetime | None) -> str:
    if not dt:
        return "-"
    try:
        dt = _to_ist(dt)
    except Exception:
        pass
    return dt.strftime("%d-%m-%Y %H:%M")


def _parse_cell_dt(value: str) -> datetime | None:
    if not value:
        return None
    for fmt in ("%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(value.strip(), fmt)
        except Exception:
            continue
    return None


def _extract_body(msg) -> tuple[str, str]:
    plain = ""
    html_body = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                ctype = (part.get_content_type() or "").lower()
                try:
                    content = part.get_content()
                except Exception:
                    content = None
                if ctype == "text/plain" and isinstance(content, str) and not plain:
                    plain = content
                elif ctype == "text/html" and isinstance(content, str) and not html_body:
                    html_body = content
        else:
            try:
                content = msg.get_content()
            except Exception:
                content = ""
            if isinstance(content, str):
                plain = content
    except Exception:
        pass
    return plain or "", html_body or ""


def _parse_eml(path: Path) -> EmailRec | None:
    try:
        with path.open("rb") as handle:
            msg = BytesParser(policy=policy.default).parse(handle)
        plain, html_body = _extract_body(msg)
        sender_name = ""
        sender_email = ""
        try:
            addrs = getaddresses([msg.get("from", "")])
            if addrs:
                sender_name, sender_email = addrs[0]
        except Exception:
            pass
        sent_time = None
        try:
            sent_time = parsedate_to_datetime(msg.get("date", ""))
        except Exception:
            sent_time = None
        return EmailRec(
            subject=str(msg.get("subject", "") or ""),
            sender_name=sender_name or "",
            sender_email=(sender_email or "").lower(),
            sent_time=sent_time,
            body=plain,
            body_html=html_body,
            path=path,
        )
    except Exception:
        return None


def _family_subject(row: dict) -> str:
    return normalize_subject(extract_subject_from_description(_get(row, "Description")))


def _same_subject(a: str, b: str) -> bool:
    a_norm = normalize_subject(a or "")
    b_norm = normalize_subject(b or "")
    return bool(a_norm and b_norm and (a_norm == b_norm or a_norm in b_norm or b_norm in a_norm))


def _current_clean_lines(email_obj: EmailRec) -> list[str]:
    raw = f"{email_obj.body}\n{email_obj.body_html}"
    if not raw:
        return []
    txt = re.sub(r"(?is)<style.*?>.*?</style>", " ", raw)
    txt = re.sub(r"(?is)<script.*?>.*?</script>", " ", txt)
    txt = re.sub(r"(?i)<\s*br\s*/?>", "\n", txt)
    txt = re.sub(r"(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>", "\n", txt)
    txt = re.sub(r"(?is)<[^>]+>", " ", txt)
    txt = html.unescape(txt)
    return [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]


def _bs4_clean_lines(email_obj: EmailRec) -> list[str]:
    if BeautifulSoup is None:
        return []
    html_source = email_obj.body_html or ""
    if html_source:
        soup = BeautifulSoup(html_source, "html.parser")
        text = soup.get_text("\n")
    else:
        text = email_obj.body
    text = html.unescape(text or "")
    return [ln.strip() for ln in text.splitlines() if ln and ln.strip()]


def _parse_quoted_sent_time(line: str) -> datetime | None:
    line = re.sub(r"(?i)^sent\b\s*:?\s*", "", (line or "")).strip()
    patterns = [
        "%A, %d %B %Y %I:%M %p",
        "%A, %d %B %Y %H:%M",
        "%A, %B %d, %Y %I:%M %p",
        "%d %B %Y %I:%M %p",
        "%d %B %Y %H:%M",
        "%d-%m-%Y %I:%M %p",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y %I:%M %p",
        "%d/%m/%Y %H:%M",
        "%m/%d/%Y %I:%M %p",
    ]
    for fmt in patterns:
        try:
            return datetime.strptime(line, fmt)
        except Exception:
            continue
    # Primary rule: use the first real date-like fragment from the Sent line and
    # preserve AM/PM exactly when present.
    fragments = []
    regexes = [
        r"\b[A-Za-z]+,\s+\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}\s*[APMapm]{2}\b",
        r"\b[A-Za-z]+,\s+\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}\b",
        r"\b[A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s*[APMapm]{2}\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}\s*[APMapm]{2}\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\s+\d{1,2}:\d{2}\s*[APMapm]{2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\s+\d{1,2}:\d{2}\b",
    ]
    for rx in regexes:
        m = re.search(rx, line)
        if m:
            fragments.append(m.group(0).strip())
    for frag in fragments:
        for fmt in patterns:
            try:
                return datetime.strptime(frag, fmt)
            except Exception:
                continue
    return None


def _extract_blocks_from_lines(lines: list[str]) -> list[tuple[int, str, str, datetime | None, str, list[str]]]:
    def _header_label(line: str):
        m = re.match(r"(?i)^(from|sent|to|cc|subject|objet)\b\s*:?\s*(.*)$", line or "")
        return m.group(1).lower() if m else None

    def _header_value(line: str, label: str) -> str:
        m = re.match(rf"(?i)^{label}\b\s*:?\s*(.*)$", line or "")
        return (m.group(1) if m else "").strip()

    def _from_start(line: str) -> bool:
        return bool(re.search(r"(?i)\bfrom\b\s*:", line or ""))

    def _looks_like_body_line(line: str) -> bool:
        text = (line or "").strip()
        if not text:
            return False
        low = text.lower()
        if _header_label(text):
            return False
        if re.match(r"(?i)^(hi|hello|dear|regards|kind regards|best regards|thanks|thank you|please)\b", text):
            return True
        if low.startswith("@"):
            return True
        if len(text.split()) >= 7 and not re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, flags=re.I):
            return True
        return False

    out = []
    i = 0
    while i < len(lines):
        if not _from_start(lines[i]):
            i += 1
            continue
        from_line = ""
        sent_line = ""
        subj_line = ""
        block_lines = []
        end = i
        for j in range(i, min(i + 16, len(lines))):
            cur = (lines[j] or "").strip()
            block_lines.append(cur)
            if j > i and (_from_start(cur) or re.match(r"(?i)^[-_]{3,}$", cur)):
                break
            label = _header_label(cur)
            if label == "from" and not from_line:
                from_line = _header_value(cur, "from")
            elif label == "sent" and not sent_line:
                sent_line = cur
            elif label in {"subject", "objet"} and not subj_line:
                subj_line = cur
            elif j > i and label is None and _looks_like_body_line(cur):
                end = j - 1
                break
            elif j > i and label is None and sent_line and (subj_line or from_line):
                end = j - 1
                break
            end = j
        sent_dt = _parse_quoted_sent_time(sent_line) if sent_line else None
        subj = re.sub(r"(?i)^(subject|objet)\b\s*:?\s*", "", subj_line).strip() if subj_line else ""
        out.append((i, from_line, sent_line, sent_dt, subj, block_lines))
        i = max(i + 1, end + 1)
    return out


def _print_lines(title: str, lines: list[str], limit: int = 80):
    print("-" * 100)
    print(title)
    if not lines:
        print("  <no lines>")
        return
    for idx, line in enumerate(lines[:limit], start=1):
        print(f"  {idx:02d}: {line}")


def _print_blocks(title: str, blocks):
    print("-" * 100)
    print(title)
    if not blocks:
        print("  <no blocks>")
        return
    for idx, (start_idx, from_line, sent_line, sent_dt, subj, raw_lines) in enumerate(blocks, start=1):
        print(f"  block {idx}: start_line={start_idx + 1} sent={_fmt(sent_dt)}")
        print(f"    from={from_line or '-'}")
        print(f"    sent_line={sent_line or '-'}")
        print(f"    subj={normalize_subject(subj) or '-'}")
        for raw in raw_lines[:8]:
            print(f"      {raw}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-csv", required=True)
    parser.add_argument("--eml-dir", required=True)
    parser.add_argument("--subject", required=True)
    args = parser.parse_args()

    rows = _load_csv(Path(args.output_csv))
    matches = [r for r in rows if args.subject.lower() in _get(r, "Description").lower()]
    if not matches:
        print("No matching rows found.")
        return 1

    emails = []
    for path in sorted(Path(args.eml_dir).rglob("*.eml")):
        rec = _parse_eml(path)
        if rec:
            emails.append(rec)

    for row in matches:
        family_subject = _family_subject(row)
        requester = _get(row, "Requested By", "Requester")
        resolved = _get(row, "Actual Resolved Date & Time")
        resolved_dt = _parse_cell_dt(resolved)
        resolved_min = _to_ist(resolved_dt).replace(second=0, microsecond=0) if resolved_dt else None

        print("=" * 100)
        print(f"line={row.get('_line')} desc={_get(row, 'Description')}")
        print(f"family_subject={family_subject}")
        print(f"row_triplet={_get(row, 'Created Date & Time')} / {_get(row, 'Actual Response Date & Time')} / {resolved}")

        live_replies = []
        for e in emails:
            if not e.sent_time:
                continue
            if not _same_subject(e.subject, family_subject):
                continue
            if requester and not _match_requester(e.sender_name, e.sender_email, requester):
                continue
            live_replies.append(e)
        live_replies.sort(key=lambda e: _to_ist(e.sent_time))

        target = None
        for e in live_replies:
            e_min = _to_ist(e.sent_time).replace(second=0, microsecond=0)
            if resolved_min and e_min == resolved_min:
                target = e
                break
        if target is None and live_replies:
            target = live_replies[-1]

        print("-" * 100)
        print("live requester replies:")
        for e in live_replies:
            mark = ""
            e_min = _to_ist(e.sent_time).replace(second=0, microsecond=0)
            if resolved_min and e_min == resolved_min:
                mark = "  <-- output resolved"
            print(f"  {_fmt(e.sent_time)} | from={e.sender_email or e.sender_name} | subj={normalize_subject(e.subject)}{mark}")

        if not target:
            print("No target live reply found.")
            continue

        print("-" * 100)
        print(f"target reply path={target.path}")
        print(f"bs4_available={'yes' if BeautifulSoup is not None else 'no'}")

        current_lines = _current_clean_lines(target)
        current_blocks = _extract_blocks_from_lines(current_lines)
        _print_lines("current cleaned lines", current_lines)
        _print_blocks("current parsed blocks", current_blocks)

        bs4_lines = _bs4_clean_lines(target)
        if bs4_lines:
            bs4_blocks = _extract_blocks_from_lines(bs4_lines)
            _print_lines("BeautifulSoup cleaned lines", bs4_lines)
            _print_blocks("BeautifulSoup parsed blocks", bs4_blocks)
        else:
            print("-" * 100)
            print("BeautifulSoup cleaned lines")
            print("  <bs4 not available>")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
