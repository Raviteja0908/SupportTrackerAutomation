import argparse
import csv
import json
import re
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path

from src.rules.subject_normalizer import extract_subject_from_description, normalize_subject
from src.rules.time_resolver import (
    _extract_canonical_message_lines,
    _is_ess_sender,
    _match_requester,
    _to_ist,
)


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


def _parse_eml(path: Path) -> EmailRec | None:
    try:
        with path.open("rb") as handle:
            msg = BytesParser(policy=policy.default).parse(handle)
        plain = ""
        html_body = ""
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
            body=plain or "",
            body_html=html_body or "",
            path=path,
        )
    except Exception:
        return None


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


def _clean_lines(email_obj: EmailRec) -> list[str]:
    return _extract_canonical_message_lines(email_obj)


def _parse_quoted_sent_time(line: str) -> datetime | None:
    line = re.sub(r"(?i)^sent\b\s*:?\s*", "", (line or "")).strip()
    patterns = [
        "%A, %d %B %Y %H:%M",
        "%A, %B %d, %Y %I:%M %p",
        "%d %B %Y %H:%M",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y %H:%M",
        "%m/%d/%Y %I:%M %p",
    ]
    for fmt in patterns:
        try:
            return datetime.strptime(line, fmt)
        except Exception:
            continue
    return None


def _extract_quoted_blocks(email_obj: EmailRec):
    lines = _clean_lines(email_obj)
    out = []

    def _label(line: str):
        m = re.match(r"(?i)^(from|sent|to|cc|subject|objet)\b\s*:?\s*(.*)$", line or "")
        return m.group(1).lower() if m else None

    def _value(line: str, label: str) -> str:
        m = re.match(rf"(?i)^{label}\b\s*:?\s*(.*)$", line or "")
        return (m.group(1) if m else "").strip()

    i = 0
    while i < len(lines):
        if not re.search(r"(?i)\bfrom\b\s*:", lines[i] or ""):
            i += 1
            continue
        from_line = ""
        sent_line = ""
        subj_line = ""
        end = i
        for j in range(i, min(i + 16, len(lines))):
            cur = (lines[j] or "").strip()
            if j > i and (re.search(r"(?i)\bfrom\b\s*:", cur or "") or re.match(r"(?i)^[-_]{3,}$", cur)):
                break
            label = _label(cur)
            if label == "from" and not from_line:
                from_line = _value(cur, "from")
            elif label == "sent" and not sent_line:
                sent_line = cur
            elif label in {"subject", "objet"} and not subj_line:
                subj_line = cur
            end = j
        if sent_line:
            sent_dt = _parse_quoted_sent_time(sent_line)
            subj = re.sub(r"(?i)^(subject|objet)\b\s*:?\s*", "", subj_line).strip() if subj_line else ""
            out.append((from_line, sent_dt, subj))
        i = max(i + 1, end + 1)
    return out


def _family_subject(row: dict) -> str:
    return normalize_subject(extract_subject_from_description(_get(row, "Description")))


def _same_subject(email_subject: str, family_subject: str) -> bool:
    e = normalize_subject(email_subject or "")
    f = normalize_subject(family_subject or "")
    return bool(e and f and (e == f or e in f or f in e))


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
    config_path = Path("/app/config/ess_team.json")
    if not config_path.exists():
        config_path = Path("config/ess_team.json")
    ess_team = json.loads(config_path.read_text(encoding="utf-8"))
    for path in sorted(Path(args.eml_dir).rglob("*.eml")):
        rec = _parse_eml(path)
        if rec:
            emails.append(rec)

    for row in matches:
        family_subject = _family_subject(row)
        requester = _get(row, "Requested By", "Requester")
        created = _get(row, "Created Date & Time")
        response = _get(row, "Actual Response Date & Time")
        resolved = _get(row, "Actual Resolved Date & Time")
        resolved_dt = _parse_cell_dt(resolved)
        resolved_min = _to_ist(resolved_dt).replace(second=0, microsecond=0) if resolved_dt else None

        print("=" * 100)
        print(f"line={row.get('_line')} desc={_get(row, 'Description')}")
        print(f"family_subject={family_subject}")
        print(f"row_triplet={created} / {response} / {resolved}")

        candidates = [
            e for e in emails
            if _same_subject(e.subject, family_subject)
        ]
        candidates.sort(key=lambda e: _to_ist(e.sent_time) if e.sent_time else datetime.min)

        live_replies = []
        for e in candidates:
            if not e.sent_time:
                continue
            if requester and not _match_requester(e.sender_name, e.sender_email, requester):
                continue
            live_replies.append(e)

        print("-" * 100)
        print("live requester replies:")
        for e in live_replies:
            tag = ""
            e_min = _to_ist(e.sent_time).replace(second=0, microsecond=0)
            if resolved_min and e_min == resolved_min:
                tag = "  <-- output resolved"
            print(f"  {_fmt(e.sent_time)} | from={e.sender_email or e.sender_name} | subj={normalize_subject(e.subject)}{tag}")

        target = None
        for e in live_replies:
            e_min = _to_ist(e.sent_time).replace(second=0, microsecond=0)
            if resolved_min and e_min == resolved_min:
                target = e
                break
        if not target and live_replies:
            target = live_replies[-1]

        print("-" * 100)
        if not target:
            print("No live requester reply found for this row.")
            continue

        print(f"quoted blocks under target reply: {_fmt(target.sent_time)} | {target.sender_email or target.sender_name}")
        blocks = _extract_quoted_blocks(target)
        if not blocks:
            print("  no quoted header blocks found")
            continue

        for idx, (from_line, sent_dt, subj) in enumerate(blocks[:12], start=1):
            sent_ist = _to_ist(sent_dt) if sent_dt else None
            ess = False
            try:
                temp = type("T", (), {"sender_name": from_line, "sender_email": from_line})()
                ess = _is_ess_sender(temp, ess_team)
            except Exception:
                ess = False
            subj_norm = normalize_subject(subj or "")
            subj_match = _same_subject(subj_norm, family_subject)
            print(
                f"  {idx}. {_fmt(sent_ist)} | {'ESS' if ess else 'NON-ESS'} | subj_match={subj_match}"
            )
            print(f"     from={from_line}")
            print(f"     subj={subj_norm or '-'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
