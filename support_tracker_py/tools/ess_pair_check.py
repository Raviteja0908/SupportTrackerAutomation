# Tool: ESS-only quoted pair checker
# Usage:
#   python support_tracker_py/tools/ess_pair_check.py --subject "PE-15800052-S8 claim skipped to process in ExpenSys" --eml-root "D:\Support_Tracker\DockerOutput\eml\My Outlook Data File(1)\My Outlook Data File(1)\raviteja.dwarampudi@invenio-solutions.com\My Team"

import argparse
import email
import html as _html
import json
import re
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from pathlib import Path
import sys

THIS_DIR = Path(__file__).resolve().parent
ROOT = THIS_DIR.parent
sys.path.append(str(ROOT))

try:
    from src.rules.subject_normalizer import normalize_subject
except Exception:
    def normalize_subject(subject: str) -> str:
        return (subject or '').strip()


def _match_tokens(text: str):
    if not text:
        return set()
    t = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return {p for p in t.split() if p}


def _token_overlap_score(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    return (2 * inter) / (len(a) + len(b))


def _parse_sent_line(sent_line: str):
    if not sent_line:
        return None
    s = re.sub(r"(?i)^sent\s*:\s*", "", sent_line).strip()
    s = re.sub(r"(?i)^(mon|tue|wed|thu|fri|sat|sun)\w*,?\s*", "", s).strip()
    s = re.sub(r"\(.*?\)", "", s).strip()
    s = " ".join(s.replace(",", " ").split())
    fmts = [
        "%d %B %Y %H:%M:%S",
        "%d %B %Y %H:%M",
        "%d %b %Y %H:%M:%S",
        "%d %b %Y %H:%M",
        "%d %B %Y %I:%M %p",
        "%d %b %Y %I:%M %p",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _strip_html(text: str) -> str:
    if not text:
        return ""
    txt = text
    txt = re.sub(r"(?is)<style.*?>.*?</style>", " ", txt)
    txt = re.sub(r"(?is)<script.*?>.*?</script>", " ", txt)
    txt = re.sub(r"(?i)<\s*br\s*/?>", "\n", txt)
    txt = re.sub(r"(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>", "\n", txt)
    txt = re.sub(r"(?is)<[^>]+>", " ", txt)
    txt = _html.unescape(txt)
    return txt


def _extract_body(msg):
    if msg.is_multipart():
        plain_parts = []
        html_parts = []
        for part in msg.walk():
            ctype = part.get_content_type()
            try:
                content = part.get_content()
            except Exception:
                content = None
            if ctype == "text/plain" and isinstance(content, str):
                plain_parts.append(content)
            elif ctype == "text/html" and isinstance(content, str):
                html_parts.append(content)
        plain_text = "\n".join(plain_parts).strip()
        html_text = "\n".join(html_parts).strip()
        return plain_text, html_text
    try:
        content = msg.get_content()
        if isinstance(content, str):
            return content.strip(), ""
    except Exception:
        pass
    return "", ""


def _extract_quoted_blocks(raw: str, subject_norm_value: str):
    txt = _strip_html(raw)
    lines = [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]
    if not lines:
        return []

    def _inline_field(line: str, label: str) -> str:
        m = re.search(rf"(?i)\b{label}\s*:\s*", line)
        if not m:
            return ""
        rest = line[m.end():]
        stop = re.search(r"(?i)\b(from|sent|to|cc|subject|objet)\s*:", rest)
        if stop:
            rest = rest[: stop.start()]
        return f"{label.capitalize()}: {rest.strip()}"

    row_tokens = _match_tokens(subject_norm_value or "")
    blocks = []
    for i, ln in enumerate(lines):
        if "from:" not in ln.lower():
            continue
        from_idx = ln.lower().find("from:")
        from_line = ln[from_idx + 5:].strip()
        cut = re.search(r"(?i)\b(sent|to|cc|subject|objet)\s*:", from_line)
        if cut:
            from_line = from_line[: cut.start()].strip()
        sent_line = _inline_field(ln, "sent")
        subj_line = _inline_field(ln, "subject") or _inline_field(ln, "objet")
        for j in range(i + 1, min(i + 18, len(lines))):
            low = lines[j].lower()
            if low.startswith("from:"):
                break
            if not sent_line and low.startswith("sent:"):
                sent_line = lines[j]
            if not subj_line and (low.startswith("subject:") or low.startswith("objet")):
                subj_line = lines[j]
            if sent_line and subj_line:
                break
        if not sent_line:
            continue
        if subj_line and row_tokens:
            subj_text = re.sub(r"(?i)^(subject|objet)\s*:\s*", "", subj_line).strip()
            subj_norm = normalize_subject(subj_text)
            subj_tokens = _match_tokens(subj_norm)
            score = _token_overlap_score(row_tokens, subj_tokens) if subj_tokens else 0.0
            contains = bool(subject_norm_value and subj_norm and (subject_norm_value in subj_norm or subj_norm in subject_norm_value))
            if score < 0.45 and not contains:
                continue
        sent_dt = _parse_sent_line(sent_line)
        if not sent_dt:
            continue
        blocks.append((from_line, sent_dt))
    return blocks


IST = timezone(timedelta(hours=5, minutes=30))


def _to_ist_naive(dt):
    if not dt:
        return None
    if dt.tzinfo:
        return dt.astimezone(IST).replace(tzinfo=None)
    # If no tzinfo, assume it's already IST (quoted Sent lines)
    return dt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--eml-root", required=True)
    ap.add_argument("--ess-list", default=str(ROOT / "config" / "ess_team.json"))
    ap.add_argument("--ack-mins", type=int, default=16)
    ap.add_argument("--reply-hours", type=int, default=48)
    args = ap.parse_args()

    subject_norm = normalize_subject(args.subject)
    row_tokens = _match_tokens(subject_norm)

    ess = json.loads(Path(args.ess_list).read_text())
    ess_set = {e.strip().lower() for e in ess}

    eml_root = Path(args.eml_root)
    if not eml_root.exists():
        print(f"EML root not found: {eml_root}")
        return 2

    pairs = []
    consultant_replies = []

    for path in eml_root.rglob("*.eml"):
        try:
            with path.open("rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
        except Exception:
            continue

        subj = normalize_subject(msg.get("subject") or "")
        if row_tokens:
            s_tokens = _match_tokens(subj)
            score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
            contains = bool(subject_norm and subj and (subject_norm in subj or subj in subject_norm))
            if score < 0.45 and not contains:
                continue

        from_header = msg.get("from") or ""
        from_lower = from_header.lower()
        if any(em in from_lower for em in ess_set):
            dt = msg.get("date")
            try:
                parsed = email.utils.parsedate_to_datetime(dt) if dt else None
            except Exception:
                parsed = None
            if parsed:
                consultant_replies.append((_to_ist_naive(parsed), path.name, from_header, subj))

        body_plain, body_html = _extract_body(msg)
        raw = (body_plain or "") + "\n" + (body_html or "")

        q_ess = []
        q_non = []
        for from_line, sent_dt in _extract_quoted_blocks(raw, subject_norm):
            addrs = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", from_line, flags=re.I)
            addrs = [a.lower() for a in addrs]
            if any(a in ess_set for a in addrs):
                q_ess.append(sent_dt)
            else:
                q_non.append(sent_dt)

        for ack in sorted(q_ess):
            reqs = [r for r in q_non if r < ack and r.date() == ack.date()]
            if not reqs:
                continue
            req = reqs[-1]
            if (ack - req) <= timedelta(minutes=args.ack_mins):
                pairs.append((req, ack, path.name))

    pairs.sort()
    consultant_replies.sort()

    print("PAIRS:")
    for p in pairs:
        print(p)

    print("\nCONSULTANT REPLIES:")
    for r in consultant_replies:
        print(r)

    if pairs and consultant_replies:
        req, ack, src = pairs[0]
        reply = next((r for r in consultant_replies if r[0] and r[0] >= ack and r[0] <= ack + timedelta(hours=args.reply_hours)), None)
        print("\nCHOSEN:", req, ack, reply)


if __name__ == "__main__":
    raise SystemExit(main())
