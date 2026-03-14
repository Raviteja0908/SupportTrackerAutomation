"""
Targeted scan / debug for a single subject using the same core helpers.

Usage:
  python support_tracker_py/tools/targeted_scan.py \
    --subject "PE-15800052-S8 claim skipped to process in ExpenSys" \
    --eml-root "D:\\Support_Tracker\\DockerOutput\\eml\\My Outlook Data File(1)\\My Outlook Data File(1)\\raviteja.dwarampudi@invenio-solutions.com\\My Team" \
    --occurrence 1
"""

import argparse
import email
import html as _html
import json
import re
import sys
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent
ROOT = THIS_DIR.parent
sys.path.append(str(ROOT))

try:
    from src.rules.subject_normalizer import normalize_subject
except Exception:
    def normalize_subject(subject: str) -> str:
        return (subject or "").strip()


IST = timezone(timedelta(hours=5, minutes=30))


def _to_ist_naive(dt):
    if not dt:
        return None
    if dt.tzinfo:
        return dt.astimezone(IST).replace(tzinfo=None)
    return dt


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


def _id_like_tokens(text: str) -> set:
    if not text:
        return set()
    # normalize odd dash variants and zero-width chars
    text = re.sub(r"[â€â€‘â€’â€“â€”â€•âˆ’ï¹£ï¼\u00ad]", "-", text)
    text = re.sub(r"[\u200b\u200c\u200d\u2060\ufeff]", "", text)
    tokens = {
        m.group(0).lower()
        for m in re.finditer(
            r"(?<![a-z0-9])[a-z]{2,}[a-z0-9\-]*\d[a-z0-9\-]*(?![a-z0-9])",
            text.lower(),
        )
    }
    if tokens:
        return tokens
    # fallback: split on whitespace and pick id-like tokens
    out = set()
    for p in re.split(r"\s+", text.lower()):
        if re.search(r"[a-z]{2,}.*\d", p):
            out.add(p.strip("-_.,"))
    return {t for t in out if t}


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
        return "\n".join(plain_parts).strip(), "\n".join(html_parts).strip()
    try:
        content = msg.get_content()
        if isinstance(content, str):
            return content.strip(), ""
    except Exception:
        pass
    return "", ""


def _extract_quoted_blocks_with_subject(raw: str):
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

    blocks = []
    for i, ln in enumerate(lines):
        lower_ln = ln.lower()
        if "from:" not in lower_ln:
            continue
        from_idx = lower_ln.find("from:")
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
        sent_dt = _parse_sent_line(sent_line)
        if not sent_dt:
            continue
        subj_text = ""
        if subj_line:
            subj_text = re.sub(r"(?i)^(subject|objet)\s*:\s*", "", subj_line).strip()
        blocks.append((from_line, sent_dt, subj_text))
    return blocks


def _row_id_in_raw(subject: str, raw: str, row_id_tokens: set) -> bool:
    if not row_id_tokens:
        return True
    txt = f"{subject or ''} {raw or ''}"
    if "<" in txt or "&" in txt:
        txt = _html.unescape(txt)
        txt = re.sub(r"(?is)<[^>]+>", " ", txt)
    txt = txt.lower()
    txt = txt.replace("\u2013", "-").replace("\u2014", "-").replace("\u2011", "-")
    txt = txt.replace("&ndash;", "-").replace("&mdash;", "-").replace("&#8209;", "-")
    tokens = {
        m.group(0)
        for m in re.finditer(r"(?<![a-z0-9])[a-z]{2,}[a-z0-9\-]*\d[a-z0-9\-]*(?![a-z0-9])", txt)
    }
    return bool(tokens & row_id_tokens)


def _is_ess_sender(sender: str, ess_set: set) -> bool:
    s = (sender or "").lower()
    for e in ess_set:
        if e and e in s:
            return True
    return False


def _name_tokens(name: str) -> set:
    if not name:
        return set()
    return {t for t in re.split(r"[^a-z0-9]+", name.lower()) if len(t) >= 3}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--eml-root", required=True)
    ap.add_argument("--ess-list", default=str(ROOT / "config" / "ess_team.json"))
    ap.add_argument("--ack-mins", type=int, default=16)
    ap.add_argument("--occurrence", type=int, default=1, help="1-based occurrence index for duplicate subjects")
    ap.add_argument("--debug", action="store_true", help="Print quoted block extraction details")
    args = ap.parse_args()

    subject_norm = normalize_subject(args.subject)
    row_tokens = _match_tokens(subject_norm)
    row_id_tokens = _id_like_tokens(subject_norm)

    ess = json.loads(Path(args.ess_list).read_text())
    ess_set = {e.strip().lower() for e in ess}

    eml_root = Path(args.eml_root)
    if not eml_root.exists():
        print(f"EML root not found: {eml_root}")
        return 2

    print(f"SUBJECT_NORM: {subject_norm}")
    print(f"ROW_TOKENS: {sorted(row_tokens)}")
    print(f"ROW_ID_TOKENS: {sorted(row_id_tokens)}")

    quoted_blocks = []
    consultant_msgs = []
    ess_name_tokens = set()

    for eml_path in sorted(eml_root.rglob("*.eml")):
        try:
            msg = BytesParser(policy=policy.default).parsebytes(eml_path.read_bytes())
        except Exception:
            continue
        subj = msg.get("subject", "") or ""
        subj_norm = normalize_subject(subj)
        subj_tokens = _match_tokens(subj_norm)
        score = _token_overlap_score(row_tokens, subj_tokens) if subj_tokens else 0.0
        contains = bool(subject_norm and subj_norm and (subject_norm in subj_norm or subj_norm in subject_norm))
        if row_tokens and score < 0.45 and not contains:
            continue
        sent = msg.get("date")
        try:
            sent_dt = email.utils.parsedate_to_datetime(sent) if sent else None
        except Exception:
            sent_dt = None
        sender = (msg.get("from") or "")
        if sent_dt:
            if _is_ess_sender(sender, ess_set):
                consultant_msgs.append((_to_ist_naive(sent_dt), eml_path.name, sender, subj))
                ess_name_tokens |= _name_tokens(sender)
        plain, html = _extract_body(msg)
        raw = plain or ""
        if not raw and html:
            raw = _strip_html(html)
        if raw:
            blocks = _extract_quoted_blocks_with_subject(raw)
            for from_line, sent_dt, subj_text in blocks:
                quoted_blocks.append((eml_path.name, from_line, sent_dt, subj_text))

    # Build quoted request+ack pairs
    pairs_all = []
    if args.debug:
        print(f"QUOTED_BLOCKS_TOTAL: {len(quoted_blocks)}")
        for eml_name, from_line, sent_dt, subj_text in quoted_blocks[:30]:
            subj_norm_dbg = normalize_subject(subj_text or "")
            has_id = bool(_id_like_tokens(subj_norm_dbg) & row_id_tokens) if subj_norm_dbg else False
            print(
                f"  BLOCK file={eml_name} sent={sent_dt} from='{from_line}' subj='{subj_text}' "
                f"id_match={has_id}"
            )

    for eml_name, from_line, sent_dt, subj_text in quoted_blocks:
        subj_norm = normalize_subject(subj_text or "")
        subj_tokens = _match_tokens(subj_norm)
        score = _token_overlap_score(row_tokens, subj_tokens) if subj_tokens else 0.0
        contains = bool(subject_norm and subj_norm and (subject_norm in subj_norm or subj_norm in subject_norm))
        if row_tokens and score < 0.45 and not contains:
            continue
        if row_id_tokens and not _row_id_in_raw(subj_text, subj_text, row_id_tokens):
            continue
        sender = from_line or ""
        is_ess = _is_ess_sender(sender, ess_set)
        if (not is_ess) and (not re.search(r"@[A-Z0-9._%+-]+", sender, flags=re.I)):
            # name-only: match against ESS name tokens extracted from real senders
            if _name_tokens(sender) & ess_name_tokens:
                is_ess = True
        pairs_all.append((eml_name, sent_dt, sender, subj_text, is_ess))

    reqs = [p for p in pairs_all if not p[4]]
    acks = [p for p in pairs_all if p[4]]

    pairs = []
    for r in reqs:
        r_dt = r[1]
        for a in acks:
            a_dt = a[1]
            if not (r_dt and a_dt):
                continue
            if r_dt.date() != a_dt.date():
                continue
            gap = abs((a_dt - r_dt).total_seconds()) / 60.0
            if gap <= args.ack_mins:
                pairs.append((r_dt, a_dt, r, a))

    pairs = sorted(pairs, key=lambda x: (x[1], x[0]))
    consultant_msgs = sorted(consultant_msgs, key=lambda x: x[0] or datetime.min)

    print("PAIRS:")
    for r_dt, a_dt, r, a in pairs:
        print(f"  req={r_dt} ack={a_dt} req_from='{r[2]}' ack_from='{a[2]}' file={r[0]}")

    print("CONSULTANT REPLIES:")
    for t, name, sender, subj in consultant_msgs:
        print(f"  {t} | {sender} | {name} | {subj}")

    # Choose pair based on occurrence and consultant reply proximity
    occ = max(1, args.occurrence)
    pick_pair = pairs[occ - 1] if len(pairs) >= occ else (pairs[0] if pairs else None)
    pick_cons = consultant_msgs[occ - 1] if len(consultant_msgs) >= occ else (consultant_msgs[0] if consultant_msgs else None)

    print("CHOSEN:")
    if pick_pair:
        print(f"  req={pick_pair[0]} ack={pick_pair[1]}")
    else:
        print("  req=None ack=None")
    if pick_cons:
        print(f"  reply={pick_cons[0]} sender={pick_cons[2]} file={pick_cons[1]}")
    else:
        print("  reply=None")


if __name__ == "__main__":
    raise SystemExit(main())
