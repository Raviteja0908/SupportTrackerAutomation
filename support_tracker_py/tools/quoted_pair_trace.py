# Tool: Quoted-pair trace (mirrors main.py quoted-pair filtering)
# Usage:
#   python support_tracker_py/tools/quoted_pair_trace.py --subject "PE-15800052-S8 claim skipped to process in ExpenSys" \
#     --eml-root "D:\Support_Tracker\DockerOutput\eml\My Outlook Data File(1)\My Outlook Data File(1)\raviteja.dwarampudi@invenio-solutions.com\My Team"

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
    def normalize_subject(s: str) -> str:
        return (s or "").strip()


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
    text = re.sub(r"[‐‑‒–—―−﹣－\u00ad]", "-", text)
    text = re.sub(r"[\u200b\u200c\u200d\u2060\ufeff]", "", text)
    return {
        m.group(0).lower()
        for m in re.finditer(r"(?<![a-z0-9])[a-z]{2,}[a-z0-9\-]*\d[a-z0-9\-]*(?![a-z0-9])", text.lower())
    }


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


IST = timezone(timedelta(hours=5, minutes=30))


def _to_ist_naive(dt):
    if not dt:
        return None
    if dt.tzinfo:
        return dt.astimezone(IST).replace(tzinfo=None)
    return dt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--eml-root", required=True)
    ap.add_argument("--ess-list", default=str(ROOT / "config" / "ess_team.json"))
    args = ap.parse_args()

    subject_norm = normalize_subject(args.subject)
    row_tokens = _match_tokens(subject_norm)
    print("SUBJECT_NORM:", repr(subject_norm))
    print("ID_TOKENS:", _id_like_tokens(subject_norm))
    row_id_tokens = _id_like_tokens(subject_norm)

    ess = json.loads(Path(args.ess_list).read_text())
    ess_set = {e.strip().lower() for e in ess}
    ess_domains = {"invenio-solutions.com", "inveniolsi.com"}

    eml_root = Path(args.eml_root)
    if not eml_root.exists():
        print(f"EML root not found: {eml_root}")
        return 2

    for path in sorted(eml_root.rglob("*.eml")):
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

        body_plain, body_html = _extract_body(msg)
        raw = (body_plain or "") + "\n" + (body_html or "")
        blocks = _extract_quoted_blocks_with_subject(raw)
        if not blocks:
            continue

        print(f"\nFILE: {path.name}")
        for from_line, sent_dt, q_subj in blocks:
            q_norm = normalize_subject(q_subj or "")
            q_ids = _id_like_tokens(q_norm) if q_subj else set()
            id_ok = True
            if row_id_tokens:
                if q_subj:
                    id_ok = bool(q_ids and not row_id_tokens.isdisjoint(q_ids))
                else:
                    id_ok = False
            if not id_ok:
                print(f"  SKIP (ID): {sent_dt} | {from_line} | subj='{q_subj}'")
                continue

            addr_hits = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}", from_line, flags=re.I)
            emails_l = [em.lower() for em in addr_hits]
            domains_l = [em.split("@", 1)[-1] for em in emails_l if "@" in em]
            is_ess = any(em in ess_set for em in emails_l) or any(d in ess_domains for d in domains_l)
            role = "ESS" if is_ess else "NON-ESS"
            print(f"  OK ({role}) {sent_dt} | {from_line} | subj='{q_subj}'")


if __name__ == "__main__":
    raise SystemExit(main())
