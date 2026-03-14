"""
Tool: Episode trace (mirrors QuotedRequestOnly + ESS-only selection)

Usage:
  python support_tracker_py/tools/episode_trace.py \
    --subject "PE-15800052-S8 claim skipped to process in ExpenSys" \
    --eml-root "D:\\Support_Tracker\\DockerOutput\\eml\\My Outlook Data File(1)\\My Outlook Data File(1)\\raviteja.dwarampudi@invenio-solutions.com\\My Team" \
    --requester "Aniket Kumar" \
    --rows 1 --occurrence 1
"""

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
        return (subject or "").strip()


IST = timezone(timedelta(hours=5, minutes=30))


def _to_ist_naive(dt):
    if not dt:
        return None
    if dt.tzinfo:
        return dt.astimezone(IST).replace(tzinfo=None)
    return dt


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
    out = set()
    for p in re.split(r"[^a-z0-9\\-]+", text.lower()):
        if not p:
            continue
        if any(c.isalpha() for c in p) and any(c.isdigit() for c in p):
            out.add(p)
    return {t for t in out if t}


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


def _extract_quoted_blocks_relaxed(raw: str):
    txt = _strip_html(raw)
    lines = [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]
    if not lines:
        return []
    blocks = []
    for i, ln in enumerate(lines):
        if not re.search(r"(?i)\bfrom\b\s*:?", ln):
            continue
        from_idx = re.search(r"(?i)\bfrom\b\s*:?", ln)
        if not from_idx:
            continue
        from_line = ln[from_idx.end():].strip()
        cut = re.search(r"(?i)\b(sent|to|cc|subject|objet)\b\s*:?", from_line)
        if cut:
            from_line = from_line[: cut.start()].strip()
        sent_line = ""
        subj_line = ""
        for j in range(i + 1, min(i + 18, len(lines))):
            low = lines[j].lower()
            if re.search(r"(?i)\bfrom\b\s*:?", low):
                break
            if (not sent_line) and re.search(r"(?i)\bsent\b\s*:?", low):
                sent_line = lines[j]
            if (not subj_line) and re.search(r"(?i)\b(subject|objet)\b\s*:?", low):
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
            subj_text = re.sub(r"(?i)^(subject|objet)\b\s*:?\s*", "", subj_line).strip()
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


def _name_tokens(name: str) -> set:
    if not name:
        return set()
    return {t for t in re.split(r"[^a-z0-9]+", name.lower()) if len(t) >= 3}


def _match_requester(sender_name: str, sender_email: str, requester_name: str) -> bool:
    if not requester_name:
        return True
    rn = requester_name.lower()
    se = (sender_email or "").lower()
    sn = (sender_name or "").lower()
    if rn in se or rn in sn:
        return True
    req_tokens = _name_tokens(rn)
    if not req_tokens:
        return False
    return req_tokens.issubset(_name_tokens(sn)) or req_tokens.issubset(_name_tokens(se))

_ESS_NAME_STOP = {
    "admin", "support", "team", "service", "services", "ops", "operations",
    "enterprise", "ess", "es", "help", "helpdesk", "desk", "mailbox",
    "noreply", "no", "reply", "system", "group", "global",
}

def _ess_name_only(from_line: str, ess_name_tokens: set) -> bool:
    name_blob = re.sub(r"[^a-z0-9]+", " ", (from_line or "").lower()).strip()
    if not name_blob:
        return False
    tokens = {t for t in name_blob.split() if len(t) >= 3 and t not in _ESS_NAME_STOP}
    if not tokens:
        return False
    ess_tokens = {t for t in ess_name_tokens if t not in _ESS_NAME_STOP}
    matches = tokens & ess_tokens
    if not matches:
        return False
    if len(matches) >= 2:
        return True
    tok = next(iter(matches))
    return len(tok) >= 6


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--eml-root", required=True)
    ap.add_argument("--ess-list", default=str(ROOT / "config" / "ess_team.json"))
    ap.add_argument("--requester", default="")
    ap.add_argument("--ack-mins", type=int, default=16)
    ap.add_argument("--rows", type=int, default=1)
    ap.add_argument("--occurrence", type=int, default=1)
    args = ap.parse_args()

    subject_norm = normalize_subject(args.subject)
    row_tokens = _match_tokens(subject_norm)
    row_id_tokens = _id_like_tokens(subject_norm)

    ess = json.loads(Path(args.ess_list).read_text())
    ess_set = {e.strip().lower() for e in ess}
    ess_domains = {"invenio-solutions.com", "inveniolsi.com"}

    eml_root = Path(args.eml_root)
    if not eml_root.exists():
        print(f"EML root not found: {eml_root}")
        return 2

    consultant_msgs = []
    quoted_blocks = []
    merged_msgs = []
    has_live_non_ess = False
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
        sender_email = email.utils.parseaddr(sender)[1].lower()
        sender_domain = sender_email.split("@", 1)[-1] if "@" in sender_email else ""
        is_ess = (sender_email in ess_set) or (sender_domain in ess_domains)
        if is_ess:
            ess_name_tokens |= _name_tokens(sender)
        if sent_dt:
            merged_msgs.append((msg, sender_email, is_ess, _to_ist_naive(sent_dt), eml_path.name))
            if _match_requester(msg.get("from", ""), sender_email, args.requester):
                if not is_ess:
                    has_live_non_ess = True
                else:
                    consultant_msgs.append((_to_ist_naive(sent_dt), eml_path.name, sender, subj))
        body_plain = ""
        body_html = ""
        try:
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
                body_plain = "\n".join(plain_parts).strip()
                body_html = "\n".join(html_parts).strip()
            else:
                content = msg.get_content()
                if isinstance(content, str):
                    body_plain = content.strip()
        except Exception:
            pass
        raw = (body_plain or "") + "\n" + (body_html or "")
        if raw.strip():
            quoted_blocks.extend(
                [(eml_path.name,)+b for b in _extract_quoted_blocks_with_subject(raw)]
            )

    print("HAS_LIVE_NON_ESS:", has_live_non_ess)

    # classify quoted blocks
    def _classify_block(from_line: str) -> bool:
        addr_hits = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}", from_line, flags=re.I)
        if not addr_hits:
            return _ess_name_only(from_line, ess_name_tokens)
        emails_l = [em.lower() for em in addr_hits]
        domains_l = [em.split("@", 1)[-1] for em in emails_l if "@" in em]
        return any(em in ess_set for em in emails_l) or any(d in ess_domains for d in domains_l)

    def _build_pairs(blocks):
        q_non_ess = []
        q_ess = []
        for _fname, from_line, sent_dt, subj_text in blocks:
            q_norm = normalize_subject(subj_text or "")
            if row_id_tokens:
                q_ids = _id_like_tokens(q_norm)
                if q_ids and row_id_tokens.isdisjoint(q_ids):
                    continue
                if (not q_ids) and (not _row_id_in_raw(subj_text, subj_text, row_id_tokens)):
                    continue
            else:
                q_tokens = _match_tokens(q_norm)
                score = _token_overlap_score(row_tokens, q_tokens) if q_tokens else 0.0
                contains = bool(subject_norm and q_norm and (subject_norm in q_norm or q_norm in subject_norm))
                if score < 0.45 and not contains:
                    continue
            if _classify_block(from_line):
                q_ess.append(_to_ist_naive(sent_dt))
            else:
                q_non_ess.append(_to_ist_naive(sent_dt))
        pairs = []
        if q_non_ess and q_ess:
            q_non_ess.sort()
            q_ess.sort()
            for ack_ist in q_ess:
                reqs = [r for r in q_non_ess if r < ack_ist and r.date() == ack_ist.date()]
                if not reqs:
                    continue
                req_ist = reqs[-1]
                if (ack_ist - req_ist) <= timedelta(minutes=args.ack_mins):
                    pairs.append((req_ist, ack_ist))
        return pairs, q_non_ess

    strict_pairs, strict_non_ess = _build_pairs(quoted_blocks)
    relaxed_pairs = []
    relaxed_non_ess = []
    if not strict_pairs:
        # relaxed parse only when strict found nothing
        relaxed = []
        for eml_path in sorted(eml_root.rglob("*.eml")):
            try:
                msg = BytesParser(policy=policy.default).parsebytes(eml_path.read_bytes())
            except Exception:
                continue
            body_plain = ""
            body_html = ""
            try:
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
                    body_plain = "\n".join(plain_parts).strip()
                    body_html = "\n".join(html_parts).strip()
                else:
                    content = msg.get_content()
                    if isinstance(content, str):
                        body_plain = content.strip()
            except Exception:
                pass
            raw = (body_plain or "") + "\n" + (body_html or "")
            if raw.strip():
                relaxed.extend([(eml_path.name,)+b for b in _extract_quoted_blocks_relaxed(raw)])
        relaxed_pairs, relaxed_non_ess = _build_pairs(relaxed)

    consultant_msgs.sort(key=lambda x: x[0])
    pairs = strict_pairs or relaxed_pairs
    quoted_non_ess = strict_non_ess or relaxed_non_ess
    used_hybrid = False
    if (not pairs) and quoted_non_ess:
        quoted_non_ess = sorted(set(quoted_non_ess))
        req_time = quoted_non_ess[-1]
        # Hybrid: quoted non-ESS request + live ESS reply within ack window
        live_ack = None
        for t, name, sender, subj in consultant_msgs:
            if t and t > req_time and (t - req_time) <= timedelta(minutes=args.ack_mins):
                live_ack = t
                break
        if live_ack:
            pairs = [(req_time, live_ack)]
            used_hybrid = True

    print("PAIRS:", pairs)
    print("CONSULTANT_REPLIES:", consultant_msgs)

    occ = max(1, args.occurrence)
    pick_pair = pairs[occ - 1] if len(pairs) >= occ else (pairs[0] if pairs else None)
    replies_after_ack = []
    if pick_pair:
        for t, name, sender, subj in consultant_msgs:
            if t and t >= pick_pair[1]:
                if (t - pick_pair[1]) <= timedelta(hours=48):
                    replies_after_ack.append((t, name, sender, subj))

    resolved_pick = None
    if replies_after_ack:
        if args.rows <= 1:
            resolved_pick = replies_after_ack[-1]
        else:
            resolved_pick = replies_after_ack[min(occ - 1, len(replies_after_ack) - 1)]

    print("CHOSEN:")
    print("  pair:", pick_pair)
    print("  resolved:", resolved_pick)
    print("WOULD_TAG_QUOTED_REQUEST_ONLY:", (not has_live_non_ess) and bool(pairs))
    print("HYBRID_QUOTED_REQUEST_LIVE_ACK:", used_hybrid)


if __name__ == "__main__":
    raise SystemExit(main())
