import argparse
import csv
import json
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Iterable

from src.rules.subject_normalizer import extract_subject_from_description, normalize_subject, normalize_subject_for_match
from src.rules.time_resolver import (
    _classify_reply_kind,
    _extract_canonical_message_lines,
    _is_ess_sender,
    _match_requester,
    _to_ist,
)

try:
    from dateutil.parser import parse as _dateutil_parse
except Exception:
    _dateutil_parse = None


@dataclass
class DebugEmail:
    subject: str
    sender_name: str
    sender_email: str
    sent_time: datetime | None
    body: str
    body_html: str
    path: Path


@dataclass
class QuotedHeaderCandidate:
    from_line: str
    quoted_ist: datetime
    quoted_subj: str
    raw_sent: str
    normalized_sent: str
    has_am_pm: bool
    am_or_pm: str
    is_ambiguous: bool
    header_lines: tuple[str, ...]


def _read_json_list(path: Path) -> list[str]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []


def _get_col(row: dict, *names: str) -> str:
    for name in names:
        if name in row and row[name] not in (None, ""):
            return str(row[name]).strip()
    lowered = {str(k).strip().lower(): v for k, v in row.items()}
    for name in names:
        v = lowered.get(name.lower())
        if v not in (None, ""):
            return str(v).strip()
    return ""


def _load_csv_rows(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = []
        for idx, row in enumerate(reader, start=2):
            row["_line"] = idx
            rows.append(row)
        return rows


def _match_rows(rows: list[dict], needles: list[str]) -> list[dict]:
    out = []
    for row in rows:
        desc = _get_col(row, "Description")
        if any(n.lower() in desc.lower() for n in needles):
            out.append(row)
    return out


def _family_subject_norm(row: dict) -> str:
    desc = _get_col(row, "Description")
    return normalize_subject(extract_subject_from_description(desc))


def _fmt(dt: datetime | None) -> str:
    if not dt:
        return "-"
    ist = _to_ist(dt)
    return ist.strftime("%d-%m-%Y %H:%M") if ist else "-"


def _extract_body(msg) -> tuple[str, str]:
    plain = ""
    html = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                ctype = (part.get_content_type() or "").lower()
                if ctype == "text/plain" and not plain:
                    plain = part.get_content()
                elif ctype == "text/html" and not html:
                    html = part.get_content()
        else:
            ctype = (msg.get_content_type() or "").lower()
            if ctype == "text/plain":
                plain = msg.get_content()
            elif ctype == "text/html":
                html = msg.get_content()
    except Exception:
        pass
    return plain or "", html or ""


def _parse_eml(path: Path) -> DebugEmail | None:
    try:
        msg = BytesParser(policy=policy.default).parse(path.open("rb"))
    except Exception:
        return None
    try:
        sent = parsedate_to_datetime(msg.get("Date")) if msg.get("Date") else None
    except Exception:
        sent = None
    sender_name = ""
    sender_email = ""
    try:
        addrs = getaddresses([msg.get("From", "")])
        if addrs:
            sender_name, sender_email = addrs[0]
    except Exception:
        pass
    body, body_html = _extract_body(msg)
    return DebugEmail(
        subject=str(msg.get("Subject", "") or ""),
        sender_name=sender_name or "",
        sender_email=sender_email or "",
        sent_time=sent,
        body=body,
        body_html=body_html,
        path=path,
    )


def _iter_emails(eml_dir: Path) -> Iterable[DebugEmail]:
    for path in eml_dir.rglob("*.eml"):
        parsed = _parse_eml(path)
        if parsed:
            yield parsed


def _match_tokens(text: str) -> set[str]:
    if not text:
        return set()
    norm = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return {part for part in norm.split() if part}


def _token_overlap_score(left: set[str], right: set[str]) -> float:
    if not left or not right:
        return 0.0
    inter = len(left & right)
    if inter <= 0:
        return 0.0
    return inter / float(max(len(left), len(right)))


def _id_like_tokens(text: str) -> set[str]:
    if not text:
        return set()
    text = text.lower()
    tokens = {
        m.group(0)
        for m in re.finditer(r"(?<![a-z0-9])[a-z]{2,}[a-z0-9\\-]*\\d[a-z0-9\\-]*(?![a-z0-9])", text)
    }
    if not tokens:
        for part in re.split(r"[^a-z0-9\\-]+", text):
            if part and any(c.isalpha() for c in part) and any(c.isdigit() for c in part):
                tokens.add(part)
    return tokens


def _row_subject_match(subject_norm_value: str, row_tokens: set[str], row_id_tokens: set[str], email_subject: str) -> bool:
    email_norm = normalize_subject(email_subject or "")
    if not subject_norm_value or not email_norm:
        return False
    if row_id_tokens:
        email_ids = _id_like_tokens(email_norm)
        if not email_ids or row_id_tokens.isdisjoint(email_ids):
            return False
    if not row_tokens:
        return True
    email_tokens = _match_tokens(email_norm)
    score = _token_overlap_score(row_tokens, email_tokens) if email_tokens else 0.0
    contains = subject_norm_value in email_norm or email_norm in subject_norm_value
    if score >= 0.45 or contains:
        return True
    subj_match = normalize_subject_for_match(subject_norm_value)
    email_match = normalize_subject_for_match(email_norm)
    if not subj_match or not email_match:
        return False
    return subj_match == email_match or subj_match in email_match or email_match in subj_match


def _system_like_sender(email: DebugEmail) -> bool:
    sender = (email.sender_email or email.sender_name or "").lower()
    return any(token in sender for token in ("noreply", "no-reply", "daemon", "postmaster", "system"))


def _parse_quoted_sent_time(sent_line: str) -> datetime | None:
    if not sent_line:
        return None
    text = re.sub(r"(?i)^sent\s*:\s*", "", sent_line).strip()
    text = re.sub(r"(?i)^(mon|tue|wed|thu|fri|sat|sun)\w*,?\s*", "", text).strip()
    text = re.sub(r"\(.*?\)", " ", text).strip()
    text = re.sub(r"(?i)(\d)(am|pm)\b", r"\1 \2", text)
    text = re.sub(r"(?i)\b(a\.m\.|p\.m\.)\b", lambda m: m.group(1).replace(".", "").upper(), text)
    normalized = " ".join(text.replace(",", " ").replace(" at ", " ").split())
    candidates = []
    for candidate in (text, normalized):
        if candidate and candidate not in candidates:
            candidates.append(candidate)
    has_am_pm = bool(re.search(r"(?i)\b(am|pm|a\.m\.|p\.m\.)\b", sent_line or ""))
    fmts_24h = [
        "%d %B %Y %H:%M:%S",
        "%d %B %Y %H:%M",
        "%d %b %Y %H:%M:%S",
        "%d %b %Y %H:%M",
        "%B %d %Y %H:%M:%S",
        "%B %d %Y %H:%M",
        "%b %d %Y %H:%M:%S",
        "%b %d %Y %H:%M",
        "%d %B %Y %I:%M:%S %p",
        "%d %B %Y %I:%M %p",
        "%d %b %Y %I:%M:%S %p",
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
    fmts_12h = [
        "%d %B %Y %I:%M:%S %p",
        "%d %B %Y %I:%M %p",
        "%d %b %Y %I:%M:%S %p",
        "%d %b %Y %I:%M %p",
        "%B %d %Y %I:%M:%S %p",
        "%B %d %Y %I:%M %p",
        "%b %d %Y %I:%M:%S %p",
        "%b %d %Y %I:%M %p",
        "%d-%m-%Y %I:%M:%S %p",
        "%d-%m-%Y %I:%M %p",
        "%d.%m.%Y %I:%M:%S %p",
        "%d.%m.%Y %I:%M %p",
        "%d/%m/%Y %I:%M:%S %p",
        "%d/%m/%Y %I:%M %p",
        "%Y-%m-%d %I:%M:%S %p",
        "%Y-%m-%d %I:%M %p",
    ]
    parse_fmts = fmts_12h if has_am_pm else fmts_24h
    for candidate in candidates:
        for fmt in parse_fmts:
            try:
                return datetime.strptime(candidate, fmt)
            except Exception:
                continue
        try:
            return parsedate_to_datetime(candidate)
        except Exception:
            pass
        if has_am_pm:
            for fmt in fmts_24h:
                try:
                    return datetime.strptime(candidate, fmt)
                except Exception:
                    continue
    return None


def _normalize_quoted_sent_text(sent_line: str) -> str:
    text = re.sub(r"(?i)^sent\s*:\s*", "", sent_line or "").strip()
    text = re.sub(r"(?i)^(mon|tue|wed|thu|fri|sat|sun)\w*,?\s*", "", text).strip()
    text = re.sub(r"\(.*?\)", " ", text).strip()
    text = re.sub(r"(?i)(\d)(am|pm)\b", r"\1 \2", text)
    text = re.sub(r"(?i)\b(a\.m\.|p\.m\.)\b", lambda m: m.group(1).replace(".", "").upper(), text)
    return " ".join(text.replace(",", " ").replace(" at ", " ").split())


def _parse_quoted_sent_time_strict(sent_line: str) -> datetime | None:
    if not sent_line:
        return None
    normalized = _normalize_quoted_sent_text(sent_line)
    candidates = []
    for candidate in (sent_line, normalized):
        candidate = candidate.strip()
        if candidate and candidate not in candidates:
            candidates.append(candidate)
    for candidate in candidates:
        if _dateutil_parse is not None:
            try:
                return _dateutil_parse(candidate, fuzzy=False, dayfirst=True)
            except Exception:
                pass
        try:
            return parsedate_to_datetime(candidate)
        except Exception:
            pass
    return _parse_quoted_sent_time(sent_line)


def _quoted_sent_metadata(sent_line: str) -> tuple[str, bool, str, bool]:
    raw_sent = (sent_line or "").strip()
    normalized_sent = _normalize_quoted_sent_text(raw_sent)
    meridiem_match = re.search(r"(?i)\b(am|pm|a\.m\.|p\.m\.)\b", raw_sent)
    am_or_pm = ""
    has_am_pm = False
    if meridiem_match:
        token = meridiem_match.group(1).lower().replace(".", "")
        am_or_pm = "pm" if token == "pm" else "am"
        has_am_pm = True
    time_match = re.search(r"\b(\d{1,2}):(\d{2})(?::(\d{2}))?\b", normalized_sent)
    is_ambiguous = False
    if time_match:
        hour = int(time_match.group(1))
        is_ambiguous = (1 <= hour <= 12) and (not has_am_pm)
    return normalized_sent, has_am_pm, am_or_pm, is_ambiguous


def _ess_name_only(from_line: str, ess_team: list[str]) -> bool:
    ess_tokens = set()
    for item in ess_team:
        item_l = (item or "").lower()
        for token in re.split(r"[^a-z0-9]+", item_l):
            if len(token) >= 3 and token not in {"admin", "support", "team", "service", "mailbox", "noreply", "reply"}:
                ess_tokens.add(token)
    from_tokens = {
        token
        for token in re.split(r"[^a-z0-9]+", (from_line or "").lower())
        if len(token) >= 3 and token not in {"admin", "support", "team", "service", "mailbox", "noreply", "reply"}
    }
    if not from_tokens or not ess_tokens:
        return False
    matches = from_tokens & ess_tokens
    if len(matches) >= 2:
        return True
    return any(len(token) >= 6 for token in matches)


def _clean_message_text(email: DebugEmail) -> list[str]:
    return _extract_canonical_message_lines(email)


def _extract_quoted_blocks_with_subject(email: DebugEmail) -> list[tuple[str, datetime, str]]:
    lines = _clean_message_text(email)
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
        if sent_line:
            sent_low = sent_line.lower()
            if (" am" not in sent_low) and (" pm" not in sent_low):
                for j in range(i + 1, min(i + 4, len(lines))):
                    extra = (lines[j] or "").strip()
                    if re.fullmatch(r"(?i)(am|pm|a\.m\.|p\.m\.)", extra):
                        sent_line = f"{sent_line} {extra}"
                        break
        if not sent_line:
            continue
        sent_dt = _parse_quoted_sent_time(sent_line)
        if not sent_dt:
            continue
        subj_text = ""
        if subj_line:
            subj_text = re.sub(r"(?i)^(subject|objet)\s*:\s*", "", subj_line).strip()
        blocks.append((from_line, _to_ist(sent_dt), subj_text))
    return blocks


def _extract_quoted_blocks_relaxed(email: DebugEmail) -> list[tuple[str, datetime, str]]:
    lines = _clean_message_text(email)
    if not lines:
        return []
    blocks = []
    for i, ln in enumerate(lines):
        from_match = re.search(r"(?i)\bfrom\b\s*:?", ln)
        if not from_match:
            continue
        from_line = ln[from_match.end():].strip()
        cut = re.search(r"(?i)\b(sent|to|cc|subject|objet)\b\s*:?", from_line)
        if cut:
            from_line = from_line[: cut.start()].strip()
        sent_line = ""
        subj_line = ""
        for j in range(i + 1, min(i + 18, len(lines))):
            low = lines[j].lower()
            if re.search(r"(?i)\bfrom\b\s*:?", lines[j]):
                break
            if not sent_line and re.match(r"(?i)^sent\b\s*:?", lines[j]):
                sent_line = lines[j]
            if not subj_line and re.match(r"(?i)^(subject|objet)\b\s*:?", lines[j]):
                subj_line = lines[j]
            if sent_line:
                break
        if sent_line:
            sent_low = sent_line.lower()
            if (" am" not in sent_low) and (" pm" not in sent_low):
                for j in range(i + 1, min(i + 4, len(lines))):
                    extra = (lines[j] or "").strip()
                    if re.fullmatch(r"(?i)(am|pm|a\.m\.|p\.m\.)", extra):
                        sent_line = f"{sent_line} {extra}"
                        break
        if not sent_line:
            continue
        sent_dt = _parse_quoted_sent_time(sent_line)
        if not sent_dt:
            continue
        subj_text = re.sub(r"(?i)^(subject|objet)\s*:?", "", subj_line).strip() if subj_line else ""
        blocks.append((from_line, _to_ist(sent_dt), subj_text))
    return blocks


def _extract_quoted_blocks_header_bounded(email: DebugEmail) -> list[QuotedHeaderCandidate]:
    lines = _clean_message_text(email)
    if not lines:
        return []

    def _append_meridiem(sent_line: str, start_idx: int) -> str:
        sent_low = sent_line.lower()
        if (" am" in sent_low) or (" pm" in sent_low):
            return sent_line
        for j in range(start_idx + 1, min(start_idx + 4, len(lines))):
            extra = (lines[j] or "").strip()
            if re.fullmatch(r"(?i)(am|pm|a\.m\.|p\.m\.)", extra):
                return f"{sent_line} {extra}"
        return sent_line

    blocks: list[QuotedHeaderCandidate] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        low = line.lower()
        if "from:" not in low and not re.search(r"(?i)\bfrom\b\s*:?", line):
            i += 1
            continue

        from_line = ""
        sent_line = ""
        subj_line = ""
        header_end = i
        for j in range(i, min(i + 12, len(lines))):
            cur = (lines[j] or "").strip()
            cur_low = cur.lower()
            if j > i and ("from:" in cur_low or re.match(r"(?i)^[-_]{3,}$", cur)):
                break
            if (not from_line) and re.search(r"(?i)\bfrom\b\s*:", cur):
                from_line = re.sub(r"(?i)^.*?\bfrom\b\s*:\s*", "", cur).strip()
                cut = re.search(r"(?i)\b(sent|to|cc|subject|objet)\b\s*:", from_line)
                if cut:
                    from_line = from_line[: cut.start()].strip()
            if (not sent_line) and re.search(r"(?i)\bsent\b\s*:", cur):
                sent_line = cur
            if (not subj_line) and re.search(r"(?i)\b(subject|objet)\b\s*:", cur):
                subj_line = cur
            header_end = j
            if from_line and sent_line and subj_line:
                break

        if sent_line:
            sent_line = _append_meridiem(sent_line, header_end)
            sent_dt = _parse_quoted_sent_time_strict(sent_line)
            if sent_dt:
                subj_text = re.sub(r"(?i)^(subject|objet)\s*:\s*", "", subj_line).strip() if subj_line else ""
                normalized_sent, has_am_pm, am_or_pm, is_ambiguous = _quoted_sent_metadata(sent_line)
                header_lines = tuple((lines[k] or "").strip() for k in range(i, min(header_end + 1, len(lines))))
                blocks.append(
                    QuotedHeaderCandidate(
                        from_line=from_line,
                        quoted_ist=_to_ist(sent_dt),
                        quoted_subj=subj_text,
                        raw_sent=sent_line.strip(),
                        normalized_sent=normalized_sent,
                        has_am_pm=has_am_pm,
                        am_or_pm=am_or_pm,
                        is_ambiguous=is_ambiguous,
                        header_lines=header_lines,
                    )
                )
        i = max(i + 1, header_end + 1)
    return blocks


def _requester_match_any(email: DebugEmail, requesters: list[str]) -> bool:
    return any(_match_requester(email.sender_name, email.sender_email, req) for req in requesters if req)


def _minute_dedupe(pairs: list[tuple[datetime, object]]) -> list[tuple[datetime, object]]:
    seen = set()
    out = []
    for when, payload in sorted(pairs, key=lambda x: x[0]):
        minute_key = _to_ist(when).replace(second=0, microsecond=0)
        if minute_key in seen:
            continue
        seen.add(minute_key)
        out.append((when, payload))
    return out


def _best_before(candidates: list[tuple[datetime, object]], upper_ist: datetime, max_gap: timedelta) -> tuple[datetime, object] | None:
    usable = []
    for when, payload in candidates:
        if when >= upper_ist:
            continue
        if max_gap and (upper_ist - when) > max_gap:
            continue
        usable.append((when, payload))
    if not usable:
        return None
    strict = [entry for entry in usable if entry[0].replace(second=0, microsecond=0) != upper_ist.replace(second=0, microsecond=0)]
    return strict[-1] if strict else usable[-1]


def _subject_variant_score(row_subject: str, candidate_subject: str) -> int:
    row_norm = normalize_subject(row_subject or "")
    candidate_norm = normalize_subject(candidate_subject or "")
    row_match = normalize_subject_for_match(row_norm)
    candidate_match = normalize_subject_for_match(candidate_norm)
    if not row_match or not candidate_match:
        return 0
    if row_match == candidate_match:
        return 100
    if row_match in candidate_match or candidate_match in row_match:
        return 85
    row_tokens = _match_tokens(row_match)
    candidate_tokens = _match_tokens(candidate_match)
    if not row_tokens or not candidate_tokens:
        return 0
    inter = len(row_tokens & candidate_tokens)
    if inter <= 0:
        return 0
    extra = len(candidate_tokens - row_tokens)
    missing = len(row_tokens - candidate_tokens)
    return max(0, inter * 14 - extra * 8 - missing * 6)


def _reply_anchored_preferred_episode(
    row: dict,
    live_requests: list[tuple[datetime, DebugEmail]],
    quoted_requests: list[tuple[datetime, tuple[DebugEmail, str, str]]],
    strict_quoted_requests: list[tuple[datetime, tuple[DebugEmail, QuotedHeaderCandidate]]],
    reply_candidates: list[tuple[datetime, DebugEmail]],
) -> dict | None:
    requester = _get_col(row, "Requester", "Consultant")
    row_subject = _family_subject_norm(row)
    row_tokens = _match_tokens(row_subject)
    row_id_tokens = _id_like_tokens(row_subject)
    eligible_replies = []
    for reply_ist, reply_email in reply_candidates:
        if requester and not _match_requester(reply_email.sender_name, reply_email.sender_email, requester):
            continue
        if not _row_subject_match(row_subject, row_tokens, row_id_tokens, reply_email.subject):
            continue
        reply_subject_score = _subject_variant_score(row_subject, reply_email.subject)
        eligible_replies.append((reply_ist, reply_email, reply_subject_score))
    if not eligible_replies:
        return None

    best_episode = None
    best_score = None
    for reply_ist, reply_email, reply_subject_score in eligible_replies:
        request_options = []
        for req_ist, req_email in live_requests:
            if req_ist >= reply_ist or (reply_ist - req_ist) > timedelta(hours=48):
                continue
            subject_score = _subject_variant_score(row_subject, req_email.subject)
            if subject_score < 60:
                continue
            gap_seconds = int((reply_ist - req_ist).total_seconds())
            request_options.append(
                ((subject_score, 3, -gap_seconds, int(req_ist.timestamp())), "live", (req_ist, req_email))
            )
        for req_ist, payload in strict_quoted_requests:
            source_email, header_candidate = payload
            if req_ist >= reply_ist or (reply_ist - req_ist) > timedelta(hours=48):
                continue
            subject_score = _subject_variant_score(row_subject, header_candidate.quoted_subj or source_email.subject)
            if subject_score < 60:
                continue
            gap_seconds = int((reply_ist - req_ist).total_seconds())
            amp_bonus = 1 if header_candidate.has_am_pm else 0
            request_options.append(
                ((subject_score, 2, amp_bonus, -gap_seconds, int(req_ist.timestamp())), "quoted-strict", (req_ist, payload))
            )
        for req_ist, payload in quoted_requests:
            source_email, from_line, quoted_subj = payload
            if req_ist >= reply_ist or (reply_ist - req_ist) > timedelta(hours=48):
                continue
            subject_score = _subject_variant_score(row_subject, quoted_subj or source_email.subject)
            if subject_score < 60:
                continue
            gap_seconds = int((reply_ist - req_ist).total_seconds())
            request_options.append(
                ((subject_score, 1, -gap_seconds, int(req_ist.timestamp())), "quoted", (req_ist, payload))
            )
        if not request_options:
            continue
        best_request = max(request_options, key=lambda item: item[0])
        episode = {
            "request": best_request[2],
            "request_kind": f"{best_request[1]}-reply-anchored",
            "ack": (reply_ist, reply_email),
            "resolved": (reply_ist, reply_email),
        }
        episode_score = (
            reply_subject_score,
            best_request[0],
            int(reply_ist.timestamp()),
        )
        if best_score is None or episode_score > best_score:
            best_score = episode_score
            best_episode = episode
    return best_episode


def _find_family_rows(debug_rows: list[dict]) -> dict[str, list[dict]]:
    families: dict[str, list[dict]] = {}
    for row in debug_rows:
        row_norm = _family_subject_norm(row)
        if not row_norm:
            continue
        chosen_key = None
        for family_key in families:
            if _row_subject_match(family_key, _match_tokens(family_key), _id_like_tokens(family_key), row_norm):
                chosen_key = family_key
                break
        if chosen_key is None:
            chosen_key = row_norm
            families[chosen_key] = []
        families[chosen_key].append(row)
    return families


def _print_rows(title: str, rows: list[dict]) -> None:
    print(title)
    if not rows:
        print("  <none>")
        return
    for row in rows:
        print(
            f"  line={row.get('_line')} | requester={_get_col(row, 'Requester', 'Consultant')} | "
            f"triplet={_get_col(row, 'Created Date & Time') or '-'} / "
            f"{_get_col(row, 'Actual Response Date & Time') or '-'} / "
            f"{_get_col(row, 'Actual Resolved Date & Time') or '-'}"
        )
        print(f"    desc={_get_col(row, 'Description')}")
        notes = _get_col(row, "Notes")
        if notes:
            print(f"    notes={notes}")


def _print_email_block(title: str, emails: list[DebugEmail], ess_team: list[str], requesters: list[str]) -> None:
    print(title)
    if not emails:
        print("  <empty>")
        return
    for idx, email in enumerate(sorted(emails, key=lambda e: e.sent_time or datetime.max), start=1):
        cls = _classify_reply_kind(email)
        print(
            f"  {idx}. {_fmt(email.sent_time)} | kind={cls.get('kind')} | "
            f"real={bool(cls.get('real_reply'))} | ack={bool(cls.get('ack_like') or cls.get('explicit_ack') or cls.get('short_ess_ack'))} | "
            f"ess={_is_ess_sender(email, ess_team)} | requester_match={_requester_match_any(email, requesters)}"
        )
        print(f"     from={email.sender_email or email.sender_name}")
        print(f"     subj={email.subject}")


def _simulate_family(
    family_rows: list[dict],
    matched_emails: list[DebugEmail],
    ess_team: list[str],
    requesters: list[str],
) -> None:
    live_requests: list[tuple[datetime, DebugEmail]] = []
    quoted_requests: list[tuple[datetime, tuple[DebugEmail, str, str]]] = []
    strict_quoted_requests: list[tuple[datetime, tuple[DebugEmail, QuotedHeaderCandidate]]] = []
    ack_candidates: list[tuple[datetime, DebugEmail]] = []
    reply_candidates: list[tuple[datetime, DebugEmail]] = []

    for email in matched_emails:
        email_ist = _to_ist(email.sent_time) if email.sent_time else None
        if not email_ist:
            continue
        cls = _classify_reply_kind(email)
        is_ess = _is_ess_sender(email, ess_team)
        req_match = _requester_match_any(email, requesters)
        if (not is_ess) and (not _system_like_sender(email)):
            live_requests.append((email_ist, email))
        if is_ess and (cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack") or (not req_match)):
            ack_candidates.append((email_ist, email))
        if is_ess and (not cls.get("ack_like")) and (not cls.get("explicit_ack")) and (not cls.get("short_ess_ack")) and (not cls.get("thanks_info")) and (not cls.get("nonfinal_followup")):
            reply_candidates.append((email_ist, email))
        quoted_blocks = _extract_quoted_blocks_with_subject(email)
        if not quoted_blocks:
            quoted_blocks = _extract_quoted_blocks_relaxed(email)
        for from_line, quoted_ist, quoted_subj in quoted_blocks:
            if quoted_ist >= email_ist:
                continue
            if (email_ist - quoted_ist) > timedelta(hours=48):
                continue
            quoted_norm = normalize_subject(quoted_subj or quoted_subj or "")
            if quoted_subj and not _row_subject_match(family_rows[0] and _family_subject_norm(family_rows[0]) or "", _match_tokens(_family_subject_norm(family_rows[0]) if family_rows else ""), _id_like_tokens(_family_subject_norm(family_rows[0]) if family_rows else ""), quoted_subj):
                continue
            addr_hits = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", from_line, flags=re.I)
            emails_l = [addr.lower() for addr in addr_hits]
            domains_l = [addr.split("@", 1)[-1] for addr in emails_l if "@" in addr]
            is_quoted_ess = False
            if emails_l:
                is_quoted_ess = any(_is_ess_sender(DebugEmail("", "", addr, None, "", "", Path()), ess_team) for addr in emails_l)
            else:
                is_quoted_ess = _ess_name_only(from_line, ess_team)
            if is_quoted_ess or any(domain.endswith("invenio-solutions.com") for domain in domains_l):
                continue
            quoted_requests.append((quoted_ist, (email, from_line, quoted_subj)))

        strict_blocks = _extract_quoted_blocks_header_bounded(email)
        for header_candidate in strict_blocks:
            from_line = header_candidate.from_line
            quoted_ist = header_candidate.quoted_ist
            quoted_subj = header_candidate.quoted_subj
            if quoted_ist >= email_ist:
                continue
            if (email_ist - quoted_ist) > timedelta(hours=48):
                continue
            quoted_norm = normalize_subject(quoted_subj or quoted_subj or "")
            if quoted_subj and not _row_subject_match(family_rows[0] and _family_subject_norm(family_rows[0]) or "", _match_tokens(_family_subject_norm(family_rows[0]) if family_rows else ""), _id_like_tokens(_family_subject_norm(family_rows[0]) if family_rows else ""), quoted_subj):
                continue
            addr_hits = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", from_line, flags=re.I)
            emails_l = [addr.lower() for addr in addr_hits]
            domains_l = [addr.split("@", 1)[-1] for addr in emails_l if "@" in addr]
            is_quoted_ess = False
            if emails_l:
                is_quoted_ess = any(_is_ess_sender(DebugEmail("", "", addr, None, "", "", Path()), ess_team) for addr in emails_l)
            else:
                is_quoted_ess = _ess_name_only(from_line, ess_team)
            if is_quoted_ess or any(domain.endswith("invenio-solutions.com") for domain in domains_l):
                continue
            strict_quoted_requests.append((quoted_ist, (email, header_candidate)))

    live_requests = _minute_dedupe(live_requests)
    quoted_requests = _minute_dedupe(quoted_requests)
    strict_quoted_requests = _minute_dedupe(strict_quoted_requests)
    ack_candidates = _minute_dedupe(ack_candidates)
    reply_candidates = _minute_dedupe(reply_candidates)

    print(f"live_request_count={len(live_requests)} quoted_request_count={len(quoted_requests)} strict_quoted_request_count={len(strict_quoted_requests)} ack_candidate_count={len(ack_candidates)} reply_candidate_count={len(reply_candidates)}")
    print("-" * 100)
    print("LIVE REQUEST CANDIDATES")
    if not live_requests:
        print("  <empty>")
    for idx, (when, email) in enumerate(live_requests, start=1):
        print(f"  {idx}. {_fmt(when)} | from={email.sender_email or email.sender_name} | subj={email.subject}")
    print("-" * 100)
    print("QUOTED REQUEST CANDIDATES")
    if not quoted_requests:
        print("  <empty>")
    for idx, (when, payload) in enumerate(quoted_requests, start=1):
        source_email, from_line, quoted_subj = payload
        print(f"  {idx}. {_fmt(when)} | from_line={from_line}")
        if quoted_subj:
            print(f"     quoted_subj={quoted_subj}")
        print(f"     found_in={source_email.subject}")
    print("-" * 100)
    print("STRICT HEADER QUOTED REQUEST CANDIDATES")
    if not strict_quoted_requests:
        print("  <empty>")
    for idx, (when, payload) in enumerate(strict_quoted_requests, start=1):
        source_email, header_candidate = payload
        print(
            f"  {idx}. {_fmt(when)} | from_line={header_candidate.from_line} | "
            f"raw_sent={header_candidate.raw_sent or '<empty>'}"
        )
        print(
            f"     normalized_sent={header_candidate.normalized_sent or '<empty>'} | "
            f"has_am_pm={header_candidate.has_am_pm} | am_or_pm={header_candidate.am_or_pm or '-'} | "
            f"ambiguous={header_candidate.is_ambiguous}"
        )
        if header_candidate.header_lines:
            print("     header_lines=" + " || ".join(header_candidate.header_lines))
        if header_candidate.quoted_subj:
            print(f"     quoted_subj={header_candidate.quoted_subj}")
        print(f"     found_in={source_email.subject}")
    print("-" * 100)
    print("ACK CANDIDATES")
    if not ack_candidates:
        print("  <empty>")
    for idx, (when, email) in enumerate(ack_candidates, start=1):
        print(f"  {idx}. {_fmt(when)} | from={email.sender_email or email.sender_name} | subj={email.subject}")
    print("-" * 100)
    print("REPLY CANDIDATES")
    if not reply_candidates:
        print("  <empty>")
    for idx, (when, email) in enumerate(reply_candidates, start=1):
        print(f"  {idx}. {_fmt(when)} | from={email.sender_email or email.sender_name} | subj={email.subject}")

    episodes = []
    for ack_ist, ack_email in ack_candidates:
        req_pick = _best_before(live_requests, ack_ist, timedelta(minutes=16))
        req_kind = "live"
        if not req_pick:
            req_pick = _best_before([(when, payload) for when, payload in quoted_requests], ack_ist, timedelta(minutes=16))
            req_kind = "quoted"
        if not req_pick:
            continue
        reply_after = [
            (reply_ist, reply_email)
            for reply_ist, reply_email in reply_candidates
            if reply_ist > ack_ist and (reply_ist - ack_ist) <= timedelta(hours=48)
        ]
        resolved_pick = reply_after[0] if reply_after else (ack_ist, ack_email)
        episodes.append(
            {
                "request": req_pick,
                "request_kind": req_kind,
                "ack": (ack_ist, ack_email),
                "resolved": resolved_pick,
            }
        )

    unique_episodes = []
    seen = set()
    for episode in episodes:
        key = (
            episode["request"][0].replace(second=0, microsecond=0),
            episode["ack"][0].replace(second=0, microsecond=0),
            episode["resolved"][0].replace(second=0, microsecond=0),
        )
        if key in seen:
            continue
        seen.add(key)
        unique_episodes.append(episode)

    strict_episodes = []
    for ack_ist, ack_email in ack_candidates:
        req_pick = _best_before(live_requests, ack_ist, timedelta(minutes=16))
        req_kind = "live"
        if not req_pick:
            req_pick = _best_before([(when, payload) for when, payload in strict_quoted_requests], ack_ist, timedelta(minutes=16))
            req_kind = "quoted-strict"
        if not req_pick:
            continue
        reply_after = [
            (reply_ist, reply_email)
            for reply_ist, reply_email in reply_candidates
            if reply_ist > ack_ist and (reply_ist - ack_ist) <= timedelta(hours=48)
        ]
        resolved_pick = reply_after[0] if reply_after else (ack_ist, ack_email)
        strict_episodes.append(
            {
                "request": req_pick,
                "request_kind": req_kind,
                "ack": (ack_ist, ack_email),
                "resolved": resolved_pick,
            }
        )
    unique_strict_episodes = []
    strict_seen = set()
    for episode in strict_episodes:
        key = (
            episode["request"][0].replace(second=0, microsecond=0),
            episode["ack"][0].replace(second=0, microsecond=0),
            episode["resolved"][0].replace(second=0, microsecond=0),
        )
        if key in strict_seen:
            continue
        strict_seen.add(key)
        unique_strict_episodes.append(episode)

    blue_direct_episodes = []
    for reply_ist, reply_email in reply_candidates:
        req_pick = _best_before(live_requests, reply_ist, timedelta(hours=48))
        req_kind = "live-blue"
        if not req_pick:
            req_pick = _best_before([(when, payload) for when, payload in quoted_requests], reply_ist, timedelta(hours=48))
            req_kind = "quoted-blue"
        if not req_pick:
            continue
        gap = reply_ist - req_pick[0]
        if gap <= timedelta(minutes=16):
            continue
        blue_direct_episodes.append(
            {
                "request": req_pick,
                "request_kind": req_kind,
                "ack": (reply_ist, reply_email),
                "resolved": (reply_ist, reply_email),
            }
        )

    blue_seen = set()
    for episode in blue_direct_episodes:
        key = (
            episode["request"][0].replace(second=0, microsecond=0),
            episode["ack"][0].replace(second=0, microsecond=0),
            episode["resolved"][0].replace(second=0, microsecond=0),
        )
        if key in seen or key in blue_seen:
            continue
        blue_seen.add(key)
        unique_episodes.append(episode)

    print("-" * 100)
    print("SIMULATED EPISODES")
    if not unique_episodes:
        print("  <empty>")
    for idx, episode in enumerate(unique_episodes, start=1):
        req_ist, req_email = episode["request"]
        ack_ist, ack_email = episode["ack"]
        res_ist, res_email = episode["resolved"]
        print(
            f"  episode={idx} | req={_fmt(req_ist)} | ack={_fmt(ack_ist)} | resolved={_fmt(res_ist)} | req_kind={episode['request_kind']}"
        )
        if episode["request_kind"] == "live":
            print(f"    req_from={req_email.sender_email or req_email.sender_name}")
        else:
            source_email, from_line, quoted_subj = req_email
            print(f"    req_from_line={from_line}")
            if quoted_subj:
                print(f"    req_quoted_subj={quoted_subj}")
            print(f"    req_found_in={source_email.subject}")
        print(f"    ack_from={ack_email.sender_email or ack_email.sender_name}")
        print(f"    resolved_from={res_email.sender_email or res_email.sender_name}")

    print("-" * 100)
    print("STRICT HEADER SIMULATED EPISODES")
    if not unique_strict_episodes:
        print("  <empty>")
    for idx, episode in enumerate(unique_strict_episodes, start=1):
        req_ist, req_email = episode["request"]
        ack_ist, ack_email = episode["ack"]
        res_ist, res_email = episode["resolved"]
        print(
            f"  episode={idx} | req={_fmt(req_ist)} | ack={_fmt(ack_ist)} | resolved={_fmt(res_ist)} | req_kind={episode['request_kind']}"
        )
        if episode["request_kind"] == "live":
            print(f"    req_from={req_email.sender_email or req_email.sender_name}")
        else:
            source_email, header_candidate = req_email
            print(f"    req_from_line={header_candidate.from_line}")
            print(
                f"    req_raw_sent={header_candidate.raw_sent or '<empty>'} | "
                f"normalized={header_candidate.normalized_sent or '<empty>'} | "
                f"has_am_pm={header_candidate.has_am_pm} | am_or_pm={header_candidate.am_or_pm or '-'} | "
                f"ambiguous={header_candidate.is_ambiguous}"
            )
            if header_candidate.quoted_subj:
                print(f"    req_quoted_subj={header_candidate.quoted_subj}")
            print(f"    req_found_in={source_email.subject}")
        print(f"    ack_from={ack_email.sender_email or ack_email.sender_name}")
        print(f"    resolved_from={res_email.sender_email or res_email.sender_name}")

    print("-" * 100)
    print("REPLY-ANCHORED PREFERRED EPISODES")
    found_preferred = False
    for row in sorted(family_rows, key=lambda r: int(r.get("_line", 10**9))):
        preferred = _reply_anchored_preferred_episode(
            row,
            live_requests,
            quoted_requests,
            strict_quoted_requests,
            reply_candidates,
        )
        print(
            f"  line={row.get('_line')} | requester={_get_col(row, 'Requester', 'Consultant')} | "
            f"desc={_get_col(row, 'Description')}"
        )
        if not preferred:
            print("    preferred=no reply-anchored episode found")
            continue
        found_preferred = True
        req_ist, req_email = preferred["request"]
        ack_ist, ack_email = preferred["ack"]
        res_ist, res_email = preferred["resolved"]
        print(
            f"    preferred={_fmt(req_ist)} / {_fmt(ack_ist)} / {_fmt(res_ist)} | "
            f"req_kind={preferred['request_kind']}"
        )
        if preferred["request_kind"].startswith("live"):
            print(f"    req_from={req_email.sender_email or req_email.sender_name}")
            print(f"    req_subject={req_email.subject}")
        elif preferred["request_kind"].startswith("quoted-strict"):
            source_email, header_candidate = req_email
            print(f"    req_from_line={header_candidate.from_line}")
            print(
                f"    req_raw_sent={header_candidate.raw_sent or '<empty>'} | "
                f"normalized={header_candidate.normalized_sent or '<empty>'} | "
                f"has_am_pm={header_candidate.has_am_pm} | am_or_pm={header_candidate.am_or_pm or '-'} | "
                f"ambiguous={header_candidate.is_ambiguous}"
            )
            if header_candidate.quoted_subj:
                print(f"    req_quoted_subj={header_candidate.quoted_subj}")
            print(f"    req_found_in={source_email.subject}")
        else:
            source_email, from_line, quoted_subj = req_email
            print(f"    req_from_line={from_line}")
            if quoted_subj:
                print(f"    req_quoted_subj={quoted_subj}")
            print(f"    req_found_in={source_email.subject}")
        print(f"    reply_from={ack_email.sender_email or ack_email.sender_name}")
        print(f"    reply_subject={ack_email.subject}")
    if not found_preferred:
        print("  <empty>")

    print("-" * 100)
    print("ROW SLOT SIMULATION")
    sorted_rows = sorted(family_rows, key=lambda r: int(r.get("_line", 10**9)))
    for slot, row in enumerate(sorted_rows, start=1):
        episode = unique_episodes[slot - 1] if slot - 1 < len(unique_episodes) else None
        print(
            f"  line={row.get('_line')} | requester={_get_col(row, 'Requester', 'Consultant')} | "
            f"current={_get_col(row, 'Created Date & Time') or '-'} / "
            f"{_get_col(row, 'Actual Response Date & Time') or '-'} / "
            f"{_get_col(row, 'Actual Resolved Date & Time') or '-'}"
        )
        if not episode:
            print("    simulated=no unique episode available for this slot")
            continue
        print(
            f"    simulated={_fmt(episode['request'][0])} / {_fmt(episode['ack'][0])} / {_fmt(episode['resolved'][0])}"
        )
        if len(unique_episodes) == 1 and slot > 1:
            print("    note=single real episode only; later duplicate rows should not blindly clone it")


def main() -> None:
    parser = argparse.ArgumentParser(description="Debug request-anchor / duplicate-row families such as EXP006")
    parser.add_argument("--debug-csv", required=True, help="debug_subjects CSV path")
    parser.add_argument("--output-csv", required=True, help="automation_output CSV path")
    parser.add_argument("--eml-dir", required=True, help="Directory containing exported EML files")
    parser.add_argument("--subject", action="append", required=True, help="Description substring to inspect")
    parser.add_argument("--ess-team", default=str(Path("config") / "ess_team.json"), help="ESS team JSON path")
    args = parser.parse_args()

    debug_rows = _match_rows(_load_csv_rows(Path(args.debug_csv)), args.subject)
    output_rows = _match_rows(_load_csv_rows(Path(args.output_csv)), args.subject)
    output_by_line = {int(row["_line"]): row for row in output_rows if row.get("_line")}
    ess_team = _read_json_list(Path(args.ess_team))

    _print_rows("DEBUG ROWS", debug_rows)
    print("-" * 100)
    _print_rows("OUTPUT ROWS", [output_by_line.get(int(row["_line"]), row) for row in debug_rows if row.get("_line")])
    print("-" * 100)

    if not debug_rows:
        print("No matching rows found.")
        return

    all_emails = list(_iter_emails(Path(args.eml_dir)))
    families = _find_family_rows(debug_rows)

    for family_subject, family_rows in families.items():
        requesters = []
        seen_requesters = set()
        for row in family_rows:
            req = _get_col(row, "Requester", "Consultant")
            key = req.lower()
            if req and key not in seen_requesters:
                requesters.append(req)
                seen_requesters.add(key)

        row_tokens = _match_tokens(family_subject)
        row_id_tokens = _id_like_tokens(family_subject)
        matched_emails = [
            email
            for email in all_emails
            if email.sent_time and _row_subject_match(family_subject, row_tokens, row_id_tokens, email.subject)
        ]

        print("=" * 100)
        print(f"family_subject_norm={family_subject}")
        print(f"family_requesters={', '.join(requesters) if requesters else '<none>'}")
        print(f"matched_email_count={len(matched_emails)}")
        print("-" * 100)
        _print_rows("FAMILY ROWS", family_rows)
        print("-" * 100)
        _print_email_block("MATCHED EMAILS", matched_emails, ess_team, requesters)
        print("-" * 100)
        _simulate_family(family_rows, matched_emails, ess_team, requesters)


if __name__ == "__main__":
    main()
