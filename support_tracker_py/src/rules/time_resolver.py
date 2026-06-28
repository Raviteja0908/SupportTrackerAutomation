from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import functools
import html
import os
import random
import re
from zoneinfo import ZoneInfo
import unicodedata
try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None

from src.rules.subject_normalizer import normalize_subject, normalize_subject_for_match

_quoted_header_datetime_cache = {}
_bounded_outlook_header_cache = {}
_canonical_message_lines_cache = {}
_canonical_current_text_cache = {}
_canonical_quoted_header_candidates_cache = {}
_classify_reply_kind_cache = {}


def _email_stable_key(email_record):
    p = getattr(email_record, "path", None)
    if p:
        return ("path", p)
    return (
        "content",
        getattr(email_record, "sender_email", "") or "",
        str(getattr(email_record, "sent_time", "") or ""),
        len(getattr(email_record, "body", "") or ""),
    )


_trace_row_subject_filter = (os.getenv("TRACE_ROW_SUBJECT") or "").strip().lower()


def _trace_focus_row(subject_norm: str | None, phase: str, **fields):
    return


def _fallback_visible_lines_from_text(value: str) -> list[str]:
    if not value:
        return []
    text = re.sub(r"(?is)<style.*?>.*?</style>", " ", value)
    text = re.sub(r"(?is)<script.*?>.*?</script>", " ", text)
    text = re.sub(r"(?i)<\s*br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>", "\n", text)
    text = re.sub(r"(?is)<[^>]+>", " ", text)
    text = html.unescape(text)
    return [ln.strip() for ln in text.splitlines() if ln and ln.strip()]


def _bs4_visible_lines_from_text(value: str) -> list[str]:
    if not value or BeautifulSoup is None:
        return []
    try:
        soup = BeautifulSoup(value, "html.parser")
        for tag in soup.find_all(["br", "p", "div", "blockquote", "td", "tr", "li"]):
            tag.insert_after("\n")
        text = html.unescape(soup.get_text(" "))
    except Exception:
        return []
    return [ln.strip() for ln in text.splitlines() if ln and ln.strip()]


ACK_PHRASES = [
    "sure, we will process the file to",
    "sure we will process the file to",
    "we will check",
    "update you",
    "let you know",
    "we will process the file to",
    "we will do the same",
    "we will do the needful",
    "we have not yet received",
    "we have not received",
    "not yet received",
    "please share",
    "please provide",
    "we will update",
    "we will check and update",
    "we will get back",
    "investigating",
]

NON_ACK_PHRASES = [
    "could you please provide us an update on the below",
    "could you please provide us an update regarding the below",
    "could you please provide an update on the below request",
    "could you please provide an update on the below",
    "could you please provide an update regarding the below request",
    "could you please provide an update regarding the below",
    "could you please provide and update on the below request",
    "could you please provide and update on the below",
    "could you update us on the below",
    "could you update us regarding the below",
    "please provide us an update on the below",
    "please provide us an update",
    "please provide an update on the below request",
    "please provide an update on the below",
    "please provide an update regarding the below request",
    "please provide an update regarding the below",
    "please provide an update",
    "please update us on the below request",
    "please update us on the below",
    "please update us regarding the below request",
    "please update us regarding the below",
    "thank you for the information",
    "thanks for the information",
    "thank you for the update",
    "thanks for the update",
    "thanks for the info",
    "thank you for the confirmation",
    "thanks for the confirmation",
    "noted with thanks",
    "duly noted",
    "please ignore",
    "kindly ignore",
    "ignore the below",
    "please ignore the below",
    "kindly ignore the below",
    "we will ignore",
]

PROMISE_ACK_PHRASES = [
    "we will process",
    "we will proceed",
    "we will proceed to process",
    "we will change",
    "we will deploy",
    "we will work on this",
    "we will look into this",
    "we will investigate",
    "we will verify",
    "we will fix",
    "we will reprocess",
    "we will resend",
    "we will share",
    "we will provide",
    "we will revert",
    "we will take this up",
]

ACK_COURTESY_PREFIXES = [
    "sure",
    "ok",
    "okay",
    "thanks",
    "thank you",
    "noted",
]

DIRECT_RESOLUTION_PHRASES = [
    "successfully processed",
    "processed successfully",
    "file processed successfully",
    "file is processed",
    "file has been processed",
    "file processed",
    "output files got generated",
    "output files generated",
    "files got generated",
    "files generated",
    "output files got uploaded",
    "output files uploaded",
    "files uploaded",
    "successfully uploaded",
    "sent to sap",
    "sent to ecc",
    "successfully sent",
    "processed and uploaded",
    "processed and sent",
]

FILE_ACTION_PHRASES = [
    "adding one more idoc",
    "adding more idoc",
    "adding one more file",
    "adding more files",
]

FORCE_PROD_SAME_TIME_PHRASES = [
    "daily task hyparchive",
    "severes warnings and errors in eai/es aws/es symphony",
    "files from es to grp",
]

INTERNAL_MARKER_RE = re.compile(r"\+{1,}\s*internal\s*\+{1,}", re.IGNORECASE)


def _normalize_name(name: str) -> str:
    if not name:
        return ""
    n = "".join(ch.lower() if ch.isalnum() or ch.isspace() else " " for ch in name)
    return " ".join(n.split())


def _tokenize(name: str):
    tokens = [t for t in _normalize_name(name).split() if t]
    # Add an acronym token when the name contains spaced initials
    # (e.g., "M V S" -> "mvs") to improve matching.
    initials = [t for t in tokens if len(t) == 1]
    if len(initials) >= 2:
        tokens.append("".join(initials))
    return tokens


def _is_force_same_time_subject(subject_norm: str | None, thread=None) -> bool:
    subj_values = []
    if subject_norm:
        subj_values.append((subject_norm or "").lower())
    for e in (thread or []):
        subj = (getattr(e, "subject", "") or "").lower()
        if subj:
            subj_values.append(subj)
    return any(
        phrase in subj
        for subj in subj_values
        for phrase in FORCE_PROD_SAME_TIME_PHRASES
    )


def _email_local(sender_email: str) -> str:
    if not sender_email:
        return ""
    return sender_email.split("@", 1)[0].replace(".", " ").replace("_", " ").replace("-", " ")


def _match_requester(sender_name: str, sender_email: str, requester_name: str) -> bool:
    s = _normalize_name(sender_name)
    r = _normalize_name(requester_name)
    if not r:
        return False

    if r and s and (r in s or s in r):
        return True

    combined = " ".join([s, _normalize_name(_email_local(sender_email))]).strip()
    if r and combined and r in combined:
        return True

    r_tokens = set(_tokenize(requester_name))
    if not r_tokens:
        return False

    s_tokens = set(_tokenize(sender_name)) | set(_tokenize(_email_local(sender_email)))
    if not s_tokens:
        return False

    overlap = r_tokens & s_tokens
    if not overlap:
        return False

    # Stricter thresholds (70% instead of 60%) to reduce false positives
    if len(overlap) / max(1, len(r_tokens)) >= 0.7:
        return True
    if len(overlap) / max(1, len(s_tokens)) >= 0.7:
        return True

    # Last-name strong match fallback
    r_list = list(r_tokens)
    if r_list:
        last = r_list[-1]
        if last in s_tokens and len(overlap) >= 1:
            return True

    return False


def _is_thanks_info_reply(email_record) -> bool:
    """Ignore non-resolution consultant replies like 'Thanks for the information'."""
    body = _leading_body_segment(email_record.body or "").lower()
    if not body:
        return False
    # Courtesy plus an explicit future-action promise is still an ack-like live
    # response, not a pure "thanks/info" mail to be ignored.
    if _is_explicit_ack_signal(body):
        return False
    return _contains_any_phrase(
        body,
        [
            "thanks for the information",
            "thank you for the information",
            "thanks for the confirmation",
            "thank you for the confirmation",
            "thanks for the update",
            "thank you for the update",
            "thanks for the info",
            "noted with thanks",
            "duly noted",
            "we will ignore",
            *NON_ACK_PHRASES,
        ],
    ) or _is_update_chasing_text(body)


def _normalize_for_phrase_match(text: str) -> str:
    if not text:
        return ""
    s = text.lower()
    # Keep apostrophes (contractions) but remove other punctuation to preserve word meaning
    s = re.sub(r"[^a-z0-9']+", " ", s)
    return " ".join(s.split())


def _contains_any_phrase(text: str, phrases: list[str]) -> bool:
    t = _normalize_for_phrase_match(text)
    if not t:
        return False
    for p in phrases:
        pn = _normalize_for_phrase_match(p)
        if pn and pn in t:
            return True
    return False


def _is_update_chasing_text(text: str) -> bool:
    t = _normalize_for_phrase_match(text)
    if not t:
        return False
    if _contains_any_phrase(t, NON_ACK_PHRASES):
        return True
    patterns = [
        r"\bcould you(?: please)? provide(?: us)? (?:an|and)? update(?: on| regarding)?(?: the)? below(?: request)?\b",
        r"\bplease provide(?: us)? (?:an|and)? update(?: on| regarding)?(?: the)? below(?: request)?\b",
        r"\bcould you(?: please)? update us(?: on| regarding)?(?: the)? below(?: request)?\b",
        r"\bplease update us(?: on| regarding)?(?: the)? below(?: request)?\b",
    ]
    return any(re.search(pat, t) for pat in patterns)


def _is_explicit_ack_signal(text: str) -> bool:
    t = _normalize_for_phrase_match(text)
    if not t:
        return False
    if _contains_any_phrase(t, ACK_PHRASES) or _contains_any_phrase(t, PROMISE_ACK_PHRASES):
        return True
    has_courtesy = any(t.startswith(prefix) or f"{prefix} " in t for prefix in ACK_COURTESY_PREFIXES)
    # Forward-action check: look for future-tense indicators not captured by
    # the exact-phrase test above (e.g. "Sure, will do" / "OK, will check").
    has_future_action = bool(re.search(r"\bwill\b|\bshall\b", t))
    return has_courtesy and has_future_action


@functools.lru_cache(maxsize=2000)
def _leading_body_segment(body: str, max_chars: int = 4000) -> str:
    """
    Return only the newest/top body portion, trimming quoted history blocks.
    This avoids false ack/non-ack hits from quoted older conversation text.
    """
    if not body:
        return ""
    lines = body.splitlines()
    cut = len(lines)
    quoted_markers = (
        "-----original message-----",
        "________________________________",
    )
    wrote_re = re.compile(r"^on .+wrote:\s*$", re.IGNORECASE)
    for i, raw in enumerate(lines):
        s = (raw or "").strip()
        sl = s.lower()
        if not s:
            continue
        if any(m in sl for m in quoted_markers):
            cut = i
            break
        if wrote_re.match(s):
            cut = i
            break
        if _is_from_header_line(s) or _is_sent_header_line(s):
            cut = i
            break
    top = "\n".join(lines[:cut]).strip()
    if len(top) > max_chars:
        top = top[:max_chars]
    return top


def _is_ack_body(body: str) -> bool:
    if not body:
        return False
    top = _leading_body_segment(body)
    if _is_update_chasing_text(top):
        return False
    return _is_explicit_ack_signal(top)


def _email_has_explicit_ack_signal(email_record) -> bool:
    body = email_record.body or ""
    if _is_ack_body(body):
        return True
    raw_full = f"{email_record.body or ''}\n{getattr(email_record, 'body_html', '') or ''}"
    if not raw_full:
        return False
    top = _leading_body_segment(raw_full)
    if _is_update_chasing_text(top):
        return False
    fallback = raw_full[:2000]
    return _is_explicit_ack_signal(top) or (
        len(_normalize_for_phrase_match(top)) < 80 and _is_explicit_ack_signal(fallback)
    )


def _email_has_short_ess_ack_signal(email_record) -> bool:
    raw_full = f"{email_record.body or ''}\n{getattr(email_record, 'body_html', '') or ''}"
    if not raw_full:
        return False
    top = _leading_body_segment(raw_full)
    if not top:
        return False
    top_l = top.lower()
    if _contains_any_phrase(top, DIRECT_RESOLUTION_PHRASES):
        return False
    if _contains_any_phrase(top, FILE_ACTION_PHRASES):
        return False
    if _is_update_chasing_text(top):
        return False
    if ("received" in top_l) and ("idoc" in top_l or "file" in top_l or "files" in top_l):
        return False
    lines = [ln.strip() for ln in top.splitlines() if ln and ln.strip()]
    if not lines:
        return False
    content = _normalize_for_phrase_match(" ".join(lines))
    if not content:
        return False
    if len(lines) > 2 or len(content) > 140:
        return False
    strong = (
        "resolved",
        "fixed",
        "completed",
        "success",
        "processed",
        "root cause",
        "issue was",
        "closed",
        "done",
        "sent to sap",
        "sent to ecc",
        "uploaded",
        "generated",
    )
    if any(w in content for w in strong):
        return False
    if any(k in content for k in ("attachment", "attached", "snippet", "snippets", "see attached")):
        return False
    if ("cid:" in raw_full.lower()) or ("<img" in raw_full.lower()):
        return False
    return True


def _is_nonfinal_followup_reply(email_record) -> bool:
    raw_full = f"{email_record.body or ''}\n{getattr(email_record, 'body_html', '') or ''}"
    if not raw_full:
        return False
    top = _leading_body_segment(raw_full)
    if not top:
        return False
    if _is_thanks_info_reply(email_record):
        return True
    return _is_update_chasing_text(top) or _contains_any_phrase(top, NON_ACK_PHRASES)


def _has_direct_resolution_signal(email_record) -> bool:
    raw_full = f"{email_record.body or ''}\n{getattr(email_record, 'body_html', '') or ''}"
    if not raw_full:
        return False
    top = _leading_body_segment(raw_full)
    if not top:
        return False
    if _is_update_chasing_text(top):
        return False
    return _contains_any_phrase(top, DIRECT_RESOLUTION_PHRASES)


def _classify_reply_kind(email_record) -> dict:
    if not email_record:
        return {
            "direct_resolution": False,
            "thanks_info": False,
            "nonfinal_followup": False,
            "explicit_ack": False,
            "short_ess_ack": False,
            "ack_like": False,
            "real_reply": False,
            "kind": "none",
        }
    _ck = _email_stable_key(email_record)
    if _ck in _classify_reply_kind_cache:
        return _classify_reply_kind_cache[_ck]

    direct_resolution = _has_direct_resolution_signal(email_record)
    thanks_info = _is_thanks_info_reply(email_record)
    nonfinal_followup = _is_nonfinal_followup_reply(email_record)
    explicit_ack = _email_has_explicit_ack_signal(email_record)
    short_ess_ack = _email_has_short_ess_ack_signal(email_record)
    ack_like = _is_ack_like_reply(email_record)
    real_reply = False
    if direct_resolution:
        real_reply = True
    elif not (nonfinal_followup or explicit_ack or short_ess_ack or ack_like):
        real_reply = True

    if direct_resolution:
        kind = "direct_resolution"
    elif real_reply:
        kind = "real_reply"
    elif short_ess_ack:
        kind = "short_ess_ack"
    elif explicit_ack or ack_like:
        kind = "ack_like"
    elif thanks_info:
        kind = "thanks_info"
    elif nonfinal_followup:
        kind = "nonfinal_followup"
    else:
        kind = "other"

    result = {
        "direct_resolution": direct_resolution,
        "thanks_info": thanks_info,
        "nonfinal_followup": nonfinal_followup,
        "explicit_ack": explicit_ack,
        "short_ess_ack": short_ess_ack,
        "ack_like": ack_like,
        "real_reply": real_reply,
        "kind": kind,
    }
    _classify_reply_kind_cache[_ck] = result
    return result


def _is_real_reply_candidate(email_record) -> bool:
    return _classify_reply_kind(email_record)["real_reply"]


def _resolution_candidate_rank(email_record, ess_team) -> int:
    if not _is_real_reply_candidate(email_record):
        return -1
    if _has_direct_resolution_signal(email_record) and _is_ess_sender(email_record, ess_team):
        return 4
    if _has_direct_resolution_signal(email_record):
        return 3
    return 2


def _is_consultant_real_reply_candidate(email_record, requester_name: str, ess_team) -> bool:
    if not email_record:
        return False
    if not _match_requester(
        getattr(email_record, "sender_name", ""),
        getattr(email_record, "sender_email", ""),
        requester_name,
    ):
        return False
    return _resolution_candidate_rank(email_record, ess_team) >= 0


def _is_ack_like_reply(email_record) -> bool:
    """
    Ack-like replies should not be treated as resolved-time candidates.
    Keep direct-resolution mails out of this bucket.
    """
    body = _leading_body_segment(email_record.body or "").lower()
    raw_full = f"{email_record.body or ''}\n{getattr(email_record, 'body_html', '') or ''}"
    # Some parsed mails can yield an almost-empty top segment. Use a small
    # bounded fallback window to avoid classifying based on deep quoted history.
    body_fallback = raw_full[:2000].lower() if raw_full else ""
    if not body and not body_fallback:
        return False
    # Update-chasing reminders are non-final and should not drive resolved-time picks.
    if _contains_any_phrase(body, DIRECT_RESOLUTION_PHRASES):
        return False
    # Treat file-add/resend/received updates as non-ack to avoid suppressing
    # ESS-only update replies (e.g., "Adding one more IDOC").
    file_action_hit = _contains_any_phrase(body, FILE_ACTION_PHRASES) or (
        ("received" in body)
        and ("idoc" in body or "file" in body or "files" in body)
    )
    if not file_action_hit and body_fallback:
        file_action_hit = _contains_any_phrase(body_fallback, FILE_ACTION_PHRASES) or (
            ("received" in body_fallback)
            and ("idoc" in body_fallback or "file" in body_fallback or "files" in body_fallback)
        )
    if file_action_hit:
        return False
    if _is_explicit_ack_signal(body) or (
        len(_normalize_for_phrase_match(body)) < 80
        and _is_explicit_ack_signal(body_fallback)
        and not _contains_any_phrase(body_fallback, DIRECT_RESOLUTION_PHRASES)
    ):
        return True
    # Size-based override: long, multi-line replies are more likely to be real updates.
    # Only keep them ack-like if they explicitly contain reminder/non-ack phrases.
    size_text = body if body else body_fallback
    body_lines = [ln for ln in size_text.splitlines() if ln.strip()]
    if (len(body_lines) >= 3 or len(_normalize_for_phrase_match(size_text)) >= 160):
        if not _contains_any_phrase(size_text, NON_ACK_PHRASES):
            return False
    if _contains_any_phrase(body, NON_ACK_PHRASES) or (
        len(_normalize_for_phrase_match(body)) < 40
        and _contains_any_phrase(body_fallback, NON_ACK_PHRASES)
        and not _contains_any_phrase(body_fallback, DIRECT_RESOLUTION_PHRASES)
    ):
        return True
    return _is_ack_body(body)


IST = ZoneInfo("Asia/Kolkata")


def _to_ist(dt: datetime) -> datetime:
    if not isinstance(dt, datetime):
        raise ValueError(f"Expected datetime, got {type(dt).__name__}")
    if dt.year <= 1901:
        raise ValueError(f"Invalid date year {dt.year}")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    return dt.astimezone(IST)


def _has_non_ess_email_near_time(ordered, ess_team, target_dt: datetime, window_hours: int = 48) -> bool:
    if not target_dt:
        return False
    try:
        target = _to_ist(target_dt)
    except Exception:
        return False
    window = timedelta(hours=window_hours)
    for e in ordered:
        if _is_ess_sender(e, ess_team):
            continue
        try:
            sent = _to_ist(e.sent_time)
        except Exception:
            continue
        if abs(sent - target) <= window:
            return True
    return False


def _latest_live_request_before(
    emails,
    ess_team,
    max_dt: datetime | None = None,
    requester_name: str | None = None,
):
    latest_time = None
    latest_src = ""
    for e in emails:
        is_requester_owned = bool(
            requester_name and _match_requester(e.sender_name, e.sender_email, requester_name)
        )
        if _is_ess_sender(e, ess_team) and not is_requester_owned:
            continue
        if max_dt and _to_ist(e.sent_time) > _to_ist(max_dt):
            continue
        if latest_time is None or _to_ist(e.sent_time) > _to_ist(latest_time):
            latest_time = e.sent_time
            latest_src = e.sender_email or e.sender_name
    return latest_time, latest_src


def _first_local_explicit_ack_after(
    emails,
    ess_team,
    after_time: datetime,
    before_time: datetime | None = None,
    requester_name: str | None = None,
    allow_requester_owned: bool = False,
):
    candidates = []
    for e in emails:
        if not getattr(e, "sent_time", None):
            continue
        sent_ist = _to_ist(e.sent_time)
        if sent_ist <= _to_ist(after_time):
            continue
        if before_time and sent_ist > _to_ist(before_time):
            continue
        if not _email_has_explicit_ack_signal(e):
            continue
        is_ess_sender = _is_ess_sender(e, ess_team)
        is_requester_owned = bool(
            allow_requester_owned
            and requester_name
            and _match_requester(e.sender_name, e.sender_email, requester_name)
        )
        if not is_ess_sender and not is_requester_owned:
            continue
        candidates.append((0 if is_ess_sender else 1, sent_ist, e))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _first_local_episode_ack_after(
    emails,
    ess_team,
    after_time: datetime,
    before_time: datetime | None = None,
    requester_name: str | None = None,
    allow_requester_owned: bool = False,
):
    candidates = []
    for e in emails:
        if not getattr(e, "sent_time", None):
            continue
        sent_ist = _to_ist(e.sent_time)
        if sent_ist <= _to_ist(after_time):
            continue
        if before_time and sent_ist > _to_ist(before_time):
            continue
        explicit_ack = _email_has_explicit_ack_signal(e)
        short_ess_ack = _email_has_short_ess_ack_signal(e)
        if not explicit_ack and not short_ess_ack:
            continue
        is_ess_sender = _is_ess_sender(e, ess_team)
        is_requester_owned = bool(
            allow_requester_owned
            and requester_name
            and _match_requester(e.sender_name, e.sender_email, requester_name)
        )
        # Requester-owned replies must stay explicit-only. Short, unknown
        # requester-owned mails are too risky to promote as ack.
        if is_requester_owned and not explicit_ack:
            is_requester_owned = False
        if is_ess_sender:
            candidates.append((0 if explicit_ack else 1, sent_ist, e))
        elif is_requester_owned:
            candidates.append((2, sent_ist, e))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def _format_time(dt: datetime) -> str:
    if not isinstance(dt, datetime):
        return None  # Return None instead of empty string for invalid input
    if dt.year <= 1901:
        return None  # Return None for invalid/sentinel dates
    try:
        dt = _to_ist(dt)
    except Exception as exc:
        return None  # Return None on conversion failure
    return dt.strftime("%d-%m-%Y %H:%M")


def _normalize_header_line(line: str) -> str:
    if not line:
        return ""
    s = line.strip().lower()
    try:
        s = unicodedata.normalize("NFKD", s)
        s = s.encode("ascii", "ignore").decode("ascii")
    except Exception:
        pass
    return s


def _is_from_header_line(line: str) -> bool:
    s = _normalize_header_line(line)
    return s.startswith("from:") or s.startswith("de:") or s.startswith("de :")


def _is_sent_header_line(line: str) -> bool:
    s = _normalize_header_line(line)
    return s.startswith("sent:") or s.startswith("envoye")


def _is_to_header_line(line: str) -> bool:
    s = _normalize_header_line(line)
    return s.startswith("to:") or s.startswith("a:") or s.startswith("a :")


def _is_subject_header_line(line: str) -> bool:
    s = _normalize_header_line(line)
    return s.startswith("subject:") or s.startswith("objet")


@dataclass
class TimeResult:
    created: str
    response: str
    resolved: str


@dataclass
class TimeDebug:
    created_src: str
    ack_src: str
    resolved_src: str
    notes: str


def resolve_times_with_debug(thread, requester_name, ess_team, subject_norm: str | None = None):
    if not thread:
        return (
            TimeResult("", "", ""),
            TimeDebug("", "", "", "No thread found"),
        )

    ordered = sorted(thread, key=lambda e: e.sent_time)
    first = ordered[0]
    ess_emails = [e for e in ordered if _is_ess_sender(e, ess_team)]
    non_ess_emails = [e for e in ordered if not _is_ess_sender(e, ess_team)]
    has_internal_marker = _has_internal_marker(ordered)
    system_notif_emails = [e for e in ordered if _is_system_sender(e)]
    has_ack_phrase = _has_ack_phrase(ordered, ess_team)
    failed_subject = _is_failed_subject(ordered)

    requester_candidates = _requester_email_candidates(requester_name, ess_team)
    requester_emails = [
        e for e in ordered
        if (e.sender_email in requester_candidates)
        or _match_requester(e.sender_name, e.sender_email, requester_name)
    ]

    if _is_force_same_time_subject(subject_norm, ordered):
        t = _format_time(first.sent_time)
        src = first.sender_email or first.sender_name
        return (
            TimeResult(t, t, t),
            TimeDebug(src, src, src, "Force PROD subject; all times same"),
        )

    def _latest_requester_non_ack(msgs):
        for e in reversed(msgs):
            if _is_real_reply_candidate(e):
                return e
        return None

    # Global resolved selector:
    # never use ack-like / thanks-like requester replies as final resolved
    # when a non-ack requester reply exists in the same thread.
    resolved_mail = _latest_requester_non_ack(requester_emails) if requester_emails else None
    requester_first = requester_emails[0] if requester_emails else None
    requester_is_ess = bool(requester_candidates) or any(
        _is_ess_sender(e, ess_team) for e in requester_emails
    )
    def _has_prior_request_anchor(max_dt):
        if not max_dt:
            return False
        req_dt, _req_src = _latest_request_time_before(
            ordered,
            ess_team,
            max_dt=max_dt,
            subject_norm=subject_norm,
        )
        if not req_dt:
            return False
        req_ist = _to_ist(req_dt)
        max_ist = _to_ist(max_dt)
        if req_ist >= max_ist:
            return False
        return (max_ist - req_ist) >= timedelta(minutes=20)
    def _has_local_quoted_request_anchor(email_obj, max_gap_min: int = 45):
        if not email_obj or not getattr(email_obj, "sent_time", None):
            return False
        try:
            local_req = _extract_request_time_from_email(
                email_obj,
                ess_team,
                max_dt=email_obj.sent_time,
                subject_norm=None,
            )
            if not local_req:
                return False
            local_gap = _to_ist(email_obj.sent_time) - _to_ist(local_req)
            return timedelta(minutes=1) <= local_gap <= timedelta(minutes=max_gap_min)
        except Exception:
            return False

    def _system_notification_episode():
        if not system_notif_emails or not requester_name:
            return None

        requester_real = [
            e for e in requester_emails
            if _is_real_reply_candidate(e) and not _is_system_sender(e)
        ]
        if not requester_real:
            return None

        def _has_stronger_live_before(cap_dt):
            for e in ordered:
                if not getattr(e, "sent_time", None):
                    continue
                e_ist = _to_ist(e.sent_time)
                if e_ist >= _to_ist(cap_dt):
                    break
                if _is_ess_sender(e, ess_team) or _is_system_sender(e):
                    continue
                if _match_requester(e.sender_name, e.sender_email, requester_name):
                    continue
                return True
            return False

        def _has_stronger_quoted_before(cap_dt):
            for e in ordered:
                parsed = _extract_request_time_from_email(
                    e,
                    ess_team,
                    max_dt=cap_dt,
                    subject_norm=subject_norm,
                )
                if parsed and _to_ist(parsed) < _to_ist(cap_dt):
                    return True
            return False

        for sys_msg in reversed(system_notif_emails):
            sys_dt = sys_msg.sent_time
            ack_mail = _first_local_episode_ack_after(
                ordered,
                ess_team,
                sys_dt,
                requester_name=requester_name,
                allow_requester_owned=False,
            )
            if ack_mail:
                ack_dt = ack_mail.sent_time
                if _has_stronger_live_before(ack_dt) or _has_stronger_quoted_before(ack_dt):
                    continue
                requester_after_ack = [
                    e for e in requester_real
                    if _to_ist(e.sent_time) > _to_ist(ack_dt)
                ]
                if not requester_after_ack:
                    continue
                resolved_pick = min(requester_after_ack, key=lambda e: _to_ist(e.sent_time))
                return (
                    TimeResult(
                        _format_time(sys_dt),
                        _format_time(ack_dt),
                        _format_time(resolved_pick.sent_time),
                    ),
                    TimeDebug(
                        sys_msg.sender_email or sys_msg.sender_name,
                        ack_mail.sender_email or ack_mail.sender_name,
                        resolved_pick.sender_email or resolved_pick.sender_name,
                        "SystemNotificationEpisode[with-ack]",
                    ),
                )

            requester_after_system = [
                e for e in requester_real
                if _to_ist(e.sent_time) > _to_ist(sys_dt)
            ]
            if not requester_after_system:
                continue
            reply_pick = max(requester_after_system, key=lambda e: _to_ist(e.sent_time))
            if _has_stronger_live_before(reply_pick.sent_time) or _has_stronger_quoted_before(reply_pick.sent_time):
                continue
            t = _format_time(reply_pick.sent_time)
            src = reply_pick.sender_email or reply_pick.sender_name
            return (
                TimeResult(t, t, t),
                TimeDebug(src, "ACK NOT FOUND", src, "SystemNotificationEpisode[no-ack-all-same]"),
            )

        return None

    system_notification_episode = _system_notification_episode()
    if system_notification_episode:
        return system_notification_episode

    # Global requester follow-up guard (non-hardcoded):
    # If requester posts a newer non-ack follow-up over requester/ESS mail and there is no
    # non-ESS request in-between, anchor all three times to that latest requester reply.
    # This covers ESS-initiated and running chains with stale earlier request anchors.
    if requester_emails:
        requester_non_ack = [e for e in requester_emails if _is_real_reply_candidate(e)]
        if len(requester_non_ack) >= 2:
            latest_requester = max(requester_non_ack, key=lambda e: _to_ist(e.sent_time))
            prev_candidates = [
                e for e in requester_non_ack
                if _to_ist(e.sent_time) < _to_ist(latest_requester.sent_time)
            ]
            prev_requester = max(prev_candidates, key=lambda e: _to_ist(e.sent_time)) if prev_candidates else None

            if prev_requester and (_to_ist(latest_requester.sent_time) - _to_ist(prev_requester.sent_time)) >= timedelta(minutes=30):
                window_start = _to_ist(prev_requester.sent_time)
                window_end = _to_ist(latest_requester.sent_time)

                # Live non-ESS request between previous requester and latest requester blocks this guard.
                live_non_ess_between = any(
                    (not _is_ess_sender(e, ess_team))
                    and (not _is_system_sender(e))
                    and e.sent_time
                    and (window_start < _to_ist(e.sent_time) <= window_end)
                    for e in ordered
                )

                # Also block if parsed quoted non-ESS request exists in the same window.
                parsed_req_between = False
                latest_req, latest_req_src = _latest_request_time_before(
                    ordered,
                    ess_team,
                    max_dt=latest_requester.sent_time,
                    subject_norm=subject_norm,
                )
                if latest_req and isinstance(latest_req_src, str) and latest_req_src.startswith("PARSED_"):
                    req_dt = _to_ist(latest_req)
                    if window_start < req_dt < window_end:
                        parsed_req_between = True

                # If any ESS teammate already acknowledged the earlier requester mail
                # inside this requester-to-requester span, do not collapse the whole
                # episode to the latest requester follow-up. Let the normal
                # request->ack->resolved flow preserve that in-between ESS ack.
                local_episode_ack_between = _first_local_episode_ack_after(
                    ordered,
                    ess_team,
                    after_time=prev_requester.sent_time,
                    before_time=latest_requester.sent_time,
                    requester_name=requester_name,
                    allow_requester_owned=False,
                )
                requester_hybrid_ack_between = (
                    _has_local_quoted_request_anchor(prev_requester)
                    or _has_local_quoted_request_anchor(latest_requester)
                )

                if (
                    not live_non_ess_between
                    and not parsed_req_between
                    and not local_episode_ack_between
                    and not requester_hybrid_ack_between
                ):
                    t = _format_time(latest_requester.sent_time)
                    src = latest_requester.sender_email or latest_requester.sender_name
                    return (
                        TimeResult(t, t, t),
                        TimeDebug(src, src, src, "Requester follow-up (no in-between request)"),
                    )
        # Single-top continuation fallback:
        # ESS-initiated chains can have only one visible requester non-ack on top
        # (previous request is inside quoted/older ESS context). In that case,
        # if top mail is requester non-ack and the mail right below is ESS, use
        # top requester time for all three.
        elif len(requester_non_ack) == 1 and _is_ess_sender(first, ess_team):
            latest_top = ordered[-1]
            if (
                _match_requester(latest_top.sender_name, latest_top.sender_email, requester_name)
                and _is_real_reply_candidate(latest_top)
            ):
                prev_msg = None
                for e in reversed(ordered[:-1]):
                    if getattr(e, "sent_time", None):
                        prev_msg = e
                        break
                if (
                    prev_msg
                    and _is_ess_sender(prev_msg, ess_team)
                    and not _has_prior_request_anchor(latest_top.sent_time)
                    and not _has_local_quoted_request_anchor(latest_top)
                ):
                    t = _format_time(latest_top.sent_time)
                    src = latest_top.sender_email or latest_top.sender_name
                    return (
                        TimeResult(t, t, t),
                        TimeDebug(src, src, src, "Requester follow-up(top-only)"),
                    )

    # If no ESS reply and no requester reply, mark all three same
    if not ess_emails and not requester_emails:
        t = _format_time(first.sent_time)
        return (
            TimeResult(t, t, t),
            TimeDebug(first.sender_email, first.sender_email, first.sender_email, "No ESS or requester replies"),
        )

    # ESS-only thread: use quoted request only if it's recent.
    # If no non-ESS request exists, avoid forcing all-three-same when multiple ESS emails exist.
    if not non_ess_emails:
        # Early ESS-only continuation rule (global, non-hardcoded):
        # If the latest visible mail is from requester (non-ack/non-thanks) and is
        # directly over an ESS mail (own/teammate), treat this as continuation and
        # anchor all three timestamps to that latest requester reply.
        if len(ess_emails) >= 2 and requester_emails:
            latest_msg = ordered[-1]
            if _match_requester(latest_msg.sender_name, latest_msg.sender_email, requester_name):
                if (
                    _is_real_reply_candidate(latest_msg)
                ):
                    prev_ess = None
                    for e in reversed(ordered[:-1]):
                        if not getattr(e, "sent_time", None):
                            continue
                        if _is_ess_sender(e, ess_team):
                            prev_ess = e
                            break
                    if (
                        prev_ess
                        and not _has_prior_request_anchor(latest_msg.sent_time)
                        and not _has_local_quoted_request_anchor(latest_msg)
                    ):
                        t = _format_time(latest_msg.sent_time)
                        who = latest_msg.sender_email or latest_msg.sender_name
                        return (
                            TimeResult(t, t, t),
                            TimeDebug(who, who, who, "ESS-only continuation(top requester)"),
                        )

        scan_max_dt = resolved_mail.sent_time if resolved_mail and resolved_mail.sent_time else ordered[-1].sent_time
        parsed_any = _find_latest_quoted_request_time(
            ordered,
            ess_team,
            max_dt=scan_max_dt,
            subject_norm=subject_norm,
        )
        if parsed_any:
            # Allow quoted request only if it's within 7 days of first ESS mail
            if (_to_ist(first.sent_time) - _to_ist(parsed_any)) <= timedelta(days=7):
                # Let normal flow use parsed request; do not force all-same
                pass
            else:
                parsed_any = None
        # ESS-only fallback: if no subject-matched quoted request found, try a relaxed
        # parse only when the body contains the normalized subject (safe guard).
        if not parsed_any and subject_norm:
            subj_alt = normalize_subject_for_match(subject_norm)
            relaxed_latest = None
            if subj_alt:
                for e in ordered:
                    body = (e.body or "")
                    body_html = (getattr(e, "body_html", "") or "")
                    haystack = f"{body}\n{body_html}".strip()
                    if not haystack:
                        continue
                    if subj_alt.lower() not in haystack.lower():
                        continue
                    parsed = _extract_request_time_from_email(
                        e,
                        ess_team,
                        max_dt=scan_max_dt,
                        subject_norm=None,
                    )
                    if not parsed:
                        continue
                    if (_to_ist(first.sent_time) - _to_ist(parsed)) > timedelta(days=7):
                        continue
                    if relaxed_latest is None or _to_ist(parsed) > _to_ist(relaxed_latest):
                        relaxed_latest = parsed
            if relaxed_latest:
                parsed_any = relaxed_latest
        if not parsed_any:
            if len(ess_emails) >= 2:
                # ESS-only continuation guard:
                # When the latest requester mail is a real (non-ack/non-thanks) update
                # and it is on top of an ESS continuation (own/teammate), keep all three
                # timestamps on that latest requester reply.
                requester_non_ack_tail = [e for e in requester_emails if _is_real_reply_candidate(e)]
                if requester_non_ack_tail:
                    latest_req_tail = max(requester_non_ack_tail, key=lambda e: e.sent_time)
                    latest_ess = max(ess_emails, key=lambda e: e.sent_time)
                    if latest_req_tail is latest_ess:
                        prev_ess = None
                        for e in reversed(ordered):
                            if not getattr(e, "sent_time", None):
                                continue
                            if _to_ist(e.sent_time) >= _to_ist(latest_req_tail.sent_time):
                                continue
                            if _is_ess_sender(e, ess_team):
                                prev_ess = e
                                break
                        if (
                            prev_ess
                            and not _has_prior_request_anchor(latest_req_tail.sent_time)
                            and not _has_local_quoted_request_anchor(latest_req_tail)
                        ):
                            t = _format_time(latest_req_tail.sent_time)
                            who = latest_req_tail.sender_email or latest_req_tail.sender_name
                            return (
                                TimeResult(t, t, t),
                                TimeDebug(
                                    who,
                                    who,
                                    who,
                                    "ESS-only continuation; latest requester",
                                ),
                            )

                created_mail_pick = ess_emails[0]
                if requester_emails:
                    try:
                        requester_non_ack_initial = [
                            e for e in requester_emails if _is_real_reply_candidate(e)
                        ]
                        requester_first = min(requester_emails, key=lambda e: e.sent_time)
                        requester_first_non_ack = (
                            min(requester_non_ack_initial, key=lambda e: e.sent_time)
                            if requester_non_ack_initial
                            else None
                        )
                        first_is_requester = _match_requester(
                            created_mail_pick.sender_name,
                            created_mail_pick.sender_email,
                            requester_name,
                        )
                        anchor_candidate = requester_first_non_ack or requester_first
                        gap = _to_ist(anchor_candidate.sent_time) - _to_ist(created_mail_pick.sent_time)
                        # Keep current behavior by default, but when the first ESS mail
                        # is by someone else and requester starts much later, anchor created
                        # to requester to avoid stale carry-over from older sibling chains.
                        #
                        # Safety: do not jump across large multi-day gaps on requester
                        # ack-like tails (common in long ESS-only chains).
                        if first_is_requester:
                            created_mail_pick = anchor_candidate
                        elif requester_first_non_ack and timedelta(hours=12) <= gap <= timedelta(days=3):
                            created_mail_pick = requester_first_non_ack
                    except Exception:
                        pass
                created_t = _format_time(created_mail_pick.sent_time)
                non_ack_ess = [e for e in ess_emails if _is_real_reply_candidate(e)]
                last_ess = max(non_ack_ess, key=lambda e: e.sent_time) if non_ack_ess else max(ess_emails, key=lambda e: e.sent_time)

                # Prefer the strongest real consultant reply in ESS-only
                # requester-span threads. Do not let another ESS teammate win
                # resolved when the row's consultant/requester lane should own it.
                resolution_candidates = []
                for e in requester_emails:
                    if not getattr(e, "sent_time", None):
                        continue
                    if _to_ist(e.sent_time) < _to_ist(created_mail_pick.sent_time):
                        continue
                    if not _is_consultant_real_reply_candidate(e, requester_name, ess_team):
                        continue
                    rank = _resolution_candidate_rank(e, ess_team)
                    resolution_candidates.append((rank, _to_ist(e.sent_time), e))
                if resolution_candidates:
                    resolution_candidates.sort(key=lambda item: (item[0], item[1]))
                    resolved_mail_pick = resolution_candidates[-1][2]
                    span_note = "ESS-only; no non-ESS request; requester span"
                elif requester_emails:
                    # All requester replies are ack-like/update-like; do not use them
                    # as resolved. Fall back to latest non-ack ESS reply.
                    resolved_mail_pick = last_ess
                    # Keep legacy marker for downstream guards.
                    span_note = "ESS-only; no non-ESS request; requester span(ack-like); requester span(all-ack->ess)"
                    # Safety: in all-ack requester tails, never promote requester
                    # update-chasing mails to resolved. Keep resolved on latest
                    # non-ack ESS reply.

                else:
                    resolved_mail_pick = last_ess
                    span_note = "ESS-only; no non-ESS request; span"

                resolved_t = _format_time(resolved_mail_pick.sent_time) if resolved_mail_pick else created_t

                # Ack selection for ESS-only span:
                # 1) prefer earliest ack-phrase mail between created and resolved
                # 2) fallback to earliest requester reply in that window
                # 3) fallback to earliest ESS reply in that window
                # 4) final fallback keeps historical behavior (ack=resolved)
                ack_mail_pick = None
                created_ist = _to_ist(created_mail_pick.sent_time)
                resolved_ist = _to_ist(resolved_mail_pick.sent_time) if resolved_mail_pick else None

                ack_candidates = [
                    e for e in ess_emails
                    if e.sent_time and _to_ist(e.sent_time) > created_ist
                    and (resolved_ist is None or _to_ist(e.sent_time) <= resolved_ist)
                ]
                ack_candidates.sort(key=lambda e: e.sent_time)
                dropped_stale_ack = False

                for e in ack_candidates:
                    body = (e.body or "").lower()
                    if _is_ack_body(body):
                        ack_mail_pick = e
                        break

                if not ack_mail_pick and requester_emails:
                    requester_after = [
                        e for e in requester_emails
                        if e.sent_time and _to_ist(e.sent_time) > created_ist
                        and (resolved_ist is None or _to_ist(e.sent_time) <= resolved_ist)
                    ]
                    requester_after.sort(key=lambda e: e.sent_time)
                    if requester_after:
                        ack_mail_pick = requester_after[0]

                if not ack_mail_pick and ack_candidates:
                    ack_mail_pick = ack_candidates[0]

                # Safe window guard: avoid using very-late "ack" in ESS-only span rows.
                # If selected ack is stale, try near-created fallback candidates first.
                max_ack_delay = timedelta(hours=12)
                if ack_mail_pick and (_to_ist(ack_mail_pick.sent_time) - created_ist) > max_ack_delay:
                    dropped_stale_ack = True
                    requester_near = [
                        e for e in requester_emails
                        if e.sent_time
                        and _to_ist(e.sent_time) > created_ist
                        and _to_ist(e.sent_time) <= (created_ist + max_ack_delay)
                    ]
                    requester_near.sort(key=lambda e: e.sent_time)
                    if requester_near:
                        ack_mail_pick = requester_near[0]
                    else:
                        ess_near = [
                            e for e in ess_emails
                            if e.sent_time
                            and _to_ist(e.sent_time) > created_ist
                            and _to_ist(e.sent_time) <= (created_ist + max_ack_delay)
                        ]
                        ess_near.sort(key=lambda e: e.sent_time)
                        ack_mail_pick = ess_near[0] if ess_near else None

                if ack_mail_pick:
                    ack_t = _format_time(ack_mail_pick.sent_time)
                    ack_src = ack_mail_pick.sender_email or ack_mail_pick.sender_name
                else:
                    ack_t = created_t if dropped_stale_ack else resolved_t
                    ack_src = "ACK NOT FOUND"
                    if dropped_stale_ack:
                        span_note = f"{span_note}; AckWindowGuard"
                return (
                    TimeResult(created_t, ack_t, resolved_t),
                    TimeDebug(
                        created_mail_pick.sender_email or created_mail_pick.sender_name,
                        ack_src,
                        (resolved_mail_pick.sender_email or resolved_mail_pick.sender_name) if resolved_mail_pick else (ess_emails[0].sender_email or ess_emails[0].sender_name),
                        span_note,
                    ),
                )
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(first.sender_email, first.sender_email, first.sender_email, "ESS-only; no non-ESS request"),
            )

    # ESS-initiated + system notification + no ack => all three same (use first ESS mail)
    if _is_ess_sender(first, ess_team) and system_notif_emails:
        ack_exists = _has_ack_phrase(ordered, ess_team)
        if not ack_exists:
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(first.sender_email, first.sender_email, first.sender_email, "ESS init + system notif; no ack"),
            )

    # If requester is ESS and system notification exists with no ack, force all same to requester first mail
    if requester_is_ess and system_notif_emails:
        ack_exists = _has_ack_phrase(ordered, ess_team)
        if not ack_exists and requester_first:
            t = _format_time(requester_first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(
                    requester_first.sender_email,
                    "ACK NOT FOUND",
                    requester_first.sender_email,
                    "Requester ESS + system notif; no ack",
                ),
            )

    # ESS-only thread: use all same only if single ESS email and no quoted request anywhere
    if not non_ess_emails and ess_emails and len(ess_emails) == 1 and not has_ack_phrase and not has_internal_marker:
        only = ess_emails[0]
        parsed = _extract_request_time_from_email(
            only,
            ess_team,
            max_dt=only.sent_time,
            subject_norm=subject_norm,
        )
        if not parsed:
            t = _format_time(only.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(only.sender_email, only.sender_email, only.sender_email, "Single ESS; no request found"),
            )

    # ESS-initiated + no ack + no consultant reply after request => all three same (use first ESS time)
    if _is_ess_sender(first, ess_team) and not has_ack_phrase:
        latest_req, _latest_src = _latest_request_time_before(
            ordered,
            ess_team,
            subject_norm=subject_norm,
        )
        if latest_req and resolved_mail and _to_ist(resolved_mail.sent_time) < _to_ist(latest_req):
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(
                    first.sender_email,
                    first.sender_email,
                    first.sender_email,
                    "ESS initiated; no ack; no consultant reply after request",
                ),
            )

    # Failed/skip subjects: if consultant initiated and never replied after the request, keep all three same
    if failed_subject and _match_requester(first.sender_name, first.sender_email, requester_name):
        latest_req, _latest_src = _latest_request_time_before(
            ordered,
            ess_team,
            subject_norm=subject_norm,
        )
        if latest_req and (resolved_mail is None or _to_ist(resolved_mail.sent_time) < _to_ist(latest_req)):
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(
                    first.sender_email,
                    first.sender_email,
                    first.sender_email,
                    "Failed subject; consultant initiated; no consultant reply after request",
                ),
            )

    # Failed/skip subjects with no ack phrase: handle safely without treating late ESS replies as ack
    if failed_subject and not has_ack_phrase:
        req_time, req_src = _latest_request_time_before(
            ordered,
            ess_team,
            subject_norm=subject_norm,
        )
        # If ESS initiated and no non-ESS request time found, force all three same
        if _is_ess_sender(first, ess_team) and not req_time:
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(first.sender_email, first.sender_email, first.sender_email, "Failed subject; ESS initiated; no ack phrase"),
            )
        if req_time:
            explicit_ack = _first_local_explicit_ack_after(
                ordered,
                ess_team,
                after_time=req_time,
                before_time=resolved_mail.sent_time if resolved_mail else None,
                requester_name=requester_name,
                allow_requester_owned=True,
            )
            if explicit_ack:
                resolved_pick = (
                    resolved_mail
                    if resolved_mail and _to_ist(resolved_mail.sent_time) >= _to_ist(explicit_ack.sent_time)
                    else explicit_ack
                )
                return (
                    TimeResult(
                        _format_time(req_time),
                        _format_time(explicit_ack.sent_time),
                        _format_time(resolved_pick.sent_time),
                    ),
                    TimeDebug(
                        req_src,
                        explicit_ack.sender_email or explicit_ack.sender_name,
                        resolved_pick.sender_email or resolved_pick.sender_name,
                        "Failed subject; explicit local ack fallback",
                    ),
                )
            episode_ack = _first_local_episode_ack_after(
                ordered,
                ess_team,
                after_time=req_time,
                before_time=resolved_mail.sent_time if resolved_mail else None,
                requester_name=requester_name,
                allow_requester_owned=True,
            )
            if episode_ack:
                resolved_pick = (
                    resolved_mail
                    if resolved_mail and _to_ist(resolved_mail.sent_time) >= _to_ist(episode_ack.sent_time)
                    else episode_ack
                )
                return (
                    TimeResult(
                        _format_time(req_time),
                        _format_time(episode_ack.sent_time),
                        _format_time(resolved_pick.sent_time),
                    ),
                    TimeDebug(
                        req_src,
                        episode_ack.sender_email or episode_ack.sender_name,
                        resolved_pick.sender_email or resolved_pick.sender_name,
                        "Failed subject; local ack fallback",
                    ),
                )
        # If a direct-resolution ESS reply exists, use it as response/resolved
        direct_reply = _find_direct_resolution_reply(
            ordered,
            ess_team,
            after_time=req_time,
            requester_name=requester_name,
        )
        if req_time and direct_reply:
            created_time = req_time
            created_src = req_src
            response_time = _format_time(direct_reply.sent_time)
            notes = "Direct resolution (failed subject; no ack phrase)"
            return (
                TimeResult(_format_time(created_time), response_time, response_time),
                TimeDebug(created_src, direct_reply.sender_email or "", direct_reply.sender_email or "", notes),
            )
        # No direct-resolution phrase match: if requester has a real (non-ack)
        # follow-up after the request, use that episode before collapsing all-three.
        if req_time:
            req_after = [
                e for e in requester_emails
                if getattr(e, "sent_time", None)
                and _to_ist(e.sent_time) >= _to_ist(req_time)
                and _is_real_reply_candidate(e)
            ]
            req_after.sort(key=lambda e: e.sent_time)
            if req_after:
                resp_pick = req_after[0]
                resolved_candidates = [
                    e for e in req_after
                    if _to_ist(e.sent_time) <= (_to_ist(resp_pick.sent_time) + timedelta(hours=72))
                ]
                res_pick = resolved_candidates[-1] if resolved_candidates else resp_pick
                return (
                    TimeResult(
                        _format_time(req_time),
                        _format_time(resp_pick.sent_time),
                        _format_time(res_pick.sent_time),
                    ),
                    TimeDebug(
                        req_src,
                        resp_pick.sender_email or resp_pick.sender_name,
                        res_pick.sender_email or res_pick.sender_name,
                        "Failed subject; no ack phrase; requester fallback",
                    ),
                )
            t = _format_time(req_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(req_src, "ACK NOT FOUND", req_src, "Failed subject; no ack phrase"),
            )

    # Ack: first ESS reply with ack phrase (requester not required)
    ack_mail = None
    ack_fallback = None
    non_ess_emails = [e for e in ordered if not _is_ess_sender(e, ess_team)]
    seen_non_ess = False
    for e in ordered:
        if not _is_ess_sender(e, ess_team):
            seen_non_ess = True
            continue
        if _is_ess_sender(e, ess_team):
            body = (e.body or "").lower()
            if _is_ack_body(body):
                ack_mail = e
                break
            if seen_non_ess:
                ack_fallback = e

    created_mail = None
    created_time = None
    created_src = ""
    ack_src = "ACK NOT FOUND"
    response_time = "NA"
    notes = "Ack missing"

    # Preferred path: use requester reply -> latest request before it -> first ESS reply after request
    if resolved_mail:
        req_time, req_src = _latest_request_time_before(
            ordered,
            ess_team,
            resolved_mail.sent_time,
            subject_norm=subject_norm,
        )
        if req_time:
            created_time = req_time
            created_src = req_src
            notes = "Created from requester->ack chain"
            ack_candidate = _first_ess_after(ordered, ess_team, req_time, resolved_mail.sent_time)
            if ack_candidate:
                effective_ack = ack_candidate
                ack_src = ack_candidate.sender_email
            else:
                effective_ack = None
        else:
            effective_ack = None
    else:
        effective_ack = None

    if effective_ack is None:
        # Try to pick ack based on closest ESS reply after latest request
        req_time, req_src = _latest_request_time_before(
            ordered,
            ess_team,
            subject_norm=subject_norm,
        )
        if req_time:
            ack_candidate = _first_ess_after(ordered, ess_team, req_time, None)
            if ack_candidate:
                effective_ack = ack_candidate
                ack_src = ack_candidate.sender_email
                if not created_time:
                    created_time = req_time
                    created_src = req_src
            else:
                effective_ack = ack_mail or ack_fallback
        else:
            effective_ack = ack_mail or ack_fallback

    # ESS-only thread where requester exists: use requester first mail as created,
    # and pick first ESS reply after it as ack (even if no non-ESS request exists).
    if not non_ess_emails and requester_first and not created_time:
        created_time = requester_first.sent_time
        created_src = requester_first.sender_email or requester_first.sender_name
        notes = "ESS-only; created from requester"
        if effective_ack is None:
            for e in ordered:
                if not _is_ess_sender(e, ess_team):
                    continue
                if _to_ist(e.sent_time) <= _to_ist(created_time):
                    continue
                effective_ack = e
                ack_src = e.sender_email or e.sender_name
                notes = "ESS-only; ack from ESS after requester"
                break

    # Chain alignment: ensure ack is AFTER the latest non-ESS request
    if effective_ack:
        latest_req, latest_req_src = _latest_request_time_before(
            ordered,
            ess_team,
            subject_norm=subject_norm,
        )
        ack_dt = _to_ist(effective_ack.sent_time)
        if latest_req and _to_ist(latest_req) > ack_dt:
            # Ack is before the latest request; find first ESS reply after that request
            new_ack = _first_ess_after(
                ordered,
                ess_team,
                latest_req,
                resolved_mail.sent_time if resolved_mail else None,
            )
            if new_ack:
                effective_ack = new_ack
                ack_src = new_ack.sender_email or new_ack.sender_name
                created_time = latest_req
                created_src = latest_req_src
                notes = "Ack realigned after latest request"
            else:
                # No ESS reply after the latest request
                effective_ack = None
                if not created_time:
                    created_time = latest_req
                    created_src = latest_req_src
                notes = "Ack missing (no ESS reply after latest request)"
        else:
            # Anchor created to the latest request BEFORE ack
            req_before_ack, req_before_ack_src = _latest_request_time_before(
                ordered,
                ess_team,
                effective_ack.sent_time,
                subject_norm=subject_norm,
            )
            if req_before_ack:
                if created_time is None or _to_ist(created_time) > ack_dt or _to_ist(created_time) < _to_ist(req_before_ack):
                    created_time = req_before_ack
                    created_src = req_before_ack_src
                    notes = "Created anchored to request before ack"

    # Prefer earliest ack phrase if it occurs before current ack and has a request before it
    if effective_ack or ack_mail:
        preferred_ack = None
        preferred_req = None
        preferred_src = ""
        for e in ordered:
            if not _is_ess_sender(e, ess_team):
                continue
            body = (e.body or "").lower()
            if not _is_ack_body(body):
                continue
            if created_time and _to_ist(e.sent_time) < _to_ist(created_time):
                continue
            req_time, req_src = _latest_request_time_before(
                ordered,
                ess_team,
                e.sent_time,
                subject_norm=subject_norm,
            )
            if not req_time:
                continue
            preferred_ack = e
            preferred_req = req_time
            preferred_src = req_src
            break
        if preferred_ack:
            if effective_ack is None or _to_ist(preferred_ack.sent_time) < _to_ist(effective_ack.sent_time):
                effective_ack = preferred_ack
                ack_src = preferred_ack.sender_email
                created_time = preferred_req
                created_src = preferred_src
                notes = "Ack phrase preferred (earlier)"

    # Subject-based shortcut: failed/process errors with no ack => all three same based on requester first mail
    if effective_ack is None and failed_subject:
        if requester_first:
            t = _format_time(requester_first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(requester_first.sender_email, "ACK NOT FOUND", requester_first.sender_email, "Failed subject; no ack"),
            )
        if _is_ess_sender(first, ess_team):
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(first.sender_email, "ACK NOT FOUND", first.sender_email, "Failed subject; no ack (ESS first)"),
            )

    if effective_ack and not created_time:
        if ack_mail is None:
            notes = "Ack fallback (first ESS reply)"
        ack_src = effective_ack.sender_email
        req_time, req_src = _latest_request_time_before(
            ordered,
            ess_team,
            effective_ack.sent_time,
            subject_norm=subject_norm,
        )
        if req_time:
            created_time = req_time
            created_src = req_src
        else:
            parsed = _extract_request_time_from_email(
                effective_ack,
                ess_team,
                max_dt=effective_ack.sent_time,
                subject_norm=subject_norm,
            )
            if parsed and not _has_non_ess_email_near_time(ordered, ess_team, parsed):
                parsed = None
            if parsed:
                created_time = parsed
                created_src = "PARSED_FROM_BODY"
                if ack_mail is None:
                    notes = "Ack fallback; created parsed from body"
                else:
                    notes = "Ack found; created parsed from body"
            else:
                parsed_any = _find_latest_quoted_request_time(
                    ordered,
                    ess_team,
                    max_dt=effective_ack.sent_time,
                    subject_norm=subject_norm,
                )
                if parsed_any and not _has_non_ess_email_near_time(ordered, ess_team, parsed_any):
                    parsed_any = None
                if parsed_any:
                    created_time = parsed_any
                    created_src = "PARSED_FROM_BODY_ANY"
                    if ack_mail is None:
                        notes = "Ack fallback; created parsed from any body"
                    else:
                        notes = "Ack found; created parsed from any body"

    # Prefer request time parsed from the ack body if it is closer to the ack
    if effective_ack:
        parsed_from_ack = _extract_request_time_from_email(
            effective_ack,
            ess_team,
            max_dt=effective_ack.sent_time,
            subject_norm=subject_norm,
        )
        if parsed_from_ack and not _has_non_ess_email_near_time(ordered, ess_team, parsed_from_ack):
            parsed_from_ack = None
        if parsed_from_ack:
            if not created_time:
                created_time = parsed_from_ack
                created_src = "PARSED_FROM_ACK_BODY"
                notes = "Created from ack body"
            else:
                try:
                    ack_dt = _to_ist(effective_ack.sent_time)
                    if _to_ist(parsed_from_ack) <= ack_dt and _to_ist(created_time) <= ack_dt:
                        if (ack_dt - _to_ist(parsed_from_ack)) < (ack_dt - _to_ist(created_time)):
                            created_time = parsed_from_ack
                            created_src = "PARSED_FROM_ACK_BODY"
                            notes = "Created from ack body (closer)"
                except Exception:
                    pass
        # If the current created anchor still comes from a broad body parse and is
        # much older than the ack, try the closest quoted request visible in the
        # same ack email before falling back to older history.
        try:
            if created_time and isinstance(created_src, str) and created_src.startswith("PARSED_FROM_"):
                ack_dt = _to_ist(effective_ack.sent_time)
                created_dt = _to_ist(created_time)
                if (ack_dt - created_dt) > timedelta(minutes=45):
                    closest_from_ack = _extract_request_time_from_email_closest(
                        effective_ack,
                        ess_team,
                        anchor_dt=effective_ack.sent_time,
                        max_dt=effective_ack.sent_time,
                        subject_norm=subject_norm,
                    )
                    if not closest_from_ack:
                        closest_from_ack = _extract_request_time_from_email_closest(
                            effective_ack,
                            ess_team,
                            anchor_dt=effective_ack.sent_time,
                            max_dt=effective_ack.sent_time,
                            subject_norm=None,
                        )
                    if closest_from_ack:
                        closest_dt = _to_ist(closest_from_ack)
                        if (
                            closest_dt > created_dt
                            and closest_dt <= ack_dt
                            and (ack_dt - closest_dt) <= timedelta(hours=6)
                            and (closest_dt - created_dt) >= timedelta(minutes=30)
                        ):
                            created_time = closest_from_ack
                            created_src = "PARSED_FROM_ACK_BODY_CLOSEST"
                            notes = "Created from closest ack body request"
                    if isinstance(created_src, str) and created_src.startswith("PARSED_FROM_"):
                        closest_from_thread = _extract_thread_request_time_closest_before(
                            ordered,
                            ess_team,
                            anchor_dt=effective_ack.sent_time,
                            subject_norm=subject_norm,
                        )
                        if closest_from_thread:
                            closest_thread_dt = _to_ist(closest_from_thread)
                            created_dt = _to_ist(created_time)
                            if (
                                closest_thread_dt > created_dt
                                and closest_thread_dt <= ack_dt
                                and closest_thread_dt.date() == ack_dt.date()
                                and (ack_dt - closest_thread_dt) <= timedelta(hours=6)
                                and (closest_thread_dt - created_dt) >= timedelta(minutes=30)
                            ):
                                created_time = closest_from_thread
                                created_src = "PARSED_FROM_THREAD_CLOSEST"
                                notes = "Created from closest thread request"
        except Exception:
            pass

    # If no ack phrase, treat first ESS reply after created as ack (within 20 min)
    if created_time and effective_ack is None:
        for e in ordered:
            if _is_ess_sender(e, ess_team) and _to_ist(e.sent_time) >= _to_ist(created_time):
                delta = _to_ist(e.sent_time) - _to_ist(created_time)
                if delta <= timedelta(minutes=20):
                    effective_ack = e
                    ack_src = e.sender_email
                break

    if created_time and effective_ack:
        ack_dt = _to_ist(effective_ack.sent_time)
        created_dt = _to_ist(created_time)
        if created_dt > ack_dt:
            # If ack phrase exists earlier, prefer it and recompute request before it
            for e in ordered:
                if not _is_ess_sender(e, ess_team):
                    continue
                body = (e.body or "").lower()
                if not _is_ack_body(body):
                    continue
                if _to_ist(e.sent_time) > ack_dt:
                    break
                req_time, req_src = _latest_request_time_before(
                    ordered,
                    ess_team,
                    e.sent_time,
                    subject_norm=subject_norm,
                )
                if req_time:
                    effective_ack = e
                    ack_dt = _to_ist(e.sent_time)
                    created_time = req_time
                    created_src = req_src
                    created_dt = _to_ist(created_time)
                    notes = "Ack phrase preferred (created<=ack)"
                    break
            # Ensure created time is not after ack; recompute using ack time as cap
            req_time, req_src = _latest_request_time_before(
                ordered,
                ess_team,
                effective_ack.sent_time,
                subject_norm=subject_norm,
            )
            if req_time:
                created_time = req_time
                created_src = req_src
                created_dt = _to_ist(created_time)
        if ack_dt - created_dt <= timedelta(minutes=16):
            response_time = _format_time(effective_ack.sent_time)
            if ack_mail is None:
                notes = "Ack fallback OK"
            else:
                notes = "OK"
        else:
            response_time = "NA"
            notes = "Ack delayed (>16 min)"

        # If created came from quoted parsing, re-parse ack body to find the
        # closest request immediately below the ack (common Outlook chain).
        if created_src.startswith("PARSED_FROM_BODY"):
            parsed_from_ack = _extract_request_time_from_email(
                effective_ack,
                ess_team,
                max_dt=effective_ack.sent_time,
                subject_norm=subject_norm,
            )
            if parsed_from_ack:
                parsed_dt = _to_ist(parsed_from_ack)
                if parsed_dt <= ack_dt and (ack_dt - parsed_dt) <= timedelta(minutes=20):
                    # Prefer the request under the ack if it's closer to the ack than
                    # the current parsed request time.
                    if parsed_dt > created_dt:
                        created_time = parsed_from_ack
                        created_src = "PARSED_FROM_ACK_BODY"
                        created_dt = _to_ist(created_time)
                        # Recompute response_time and note based on new delta
                        if ack_dt - created_dt <= timedelta(minutes=16):
                            response_time = _format_time(effective_ack.sent_time)
                            notes = "OK"
                        else:
                            notes = "Ack delayed (>16 min)"

        # If created_time equals ack_time, try to parse request time from ack body
        if created_dt == ack_dt:
            parsed_from_ack = _extract_request_time_from_email(
                effective_ack,
                ess_team,
                max_dt=effective_ack.sent_time,
                subject_norm=subject_norm,
            )
            if parsed_from_ack:
                parsed_dt = _to_ist(parsed_from_ack)
                if parsed_dt <= ack_dt:
                    created_time = parsed_from_ack
                    created_src = "PARSED_FROM_ACK_BODY"
                    created_dt = _to_ist(created_time)
                    notes = "Created from ack body (equal time fix)"
    else:
        # No ack: use request email below requester reply (closest non-ESS before requester)
        if resolved_mail:
            candidates = [
                e for e in ordered
                if not _is_ess_sender(e, ess_team) and e.sent_time <= resolved_mail.sent_time
            ]
            if candidates:
                created_mail = max(candidates, key=lambda e: e.sent_time)
                created_time = created_mail.sent_time
                created_src = created_mail.sender_email
                notes = "Ack missing"
            else:
                parsed = _extract_request_time_from_email(
                    resolved_mail,
                    ess_team,
                    max_dt=resolved_mail.sent_time,
                    subject_norm=subject_norm,
                )
                if parsed:
                    created_time = parsed
                    created_src = "PARSED_FROM_BODY"
                    notes = "Ack missing; created parsed from body"
                else:
                    parsed_any = _find_latest_quoted_request_time(
                        ordered,
                        ess_team,
                        max_dt=resolved_mail.sent_time,
                        subject_norm=subject_norm,
                    )
                    if parsed_any:
                        created_time = parsed_any
                        created_src = "PARSED_FROM_BODY_ANY"
                        notes = "Ack missing; created parsed from any body"
        else:
            candidates = [e for e in ordered if not _is_ess_sender(e, ess_team)]
            if candidates:
                created_mail = max(candidates, key=lambda e: e.sent_time)
                created_time = created_mail.sent_time
                created_src = created_mail.sender_email
                notes = "Ack missing"
            else:
                parsed = _extract_request_time_from_email(
                    ordered[-1],
                    ess_team,
                    subject_norm=subject_norm,
                )
                if parsed:
                    created_time = parsed
                    created_src = "PARSED_FROM_BODY"
                    notes = "Ack missing; created parsed from body"
                else:
                    parsed_any = _find_latest_quoted_request_time(
                        ordered,
                        ess_team,
                        max_dt=ordered[-1].sent_time,
                        subject_norm=subject_norm,
                    )
                    if parsed_any:
                        created_time = parsed_any
                        created_src = "PARSED_FROM_BODY_ANY"
                        notes = "Ack missing; created parsed from any body"

        # If requester replied within 20 minutes, use that as response
        if created_time and resolved_mail:
            created_dt = _to_ist(created_time)
            resolved_dt = _to_ist(resolved_mail.sent_time)
            if resolved_dt - created_dt <= timedelta(minutes=20):
                response_time = _format_time(resolved_mail.sent_time)
                notes = "No ack; requester replied within 20 min"
            else:
                # Keep response populated for downstream blue-gap audit/repair.
                response_time = _format_time(resolved_mail.sent_time)
                notes = "No ack; requester reply >20 min; response=resolved"
        elif created_time and not resolved_mail:
            # If no requester reply but created exists, allow response/resolved to match created
            response_time = _format_time(created_time)
            notes = "No ack; no requester reply; response=created"

    # If Ack exists and a consultant reply occurs within 48h after Ack,
    # prefer that reply as the resolved time (even if later follow-ups exist).
    if effective_ack and requester_emails:
        try:
            ack_dt = _to_ist(effective_ack.sent_time)
            window_end = ack_dt + timedelta(hours=48)
            candidates = []
            for e in requester_emails:
                if not e.sent_time:
                    continue
                if not _is_real_reply_candidate(e):
                    continue
                sent_dt = _to_ist(e.sent_time)
                if sent_dt > ack_dt and sent_dt <= window_end:
                    candidates.append(e)
            if candidates:
                candidates.sort(key=lambda e: _to_ist(e.sent_time))
                resolved_mail = candidates[0]
                notes = (notes + "; ResolvedWithin48hAfterAck") if notes else "ResolvedWithin48hAfterAck"
        except Exception:
            pass

    if not created_time:
        # ESS-only threads with no parsed request: fallback to first email time
        if non_ess_emails == [] and ess_emails:
            scan_max_dt = resolved_mail.sent_time if resolved_mail and resolved_mail.sent_time else ordered[-1].sent_time
            parsed_any = _find_latest_quoted_request_time(
                ordered,
                ess_team,
                max_dt=scan_max_dt,
                subject_norm=subject_norm,
            )
            if parsed_any:
                created_time = parsed_any
            else:
                if not has_internal_marker:
                    if len(ess_emails) >= 2:
                        created_t = _format_time(ess_emails[0].sent_time)
                        non_ack_ess = [e for e in ess_emails if not _is_ack_like_reply(e)]
                        last_ess = max(non_ack_ess, key=lambda e: e.sent_time) if non_ack_ess else max(ess_emails, key=lambda e: e.sent_time)
                        resolved_t = _format_time(last_ess.sent_time) if last_ess else created_t
                        ack_t = resolved_t
                        for e in ess_emails:
                            body = (e.body or "").lower()
                            if _is_ack_body(body):
                                ack_t = _format_time(e.sent_time)
                                break
                        return (
                            TimeResult(created_t, ack_t, resolved_t),
                            TimeDebug(
                                ess_emails[0].sender_email or ess_emails[0].sender_name,
                                "ACK NOT FOUND" if ack_t == resolved_t else (ess_emails[0].sender_email or ess_emails[0].sender_name),
                                (last_ess.sender_email or last_ess.sender_name) if last_ess else (ess_emails[0].sender_email or ess_emails[0].sender_name),
                                "ESS only; no request found; span",
                            ),
                        )
                    t = _format_time(first.sent_time)
                    return (
                        TimeResult(t, t, t),
                        TimeDebug(first.sender_email, first.sender_email, first.sender_email, "ESS only; no request found"),
                    )
        return (
            TimeResult("", "NA", _format_time(resolved_mail.sent_time) if resolved_mail else ""),
            TimeDebug(created_src, ack_src, resolved_mail.sender_email if resolved_mail else "", notes),
        )

    # Clamp parsed created time if it's far earlier than the thread start.
    if created_time and created_src.startswith("PARSED_FROM_"):
        try:
            created_dt = _to_ist(created_time)
            first_dt = _to_ist(first.sent_time)
            if created_dt < (first_dt - timedelta(hours=12)):
                created_time = first.sent_time
                created_src = "CREATED_CLAMPED_TO_FIRST"
                notes = "Created clamped to first email"
                _trace_focus_row(
                    subject_norm,
                    "created_clamped_to_first",
                    created=_format_time(created_time),
                    created_src=created_src,
                    first=_format_time(first.sent_time),
                )
        except Exception:
            pass

    # Shared local-live guard:
    # when created came from quoted/body parsing, prefer a newer live request in the
    # same episode if the parsed request is clearly stale relative to ack/resolved.
    if created_time and isinstance(created_src, str) and created_src.startswith("PARSED_FROM_"):
        try:
            cap_dt = effective_ack.sent_time if effective_ack else (resolved_mail.sent_time if resolved_mail else None)
            live_req_time, live_req_src = _latest_live_request_before(
                ordered,
                ess_team,
                max_dt=cap_dt,
                requester_name=requester_name,
            )
            if cap_dt and live_req_time:
                created_dt = _to_ist(created_time)
                live_req_dt = _to_ist(live_req_time)
                cap_ist = _to_ist(cap_dt)
                parsed_gap = cap_ist - created_dt
                live_gap = cap_ist - live_req_dt
                if (
                    live_req_dt > created_dt
                    and parsed_gap > timedelta(hours=1)
                    and live_gap <= timedelta(hours=6)
                    and (live_req_dt - created_dt) >= timedelta(minutes=30)
                ):
                    created_time = live_req_time
                    created_src = live_req_src
                    notes = (notes + "; ParsedCreatedReanchoredToLive") if notes else "ParsedCreatedReanchoredToLive"
        except Exception:
            pass

    # Failed-subject safety: if parsed-created is too far from ack/resolved cap,
    # fallback to the latest real non-ESS mail before that cap.
    if failed_subject and created_time and isinstance(created_src, str) and created_src.startswith("PARSED_FROM_"):
        try:
            cap_dt = effective_ack.sent_time if effective_ack else (resolved_mail.sent_time if resolved_mail else None)
            if cap_dt:
                created_dt = _to_ist(created_time)
                cap_ist = _to_ist(cap_dt)
                if created_dt > cap_ist or (cap_ist - created_dt) > timedelta(hours=6):
                    real_candidates = [
                        e for e in ordered
                        if e.sent_time
                        and not _is_ess_sender(e, ess_team)
                        and _to_ist(e.sent_time) <= cap_ist
                    ]
                    if real_candidates:
                        real_req = max(real_candidates, key=lambda e: e.sent_time)
                        created_time = real_req.sent_time
                        created_src = real_req.sender_email or real_req.sender_name
                        notes = (notes + "; ParsedCreatedFallbackRealMail") if notes else "ParsedCreatedFallbackRealMail"
        except Exception:
            pass

    # Final ordering guard: created must be <= response/ack and <= resolved
    try:
        created_dt = _to_ist(created_time)
        # If we have an ack, ensure created is not after it
        if effective_ack:
            ack_dt = _to_ist(effective_ack.sent_time)
            if created_dt > ack_dt:
                req_time, req_src = _latest_request_time_before(
                    ordered,
                    ess_team,
                    effective_ack.sent_time,
                    subject_norm=subject_norm,
                )
                if req_time:
                    created_time = req_time
                    created_src = req_src
                    created_dt = _to_ist(created_time)
        # If we have a resolved mail, ensure created is not after it
        if resolved_mail:
            resolved_dt = _to_ist(resolved_mail.sent_time)
            if created_dt > resolved_dt:
                req_time, req_src = _latest_request_time_before(
                    ordered,
                    ess_team,
                    resolved_mail.sent_time,
                    subject_norm=subject_norm,
                )
                if req_time:
                    created_time = req_time
                    created_src = req_src
                    created_dt = _to_ist(created_time)
        # If response time is set but still earlier than created, try to realign created
        if response_time not in ("", "NA"):
            try:
                resp_dt = datetime.strptime(response_time, "%d-%m-%Y %H:%M").replace(tzinfo=IST)
                if resp_dt < created_dt:
                    gap = created_dt - resp_dt
                    if gap <= timedelta(hours=12):
                        req_time, req_src = _latest_request_time_before(
                            ordered,
                            ess_team,
                            resp_dt,
                            subject_norm=subject_norm,
                        )
                        if req_time:
                            reliable_req = True
                            if isinstance(req_src, str) and req_src.startswith("PARSED_"):
                                # Parsed quoted times can point to stale chain context.
                                # Only trust parsed request anchors when we also have a
                                # nearby non-ESS email around that parsed timestamp.
                                reliable_req = _has_non_ess_email_near_time(
                                    ordered,
                                    ess_team,
                                    req_time,
                                    window_hours=24,
                                )
                            if reliable_req:
                                created_time = req_time
                                created_src = req_src
                                created_dt = _to_ist(created_time)
                                notes = "Created anchored to response time"
                            else:
                                notes = "Created retained (response anchor unreliable)"
                                _trace_focus_row(
                                    subject_norm,
                                    "created_retained_response_anchor_unreliable",
                                    created=_format_time(created_time),
                                    response=response_time,
                                    resolved=resolved_time,
                                    created_src=created_src,
                                    ack_src=ack_src,
                                )
                        else:
                            # Keep original created when no request can be found.
                            # This avoids converting "request created" into ack/response time.
                            notes = "Created retained (no request found before response)"
                    else:
                        notes = "Created retained (response earlier)"
            except Exception:
                pass
    except Exception:
        pass

    # If requester reply is missing, but we do have ESS replies, prefer the last ESS
    # reply for resolved time to avoid collapsing all three times.
    resolved_time = _format_time(resolved_mail.sent_time) if resolved_mail else ""
    if not resolved_mail and created_time and ess_emails:
        non_ack_ess = [e for e in ess_emails if not _is_ack_like_reply(e)]
        last_ess = max(non_ack_ess, key=lambda e: e.sent_time) if non_ack_ess else max(ess_emails, key=lambda e: e.sent_time)
        if last_ess and _to_ist(last_ess.sent_time) >= _to_ist(created_time):
            resolved_time = _format_time(last_ess.sent_time)
            if response_time in ("", "NA") or response_time == _format_time(created_time):
                resp_candidate = None
                for e in ordered:
                    if not _is_ess_sender(e, ess_team):
                        continue
                    if _to_ist(e.sent_time) < _to_ist(created_time):
                        continue
                    if _is_ack_body((e.body or "").lower()):
                        resp_candidate = e
                        break
                    if resp_candidate is None:
                        resp_candidate = e
                if resp_candidate:
                    response_time = _format_time(resp_candidate.sent_time)
            if "No requester match" not in notes:
                notes = notes + "; No requester match"

    if not resolved_time and created_time:
        resolved_time = _format_time(created_time)

    # Failed-subject ESS-only safety: when created came from parsed quote but the
    # requester thread is short, prefer requester-first as created anchor.
    if (
        failed_subject
        and not non_ess_emails
        and requester_first
        and requester_emails
        and created_time
        and isinstance(created_src, str)
        and created_src.startswith("PARSED_FROM_")
    ):
        try:
            first_req_dt = _to_ist(requester_first.sent_time)
            last_req_dt = _to_ist(requester_emails[-1].sent_time)
            created_dt = _to_ist(created_time)
            if (last_req_dt - first_req_dt) <= timedelta(hours=48):
                if created_dt > first_req_dt + timedelta(minutes=30):
                    created_time = requester_first.sent_time
                    created_src = requester_first.sender_email or requester_first.sender_name
                    notes = (notes + "; FailedCreatedFromRequesterFirst") if notes else "FailedCreatedFromRequesterFirst"
        except Exception:
            pass

    # Narrow episode re-anchor guard:
    # In long chains, if we already picked an older request/ack pair but there is a
    # clearly newer requester ack-like reply with a nearby request before it, re-anchor
    # created/ack to that newer episode (keep resolved unchanged).
    try:
        if created_time and response_time not in ("", "NA") and resolved_mail and requester_emails:
            created_dt = _to_ist(created_time)
            response_dt = datetime.strptime(response_time, "%d-%m-%Y %H:%M").replace(tzinfo=IST)
            resolved_dt = _to_ist(resolved_mail.sent_time)
            # Keep this strict to avoid disturbing already-correct rows.
            min_episode_gap = timedelta(hours=6)
            max_ack_after_request = timedelta(minutes=30)

            episode_candidates = []
            for e in requester_emails:
                if not e.sent_time:
                    continue
                e_dt = _to_ist(e.sent_time)
                if e_dt <= response_dt + min_episode_gap:
                    continue
                if e_dt > resolved_dt:
                    continue
                if not _is_ack_body((e.body or "").lower()):
                    continue

                req_time, req_src = _latest_request_time_before(
                    ordered,
                    ess_team,
                    e.sent_time,
                    subject_norm=subject_norm,
                )
                req_dt = _to_ist(req_time) if req_time else None

                # If live-chain request anchor is missing or too old for this newer ack-like
                # requester mail, try to extract request time from this same mail body.
                # This keeps scope tight and avoids disturbing unrelated rows.
                req_from_ack_body = _extract_request_time_from_email(
                    e,
                    ess_team,
                    max_dt=e.sent_time,
                    subject_norm=subject_norm,
                )
                req_from_ack_body_dt = _to_ist(req_from_ack_body) if req_from_ack_body else None

                use_body_req = False
                if req_from_ack_body_dt:
                    if (
                        req_from_ack_body_dt > created_dt + min_episode_gap
                        and req_from_ack_body_dt <= e_dt
                        and (e_dt - req_from_ack_body_dt) <= max_ack_after_request
                        and req_from_ack_body_dt <= resolved_dt
                    ):
                        if (not req_dt) or (req_dt <= created_dt + min_episode_gap) or ((e_dt - req_dt) > max_ack_after_request):
                            use_body_req = True

                if use_body_req:
                    req_time = req_from_ack_body
                    req_src = "PARSED_FROM_ACK_BODY_EPISODE"
                    req_dt = req_from_ack_body_dt
                elif not req_dt:
                    continue

                if req_dt <= created_dt + min_episode_gap:
                    continue
                if req_dt > e_dt:
                    continue
                if (e_dt - req_dt) > max_ack_after_request:
                    continue
                if req_dt > resolved_dt:
                    continue

                episode_candidates.append((e_dt, e, req_time, req_src))

            if episode_candidates:
                # Prefer the latest valid re-anchoring episode before resolved.
                episode_candidates.sort(key=lambda x: x[0], reverse=True)
                _ack_dt, ack_mail_new, req_time_new, req_src_new = episode_candidates[0]
                created_time = req_time_new
                created_src = req_src_new
                response_time = _format_time(ack_mail_new.sent_time)
                ack_src = ack_mail_new.sender_email or ack_mail_new.sender_name
                notes = (notes + "; ReAnchoredNewerRequestAckEpisode") if notes else "ReAnchoredNewerRequestAckEpisode"
    except Exception:
        pass

    result = TimeResult(
        _format_time(created_time),
        response_time,
        resolved_time,
    )
    debug = TimeDebug(
        created_src or (created_mail.sender_email if created_mail else ""),
        ack_src,
        resolved_mail.sender_email if resolved_mail else (created_src or ""),
        notes,
    )
    _trace_focus_row(
        subject_norm,
        "resolve_times_with_debug:return",
        created=result.created,
        response=result.response,
        resolved=result.resolved,
        created_src=debug.created_src,
        ack_src=debug.ack_src,
        resolved_src=debug.resolved_src,
        notes=debug.notes,
    )
    return (result, debug)


def _has_ack_phrase(ordered, ess_team) -> bool:
    for e in ordered:
        if _is_ess_sender(e, ess_team):
            body = (e.body or "").lower()
            if _is_ack_body(body):
                return True
    return False


def _is_failed_subject(emails) -> bool:
    keywords = [
        "failed",
        "file failed",
        "failed to process",
        "did not process",
        "not processed completely",
        "not processed",
        "records got skipped at es prod",
        "skipped at es",
    ]
    for e in emails:
        subj = (e.subject or "").lower()
        if any(k in subj for k in keywords):
            return True
    return False


def _find_latest_quoted_request_time(
    emails,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    latest = None
    for e in emails:
        parsed = _extract_request_time_from_email(
            e,
            ess_team,
            max_dt=max_dt or e.sent_time,
            subject_norm=subject_norm,
        )
        if not parsed:
            continue
        if max_dt and _to_ist(parsed) > _to_ist(max_dt):
            continue
        if latest is None or _to_ist(parsed) > _to_ist(latest):
            latest = parsed
    return latest


def _select_ack_and_request(emails, ess_team, max_time: datetime | None = None, subject_norm: str | None = None):
    candidates = []
    for e in emails:
        if not _is_ess_sender(e, ess_team):
            continue
        if max_time and _to_ist(e.sent_time) > _to_ist(max_time):
            continue
        req_time, req_src = _latest_request_time_before(emails, ess_team, e.sent_time, subject_norm=subject_norm)
        if not req_time:
            continue
        delta = _to_ist(e.sent_time) - _to_ist(req_time)
        if delta.total_seconds() < 0:
            continue
        body = (e.body or "").lower()
        has_phrase = _is_ack_body(body)
        candidates.append((delta, has_phrase, e, req_time, req_src))

    if not candidates:
        return None

    # Prefer ack with phrase; then smallest delta
    candidates.sort(key=lambda x: (not x[1], x[0]))
    delta, has_phrase, ack_email, req_time, req_src = candidates[0]
    note = "Ack phrase" if has_phrase else "Ack fallback closest"
    return ack_email, req_time, req_src, note


def _first_ess_after(emails, ess_team, after_time: datetime, max_time: datetime | None):
    candidates = []
    for e in emails:
        if not _is_ess_sender(e, ess_team):
            continue
        if _to_ist(e.sent_time) < _to_ist(after_time):
            continue
        if max_time and _to_ist(e.sent_time) > _to_ist(max_time):
            continue
        body = (e.body or "").lower()
        has_phrase = _is_ack_body(body)
        candidates.append((has_phrase, e))

    if not candidates:
        return None

    # Prefer any ESS reply within 16 minutes of the request (even without ack phrase).
    window = timedelta(minutes=16)
    within = []
    for has_phrase, e in candidates:
        if _to_ist(e.sent_time) - _to_ist(after_time) <= window:
            within.append((has_phrase, e))
    if within:
        within.sort(key=lambda x: _to_ist(x[1].sent_time))
        return within[0][1]

    # Otherwise, prefer ack phrase; if none, take earliest ESS after request
    candidates.sort(key=lambda x: (not x[0], _to_ist(x[1].sent_time)))
    return candidates[0][1]


def _find_direct_resolution_reply(
    emails,
    ess_team,
    after_time: datetime | None = None,
    requester_name: str | None = None,
):
    candidates = []
    for e in emails:
        if not _is_ess_sender(e, ess_team):
            continue
        if after_time and _to_ist(e.sent_time) < _to_ist(after_time):
            continue
        body = (e.body or "").lower()
        if any(p in body for p in DIRECT_RESOLUTION_PHRASES):
            candidates.append(e)
    if not candidates:
        return None
    if requester_name:
        requester_matches = [
            e for e in candidates
            if _match_requester(e.sender_name, e.sender_email, requester_name)
        ]
        if requester_matches:
            return min(requester_matches, key=lambda x: _to_ist(x.sent_time))
    return min(candidates, key=lambda x: _to_ist(x.sent_time))


def _latest_request_time_before(
    emails,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    latest_time = None
    latest_src = ""
    found_non_ess = False
    latest_parsed_time = None

    # Non-ESS emails
    for e in emails:
        if _is_ess_sender(e, ess_team):
            continue
        if max_dt and _to_ist(e.sent_time) > _to_ist(max_dt):
            continue
        found_non_ess = True
        if latest_time is None or _to_ist(e.sent_time) > _to_ist(latest_time):
            latest_time = e.sent_time
            latest_src = e.sender_email or e.sender_name

    # Quoted request times in email body/body_html.
    # Safe rule:
    # - If no real non-ESS request exists, keep existing behavior (use parsed when found).
    # - If subject anchor exists, allow parsed quoted request to win only when newer.
    if (not found_non_ess) or bool(subject_norm):
        for e in emails:
            parsed = _extract_request_time_from_email(
                e,
                ess_team,
                max_dt=max_dt or e.sent_time,
                subject_norm=subject_norm,
            )
            if not parsed:
                continue
            if max_dt and _to_ist(parsed) > _to_ist(max_dt):
                continue
            if latest_parsed_time is None or _to_ist(parsed) > _to_ist(latest_parsed_time):
                latest_parsed_time = parsed

    # Real non-ESS request precedence:
    # If a live non-ESS request mail exists, do not replace it with parsed quoted
    # request text from body/history in mixed threads.
    if latest_parsed_time:
        if latest_time is None:
            latest_time = latest_parsed_time
            latest_src = "PARSED_FROM_BODY"
        elif not found_non_ess and _to_ist(latest_parsed_time) > _to_ist(latest_time):
            latest_time = latest_parsed_time
            latest_src = "PARSED_FROM_BODY"

    return latest_time, latest_src


def _latest_request_time_before_internal(
    emails,
    ess_team,
    requester_name: str,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    latest_time = None
    latest_src = ""
    latest_parsed_time = None
    found_live_request = False

    # Non-ESS emails OR internal requester emails (even if ESS)
    for e in emails:
        if _is_ess_sender(e, ess_team) and not _match_requester(e.sender_name, e.sender_email, requester_name):
            continue
        if max_dt and _to_ist(e.sent_time) > _to_ist(max_dt):
            continue
        if latest_time is None or _to_ist(e.sent_time) > _to_ist(latest_time):
            latest_time = e.sent_time
            latest_src = e.sender_email or e.sender_name
            found_live_request = True

    # Quoted request times in email body/body_html.
    for e in emails:
        parsed = _extract_request_time_from_email(
            e,
            ess_team,
            max_dt=max_dt or e.sent_time,
            subject_norm=subject_norm,
        )
        if not parsed:
            continue
        if max_dt and _to_ist(parsed) > _to_ist(max_dt):
            continue
        if latest_parsed_time is None or _to_ist(parsed) > _to_ist(latest_parsed_time):
            latest_parsed_time = parsed

    # Keep live request precedence for internal-request aware path too.
    if latest_parsed_time:
        if latest_time is None:
            latest_time = latest_parsed_time
            latest_src = "PARSED_FROM_BODY"
        elif (not found_live_request) and _to_ist(latest_parsed_time) > _to_ist(latest_time):
            latest_time = latest_parsed_time
            latest_src = "PARSED_FROM_BODY"

    return latest_time, latest_src


def _is_special_case(thread, ess_team) -> bool:
    if not thread:
        return False
    first = thread[0]
    if not _is_ess_sender(first, ess_team):
        return False

    non_ess = [e for e in thread if not _is_ess_sender(e, ess_team)]
    if not non_ess:
        return False

    return all(_is_system_sender(e) for e in non_ess)


def _requester_email_candidates(requester_name: str, ess_team) -> set:
    tokens = _tokenize(requester_name)
    if not tokens:
        return set()
    candidates = set()
    for email in ess_team:
        local_tokens = set(_tokenize(_email_local(email)))
        if not local_tokens:
            continue
        overlap = set(tokens) & local_tokens
        if len(overlap) >= 2:
            candidates.add(email)
        elif len(overlap) == 1 and tokens[-1] in local_tokens:
            candidates.add(email)
    return candidates


def _is_ess_sender(email_record, ess_team) -> bool:
    sender_email = (email_record.sender_email or "").lower()
    if "enterprise-services-support" in sender_email:
        return True
    if sender_email and sender_email in ess_team:
        return True
    sender_name = (email_record.sender_name or "").lower()
    if "enterprise-services-support" in sender_name:
        return True
    if not sender_name:
        return False

    sender_tokens = _tokenize(sender_name)
    name_tokens = set(sender_tokens)
    last_token = sender_tokens[-1] if sender_tokens else ""
    if not name_tokens:
        return False

    for email in ess_team:
        local_tokens = set(_tokenize(_email_local(email)))
        if not local_tokens:
            continue
        overlap = name_tokens & local_tokens
        if len(overlap) >= 2:
            return True
        if overlap and last_token and last_token in local_tokens:
            return True
    return False


def _is_system_sender(email_record) -> bool:
    sender_email = (email_record.sender_email or "").lower()
    sender_name = (email_record.sender_name or "").lower()
    subject = (getattr(email_record, "subject", "") or "").lower()
    tokens = [
        "eai-system-notification",
        "system-notification",
        "system notification",
        "no-reply",
        "noreply",
        "do-not-reply",
        "donotreply",
    ]
    return (
        any(t in sender_email for t in tokens)
        or any(t in sender_name for t in tokens)
        or any(t in subject for t in tokens)
    )


def _extract_sent_time_from_body(body: str, max_dt: datetime | None = None):
    if not body:
        return None

    lines = [line.strip() for line in body.splitlines() if line.strip()]
    sent_lines = [l for l in lines if _is_sent_header_line(l)]
    if not sent_lines:
        # common Outlook block: "Sent: Tuesday, January 2, 2026 3:45 PM"
        sent_lines = [l for l in lines if "sent:" in l.lower() or "envoy" in l.lower()]

    candidates = []
    for line in sent_lines:
        try:
            value = line.split(":", 1)[1].strip()
        except Exception:
            continue
        dt = _parse_datetime(value)
        if dt:
            candidates.append(dt)

    if candidates:
        if max_dt:
            max_dt = _to_ist(max_dt)
            candidates = [c for c in candidates if _to_ist(c) <= max_dt]
        if candidates:
            return max(candidates)

    # Try to parse Outlook-style quoted header blocks (From/Sent/To/Subject)
    block_dt = _extract_outlook_block_sent(lines, max_dt)
    if block_dt:
        return block_dt

    # Try to extract date-time patterns from body as fallback
    patterns = [
        r"(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)?)",
        r"(\d{1,2}-\w{3}-\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)?)",
        r"(\w+,\s+\w+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM))",
    ]
    for pat in patterns:
        m = re.search(pat, body, flags=re.IGNORECASE)
        if not m:
            continue
        dt = _parse_datetime(m.group(1))
        if dt:
            if max_dt and _to_ist(dt) > _to_ist(max_dt):
                continue
            return dt

    return None


def _extract_request_time_from_body(
    body: str,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    if not body:
        return None
    lines = _body_to_clean_lines(body)
    block_dt = _extract_outlook_block_sent(
        lines,
        max_dt,
        ess_team=ess_team,
        require_non_ess=True,
        exclude_system=True,
        subject_norm=subject_norm,
    )
    if block_dt:
        return block_dt
    return _extract_sent_from_context(lines, ess_team, max_dt, subject_norm=subject_norm)


def _extract_request_time_candidates_from_body(
    body: str,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    if not body:
        return []
    lines = _body_to_clean_lines(body)
    candidates = []
    block_candidates = _extract_outlook_block_sent_candidates(
        lines,
        max_dt,
        ess_team=ess_team,
        require_non_ess=True,
        exclude_system=True,
        subject_norm=subject_norm,
    )
    candidates.extend(block_candidates)
    candidates.extend(_extract_sent_from_context_candidates(lines, ess_team, max_dt, subject_norm=subject_norm))
    unique = []
    seen = set()
    for dt in candidates:
        try:
            key = _to_ist(dt)
        except Exception:
            continue
        if key in seen:
            continue
        seen.add(key)
        unique.append(dt)
    return unique


def _extract_request_time_from_email(
    email_record,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    candidates = []
    body = getattr(email_record, "body", "") or ""
    body_html = getattr(email_record, "body_html", "") or ""

    parsed = _extract_request_time_from_body(
        body,
        ess_team,
        max_dt=max_dt,
        subject_norm=subject_norm,
    )
    if parsed:
        candidates.append(parsed)

    if body_html and body_html != body:
        parsed_html = _extract_request_time_from_body(
            body_html,
            ess_team,
            max_dt=max_dt,
            subject_norm=subject_norm,
        )
        if parsed_html:
            candidates.append(parsed_html)

    if not candidates:
        return None
    return max(candidates, key=_to_ist)


def _extract_request_time_from_email_closest(
    email_record,
    ess_team,
    anchor_dt: datetime,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    if not anchor_dt:
        return None
    candidates = []
    body = getattr(email_record, "body", "") or ""
    body_html = getattr(email_record, "body_html", "") or ""
    candidates.extend(
        _extract_request_time_candidates_from_body(
            body,
            ess_team,
            max_dt=max_dt,
            subject_norm=subject_norm,
        )
    )
    if body_html and body_html != body:
        candidates.extend(
            _extract_request_time_candidates_from_body(
                body_html,
                ess_team,
                max_dt=max_dt,
                subject_norm=subject_norm,
            )
        )
    if not candidates:
        return None
    anchor_ist = _to_ist(anchor_dt)
    filtered = []
    for dt in candidates:
        dt_ist = _to_ist(dt)
        if dt_ist <= anchor_ist:
            filtered.append(dt)
    if not filtered:
        return None
    return min(filtered, key=lambda dt: (anchor_ist - _to_ist(dt), -_to_ist(dt).timestamp()))


def _extract_thread_request_time_closest_before(
    emails,
    ess_team,
    anchor_dt: datetime,
    subject_norm: str | None = None,
):
    if not anchor_dt:
        return None
    anchor_ist = _to_ist(anchor_dt)
    candidates = []
    for e in emails:
        if not getattr(e, "sent_time", None):
            continue
        parsed = _extract_request_time_from_email_closest(
            e,
            ess_team,
            anchor_dt=anchor_dt,
            max_dt=anchor_dt,
            subject_norm=subject_norm,
        )
        if parsed:
            candidates.append(parsed)
    if not candidates and subject_norm:
        for e in emails:
            if not getattr(e, "sent_time", None):
                continue
            parsed = _extract_request_time_from_email_closest(
                e,
                ess_team,
                anchor_dt=anchor_dt,
                max_dt=anchor_dt,
                subject_norm=None,
            )
            if parsed:
                candidates.append(parsed)
    if not candidates:
        return None
    filtered = []
    seen = set()
    for dt in candidates:
        dt_ist = _to_ist(dt)
        if dt_ist > anchor_ist:
            continue
        if dt_ist in seen:
            continue
        seen.add(dt_ist)
        filtered.append(dt)
    if not filtered:
        return None
    return min(filtered, key=lambda dt: (anchor_ist - _to_ist(dt), -_to_ist(dt).timestamp()))


@dataclass
class _QuotedHeaderCandidate:
    from_line: str
    sent_line: str
    to_line: str
    subject_line: str
    block_lines: tuple[str, ...]
    sent_dt: datetime


def _normalize_quoted_sent_text(value: str) -> str:
    value = re.sub(r"(?i)^sent\b\s*:?\s*", "", value or "").strip()
    value = re.sub(r"(?i)^(mon|tue|wed|thu|fri|sat|sun)\w*,?\s*", "", value).strip()
    value = re.sub(r"\(.*?\)", " ", value).strip()
    value = re.sub(r"(?i)(\d)(am|pm)\b", r"\1 \2", value)
    value = re.sub(r"(?i)\b(a\.m\.|p\.m\.)\b", lambda m: m.group(1).replace(".", "").upper(), value)
    return " ".join(value.replace(",", " ").replace(" at ", " ").split())


def _parse_quoted_header_datetime(value: str):
    if len(_quoted_header_datetime_cache) > 5000:
        _evict_keys = list(_quoted_header_datetime_cache.keys())
        random.shuffle(_evict_keys)
        for _evict_k in _evict_keys[:1000]:
            del _quoted_header_datetime_cache[_evict_k]
    if value in _quoted_header_datetime_cache:
        return _quoted_header_datetime_cache[value]
    raw_clean = re.sub(r"(?i)^sent\b\s*:?\s*", "", (value or "")).strip()
    normalized = _normalize_quoted_sent_text(raw_clean)

    def _extract_primary_sent_fragments(*values: str):
        regexes = [
            r"\b[A-Za-z]+,\s+\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*[APMapm]{2}\b",
            r"\b[A-Za-z]+,\s+\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
            r"\b[A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*[APMapm]{2}\b",
            r"\b[A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
            r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*[APMapm]{2}\b",
            r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
            r"\b\d{1,2}[-./]\d{1,2}[-./]\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*[APMapm]{2}\b",
            r"\b\d{1,2}[-./]\d{1,2}[-./]\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
            r"\b\d{4}[-./]\d{1,2}[-./]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?\s*[APMapm]{2}\b",
            r"\b\d{4}[-./]\d{1,2}[-./]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
        ]
        out = []
        seen = set()
        for candidate_value in values:
            text = (candidate_value or "").strip()
            if not text:
                continue
            matches = []
            for rx in regexes:
                for match in re.finditer(rx, text):
                    frag = match.group(0).strip()
                    key = frag.lower()
                    if key in seen:
                        continue
                    matches.append((match.start(), -len(frag), frag))
            matches.sort()
            for _, _, frag in matches:
                key = frag.lower()
                if key in seen:
                    continue
                seen.add(key)
                out.append(frag)
            if out:
                break
        return out

    candidates = []
    for cand in _extract_primary_sent_fragments(raw_clean, normalized):
        if cand and cand not in candidates:
            candidates.append(cand)
    for cand in (normalized, raw_clean):
        if cand and cand not in candidates:
            candidates.append(cand)

    fmts_24h = [
        "%d %B %Y %H:%M:%S",
        "%d %B %Y %H:%M",
        "%d %b %Y %H:%M:%S",
        "%d %b %Y %H:%M",
        "%B %d %Y %H:%M:%S",
        "%B %d %Y %H:%M",
        "%b %d %Y %H:%M:%S",
        "%b %d %Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%A %d %B %Y %H:%M",
        "%A %d %B %Y %H:%M:%S",
        "%A %B %d %Y %H:%M",
        "%A %B %d %Y %H:%M:%S",
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
        "%A %B %d %Y %I:%M %p",
        "%A %B %d %Y %I:%M:%S %p",
        "%A %d %B %Y %I:%M %p",
        "%A %d %B %Y %I:%M:%S %p",
    ]
    for cand in candidates:
        cand_has_am_pm = bool(re.search(r"(?i)\b(am|pm|a\.m\.|p\.m\.)\b", cand or ""))
        parse_fmts = fmts_12h if cand_has_am_pm else fmts_24h
        for fmt in parse_fmts:
            try:
                parsed = datetime.strptime(cand, fmt)
                _quoted_header_datetime_cache[value] = parsed
                return parsed
            except Exception:
                continue
    for cand in candidates:
        cand_has_am_pm = bool(re.search(r"(?i)\b(am|pm|a\.m\.|p\.m\.)\b", cand or ""))
        try:
            parsed = parsedate_to_datetime(cand)
            if parsed:
                _quoted_header_datetime_cache[value] = parsed
                return parsed
        except Exception:
            pass
        if not cand_has_am_pm:
            for fmt in fmts_12h:
                try:
                    parsed = datetime.strptime(cand, fmt)
                    _quoted_header_datetime_cache[value] = parsed
                    return parsed
                except Exception:
                    continue
    _quoted_header_datetime_cache[value] = None
    return None


def _extract_bounded_outlook_header_candidates(lines, allow_relaxed: bool = False):
    cache_key = (tuple(lines or ()), bool(allow_relaxed))
    cached = _bounded_outlook_header_cache.get(cache_key)
    if cached is not None:
        return list(cached)
    if not lines:
        _bounded_outlook_header_cache[cache_key] = ()
        return []

    def _header_label(line: str):
        m = re.match(r"(?i)^(from|sent|to|cc|bcc|subject|objet)\b\s*:?\s*(.*)$", line or "")
        return m.group(1).lower() if m else None

    def _header_value(line: str, label: str) -> str:
        m = re.match(rf"(?i)^{label}\b\s*:?\s*(.*)$", line or "")
        return (m.group(1) if m else "").strip()

    def _from_start(line: str) -> bool:
        if allow_relaxed:
            return bool(re.search(r"(?i)\bfrom\b\s*:?", line or ""))
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

    def _looks_like_from_value(text: str) -> bool:
        text = (text or "").strip()
        if not text or _looks_like_body_line(text):
            return False
        if re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, flags=re.I):
            return True
        if re.search(r"<[^>]+>", text):
            return True
        parts = [p for p in re.split(r"[^A-Za-z0-9]+", text) if p]
        return 1 <= len(parts) <= 6

    def _append_meridiem(sent_line: str, start_idx: int) -> str:
        sent_low = (sent_line or "").lower()
        if (" am" in sent_low) or (" pm" in sent_low):
            return sent_line
        for j in range(start_idx + 1, min(start_idx + 4, len(lines))):
            extra = (lines[j] or "").strip()
            if re.fullmatch(r"(?i)(am|pm|a\.m\.|p\.m\.)", extra):
                return f"{sent_line} {extra}"
        return sent_line

    def _collect_header_value(start_idx: int, label: str, *, validator=None, max_continuations: int = 8):
        base_line = (lines[start_idx] or "").strip()
        value = _header_value(base_line, label)
        pieces = [value] if value else []
        last_idx = start_idx
        for j in range(start_idx + 1, min(start_idx + 1 + max_continuations, len(lines))):
            nxt = (lines[j] or "").strip()
            if not nxt:
                break
            if _header_label(nxt):
                break
            if _from_start(nxt):
                break
            if _looks_like_body_line(nxt):
                break
            candidate = " ".join(p for p in pieces + [nxt] if p).strip()
            if validator and not validator(candidate):
                break
            pieces.append(nxt)
            last_idx = j
        return " ".join(p for p in pieces if p).strip(), last_idx

    out = []
    i = 0
    while i < len(lines):
        if not _from_start(lines[i]):
            i += 1
            continue
        from_line = ""
        sent_line = ""
        to_line = ""
        subject_line = ""
        header_end = i
        for j in range(i, min(i + 32, len(lines))):
            cur = (lines[j] or "").strip()
            if j > i and (_from_start(cur) or re.match(r"(?i)^[-_]{3,}$", cur)):
                break
            label = _header_label(cur)
            if label == "from" and not from_line:
                value, consumed_idx = _collect_header_value(
                    j,
                    "from",
                    validator=_looks_like_from_value,
                    max_continuations=6,
                )
                from_line = f"From: {value}" if value else "From:"
                header_end = max(header_end, consumed_idx)
            elif label == "sent" and not sent_line:
                value, consumed_idx = _collect_header_value(j, "sent", max_continuations=4)
                sent_line = f"Sent: {value}" if value else "Sent:"
                header_end = max(header_end, consumed_idx)
            elif label in {"to", "cc", "bcc"} and not to_line:
                value = _header_value(cur, label)
                to_line = f"{label.capitalize()}: {value}" if value else f"{label.capitalize()}:"
            elif label in {"subject", "objet"} and not subject_line:
                value, consumed_idx = _collect_header_value(j, label, max_continuations=4)
                prefix = "Subject" if label == "subject" else "Objet"
                subject_line = f"{prefix}: {value}" if value else f"{prefix}:"
                header_end = max(header_end, consumed_idx)
            elif j > i and label is None and _looks_like_body_line(cur):
                header_end = j - 1
                break
            elif j > i and label is None and sent_line and (subject_line or from_line):
                header_end = j - 1
                break
            header_end = j

        if sent_line:
            sent_line = _append_meridiem(sent_line, header_end)
            sent_dt = _parse_quoted_header_datetime(sent_line)
            if sent_dt:
                block_lines = tuple((lines[k] or "").strip() for k in range(i, min(header_end + 1, len(lines))))
                out.append(
                    _QuotedHeaderCandidate(
                        from_line=from_line,
                        sent_line=sent_line,
                        to_line=to_line,
                        subject_line=subject_line,
                        block_lines=block_lines,
                        sent_dt=sent_dt,
                    )
                )
        i = max(i + 1, header_end + 1)
    if len(_bounded_outlook_header_cache) > 2000:
        _evict_keys = list(_bounded_outlook_header_cache.keys())
        random.shuffle(_evict_keys)
        for _evict_k in _evict_keys[:500]:
            del _bounded_outlook_header_cache[_evict_k]
    _bounded_outlook_header_cache[cache_key] = tuple(out)
    return list(_bounded_outlook_header_cache[cache_key])


def _extract_canonical_message_lines(email_record) -> list[str]:
    cache_key = _email_stable_key(email_record)
    cached = _canonical_message_lines_cache.get(cache_key)
    if cached is not None:
        return list(cached)

    html_source = (
        f"{getattr(email_record, 'body_html', '') or ''}\n"
        f"{getattr(email_record, 'body_html_raw', '') or ''}"
    ).strip()
    raw = (
        f"{getattr(email_record, 'body', '') or ''}\n"
        f"{getattr(email_record, 'body_html', '') or ''}\n"
        f"{getattr(email_record, 'body_html_raw', '') or ''}"
    )
    if not raw.strip():
        _canonical_message_lines_cache[cache_key] = ()
        return []

    raw_lines = _fallback_visible_lines_from_text(raw)
    bs4_lines = _bs4_visible_lines_from_text(html_source)

    # Prefer the human-visible BS4 conversation lines whenever available.
    # Fall back to raw cleaned lines only when HTML extraction produced nothing.
    lines = bs4_lines or raw_lines

    _canonical_message_lines_cache[cache_key] = tuple(lines)
    return list(_canonical_message_lines_cache[cache_key])


def _extract_canonical_current_text(email_record, max_lines: int = 48) -> str:
    cache_key = (_email_stable_key(email_record), max_lines)
    cached = _canonical_current_text_cache.get(cache_key)
    if cached is not None:
        return cached

    html_source = (
        f"{getattr(email_record, 'body_html', '') or ''}\n"
        f"{getattr(email_record, 'body_html_raw', '') or ''}"
    ).strip()
    lines = []
    if html_source:
        lines = _bs4_visible_lines_from_text(html_source)
        if not lines:
            lines = _fallback_visible_lines_from_text(html_source)
    if not lines:
        top = _leading_body_segment(getattr(email_record, "body", "") or "")
        lines = [ln.strip() for ln in top.splitlines() if ln and ln.strip()]

    current_lines = []
    header_re = re.compile(r"(?i)^(from|sent|to|cc|bcc|subject|objet)\s*:?\s*$")
    header_with_value_re = re.compile(r"(?i)^(from|sent|to|cc|bcc|subject|objet)\s*:")
    for line in lines:
        if header_re.match(line) or header_with_value_re.match(line):
            break
        current_lines.append(line)
        if len(current_lines) >= max_lines:
            break
    out = "\n".join(current_lines).strip()
    _canonical_current_text_cache[cache_key] = out
    return out


def _extract_canonical_quoted_header_candidates(email_record, allow_relaxed: bool = False):
    cache_key = (_email_stable_key(email_record), bool(allow_relaxed))
    cached = _canonical_quoted_header_candidates_cache.get(cache_key)
    if cached is not None:
        return list(cached)

    candidates = []
    lines = _extract_canonical_message_lines(email_record)
    if lines:
        candidates = _extract_bounded_outlook_header_candidates(lines, allow_relaxed=allow_relaxed)

    if not candidates:
        raw = (
            f"{getattr(email_record, 'body', '') or ''}\n"
            f"{getattr(email_record, 'body_html', '') or ''}\n"
            f"{getattr(email_record, 'body_html_raw', '') or ''}"
        )
        if raw.strip():
            raw_lines = _fallback_visible_lines_from_text(raw)
            if raw_lines:
                candidates = _extract_bounded_outlook_header_candidates(raw_lines, allow_relaxed=allow_relaxed)

    _canonical_quoted_header_candidates_cache[cache_key] = tuple(candidates)
    return list(_canonical_quoted_header_candidates_cache[cache_key])


def _extract_outlook_block_sent(
    lines,
    max_dt: datetime | None = None,
    ess_team=None,
    require_non_ess: bool = False,
    exclude_system: bool = False,
    subject_norm: str | None = None,
):
    bounded_candidates = []
    for candidate in _extract_bounded_outlook_header_candidates(lines, allow_relaxed=False):
        from_line = candidate.from_line
        to_line = candidate.to_line
        subject_line = candidate.subject_line
        if subject_norm and not _subject_line_matches(subject_line, subject_norm):
            continue
        if _is_empty_header_value(from_line):
            from_line = ""
        inferred_from = ""
        if not from_line:
            inferred_from = _infer_from_line_from_block(candidate.block_lines)

        effective_from = from_line or inferred_from
        missing_from = not effective_from
        allow_missing_from = False
        if subject_norm and missing_from:
            continue
        if missing_from and (require_non_ess or exclude_system):
            if _to_line_has_ess_dl(to_line) and subject_line and not _block_has_system_token(candidate.block_lines):
                allow_missing_from = True
            else:
                continue
        if subject_norm:
            if not effective_from:
                continue
            if ess_team is not None and _is_ess_from_line(effective_from, ess_team):
                continue
        if require_non_ess and ess_team is not None and not missing_from:
            if _is_ess_from_line(effective_from, ess_team):
                continue
        if exclude_system and not missing_from:
            if _is_system_from_line(effective_from):
                continue
        if exclude_system and missing_from and not allow_missing_from:
            continue
        dt = candidate.sent_dt
        if max_dt and _to_ist(dt) > _to_ist(max_dt):
            continue
        bounded_candidates.append(dt)
    if bounded_candidates:
        return max(bounded_candidates, key=_to_ist)

    # Look for blocks containing From/Sent/To/Subject
    block = []
    blocks = []
    has_any_from = any(_is_from_header_line(line or "") for line in lines)
    for line in lines:
        if _is_from_header_line(line):
            if block:
                blocks.append(block)
                block = []
            block.append(line)
            continue
        if block:
            block.append(line)
            if _is_subject_header_line(line):
                blocks.append(block)
                block = []
    if block:
        blocks.append(block)

    candidates = []
    for block in blocks:
        from_line = next((l for l in block if _is_from_header_line(l)), "")
        to_line = next((l for l in block if _is_to_header_line(l)), "")
        subject_line = next((l for l in block if _is_subject_header_line(l)), "")
        if subject_norm and not _subject_line_matches(subject_line, subject_norm):
            continue
        if _is_empty_header_value(from_line):
            from_line = ""
        inferred_from = ""
        if not from_line:
            inferred_from = _infer_from_line_from_block(block)

        effective_from = from_line or inferred_from
        missing_from = not effective_from
        allow_missing_from = False
        if subject_norm and missing_from:
            continue
        if missing_from and (require_non_ess or exclude_system):
            # Allow missing-From blocks only when they clearly target ESS and include a subject.
            if _to_line_has_ess_dl(to_line) and subject_line and not _block_has_system_token(block):
                allow_missing_from = True
            else:
                continue

        sent_line = next((l for l in block if _is_sent_header_line(l)), None)
        if not sent_line:
            continue

        if subject_norm:
            # If we have a subject anchor, require a non-ESS From line,
            # but do not require an explicit email address (names-only blocks exist).
            if not effective_from:
                continue
            if ess_team is not None and _is_ess_from_line(effective_from, ess_team):
                continue
        if require_non_ess and ess_team is not None and not missing_from:
            if _is_ess_from_line(effective_from, ess_team):
                continue
        if exclude_system and not missing_from:
            if _is_system_from_line(effective_from):
                continue
        if exclude_system and missing_from and not allow_missing_from:
            continue

        try:
            value = sent_line.split(":", 1)[1].strip()
        except Exception:
            continue
        dt = _parse_datetime(value)
        if not dt:
            continue
        if max_dt and _to_ist(dt) > _to_ist(max_dt):
            continue
        candidates.append(dt)

    if not candidates:
        return None
    return max(candidates, key=_to_ist)


def _extract_outlook_block_sent_candidates(
    lines,
    max_dt: datetime | None = None,
    ess_team=None,
    require_non_ess: bool = False,
    exclude_system: bool = False,
    subject_norm: str | None = None,
):
    bounded_candidates = []
    for candidate in _extract_bounded_outlook_header_candidates(lines, allow_relaxed=False):
        from_line = candidate.from_line
        to_line = candidate.to_line
        subject_line = candidate.subject_line
        if subject_norm and not _subject_line_matches(subject_line, subject_norm):
            continue
        if _is_empty_header_value(from_line):
            from_line = ""
        inferred_from = ""
        if not from_line:
            inferred_from = _infer_from_line_from_block(candidate.block_lines)

        effective_from = from_line or inferred_from
        missing_from = not effective_from
        allow_missing_from = False
        if subject_norm and missing_from:
            continue
        if missing_from and (require_non_ess or exclude_system):
            if _to_line_has_ess_dl(to_line) and subject_line and not _block_has_system_token(candidate.block_lines):
                allow_missing_from = True
            else:
                continue
        if subject_norm:
            if not effective_from:
                continue
            if ess_team is not None and _is_ess_from_line(effective_from, ess_team):
                continue
        if require_non_ess and ess_team is not None and not missing_from:
            if _is_ess_from_line(effective_from, ess_team):
                continue
        if exclude_system and not missing_from:
            if _is_system_from_line(effective_from):
                continue
        if exclude_system and missing_from and not allow_missing_from:
            continue
        dt = candidate.sent_dt
        if max_dt and _to_ist(dt) > _to_ist(max_dt):
            continue
        bounded_candidates.append(dt)
    if bounded_candidates:
        return bounded_candidates

    block = []
    blocks = []
    for line in lines:
        if _is_from_header_line(line):
            if block:
                blocks.append(block)
                block = []
            block.append(line)
            continue
        if block:
            block.append(line)
            if _is_subject_header_line(line):
                blocks.append(block)
                block = []
    if block:
        blocks.append(block)

    candidates = []
    for block in blocks:
        from_line = next((l for l in block if _is_from_header_line(l)), "")
        to_line = next((l for l in block if _is_to_header_line(l)), "")
        subject_line = next((l for l in block if _is_subject_header_line(l)), "")
        if subject_norm and not _subject_line_matches(subject_line, subject_norm):
            continue
        if _is_empty_header_value(from_line):
            from_line = ""
        inferred_from = ""
        if not from_line:
            inferred_from = _infer_from_line_from_block(block)

        effective_from = from_line or inferred_from
        missing_from = not effective_from
        allow_missing_from = False
        if subject_norm and missing_from:
            continue
        if missing_from and (require_non_ess or exclude_system):
            if _to_line_has_ess_dl(to_line) and subject_line and not _block_has_system_token(block):
                allow_missing_from = True
            else:
                continue

        sent_line = next((l for l in block if _is_sent_header_line(l)), None)
        if not sent_line:
            continue

        if subject_norm:
            if not effective_from:
                continue
            if ess_team is not None and _is_ess_from_line(effective_from, ess_team):
                continue
        if require_non_ess and ess_team is not None and not missing_from:
            if _is_ess_from_line(effective_from, ess_team):
                continue
        if exclude_system and not missing_from:
            if _is_system_from_line(effective_from):
                continue
        if exclude_system and missing_from and not allow_missing_from:
            continue

        try:
            value = sent_line.split(":", 1)[1].strip()
        except Exception:
            continue
        dt = _parse_datetime(value)
        if not dt:
            continue
        if max_dt and _to_ist(dt) > _to_ist(max_dt):
            continue
        candidates.append(dt)
    return candidates


def _extract_sent_from_context(
    lines,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    candidates = []
    for idx, line in enumerate(lines):
        lower = _normalize_header_line(line)
        if "sent:" not in lower and "envoy" not in lower:
            continue
        # Context window
        start = max(0, idx - 3)
        end = min(len(lines), idx + 4)
        context = lines[start:end]
        from_line = next((l for l in context if _is_from_header_line(l)), "")
        to_line = next((l for l in context if _is_to_header_line(l)), "")
        subject_line = next((l for l in context if _is_subject_header_line(l)), "")
        if subject_norm and not _subject_line_matches(subject_line, subject_norm):
            continue
        if _is_empty_header_value(from_line):
            from_line = ""

        if _block_has_system_token(context):
            continue

        inferred_from_line = ""
        if not from_line and idx > 0:
            prev = lines[idx - 1].strip()
            if prev and not prev.lower().startswith(("to:", "cc:", "bcc:", "subject:", "sent:", "from:")):
                inferred_from_line = f"From: {prev}"

        effective_from = from_line or inferred_from_line
        if subject_norm:
            if not effective_from:
                continue
            # Names-only From lines are allowed if they are non-ESS.
            if _is_ess_from_line(effective_from, ess_team):
                continue
        else:
            if effective_from:
                if _is_ess_from_line(effective_from, ess_team):
                    continue
            else:
                # If no From line at all, require ESS DL in To and a subject
                if not (_to_line_has_ess_dl(to_line) and subject_line):
                    continue

        try:
            value = line.split(":", 1)[1].strip()
        except Exception:
            continue
        dt = _parse_datetime(value)
        if not dt:
            continue
        if max_dt and _to_ist(dt) > _to_ist(max_dt):
            continue
        candidates.append(dt)

    if not candidates:
        return None
    return max(candidates, key=_to_ist)


def _extract_sent_from_context_candidates(
    lines,
    ess_team,
    max_dt: datetime | None = None,
    subject_norm: str | None = None,
):
    candidates = []
    for idx, line in enumerate(lines):
        lower = _normalize_header_line(line)
        if "sent:" not in lower and "envoy" not in lower:
            continue
        start = max(0, idx - 3)
        end = min(len(lines), idx + 4)
        context = lines[start:end]
        from_line = next((l for l in context if _is_from_header_line(l)), "")
        to_line = next((l for l in context if _is_to_header_line(l)), "")
        subject_line = next((l for l in context if _is_subject_header_line(l)), "")
        if subject_norm and not _subject_line_matches(subject_line, subject_norm):
            continue
        if _is_empty_header_value(from_line):
            from_line = ""

        if _block_has_system_token(context):
            continue

        inferred_from_line = ""
        if not from_line and idx > 0:
            prev = lines[idx - 1].strip()
            if prev and not prev.lower().startswith(("to:", "cc:", "bcc:", "subject:", "sent:", "from:")):
                inferred_from_line = f"From: {prev}"

        effective_from = from_line or inferred_from_line
        if subject_norm:
            if not effective_from:
                continue
            if _is_ess_from_line(effective_from, ess_team):
                continue
        else:
            if effective_from:
                if _is_ess_from_line(effective_from, ess_team):
                    continue
            else:
                if not (_to_line_has_ess_dl(to_line) and subject_line):
                    continue

        try:
            value = line.split(":", 1)[1].strip()
        except Exception:
            continue
        dt = _parse_datetime(value)
        if not dt:
            continue
        if max_dt and _to_ist(dt) > _to_ist(max_dt):
            continue
        candidates.append(dt)
    return candidates


def _extract_email_from_from_line(line: str) -> str:
    if not line:
        return ""
    try:
        value = line.split(":", 1)[1].strip()
    except Exception:
        value = line
    m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", value)
    return (m.group(0).lower() if m else "")


def _infer_from_line_from_block(block) -> str:
    # Try to infer a missing From line from the line above Sent
    for idx, line in enumerate(block):
        if (line or "").lower().startswith("sent:"):
            if idx > 0:
                prev = block[idx - 1].strip()
                if prev and not prev.lower().startswith(("to:", "cc:", "bcc:", "subject:", "sent:", "from:")):
                    return f"From: {prev}"
    return ""


def _is_empty_header_value(line: str) -> bool:
    if not line:
        return False
    if not _is_from_header_line(line):
        return False
    return line.split(":", 1)[1].strip() == ""


def _strip_ids_and_dates(text: str) -> str:
    if not text:
        return ""
    s = normalize_subject_for_match(text)
    s = re.sub(r"\binc\d{6,}\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bdr-\d{6,}-\d+\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b\d{4}[./-]\d{1,2}[./-]\d{1,2}\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(
        r"\b(?:jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\s+\d{4}\b",
        " ",
        s,
        flags=re.IGNORECASE,
    )
    s = re.sub(r"--?>", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _subject_line_matches(subject_line: str, subject_norm: str | None) -> bool:
    if not subject_norm:
        return True
    if not subject_line:
        return False
    try:
        value = subject_line.split(":", 1)[1].strip()
    except Exception:
        value = subject_line.strip()
    row_norm = normalize_subject(subject_norm)
    subj_norm = normalize_subject(value)
    if subj_norm and row_norm and subj_norm.lower() == row_norm.lower():
        return True

    row_alt = normalize_subject_for_match(subject_norm)
    subj_alt = normalize_subject_for_match(value)
    if subj_alt and row_alt and subj_alt.lower() == row_alt.lower():
        return True

    row_strip = _strip_ids_and_dates(subject_norm)
    subj_strip = _strip_ids_and_dates(value)
    if row_strip and subj_strip:
        if row_strip.lower() == subj_strip.lower():
            return True
        tokens_row = set(re.findall(r"[a-z0-9]+", row_strip.lower()))
        tokens_subj = set(re.findall(r"[a-z0-9]+", subj_strip.lower()))
        if tokens_row and tokens_subj:
            overlap = len(tokens_row & tokens_subj)
            if overlap / max(1, len(tokens_row)) >= 0.8 and overlap / max(1, len(tokens_subj)) >= 0.8:
                return True
        shorter, longer = sorted([row_strip, subj_strip], key=len)
        if len(shorter) >= 12 and shorter.lower() in longer.lower():
            return True

    return False


def _is_ess_from_line(line: str, ess_team) -> bool:
    if not line or not ess_team:
        return False
    if "enterprise-services-support" in line.lower():
        return True
    email = _extract_email_from_from_line(line)
    if email and email in ess_team:
        return True
    # No email found; try name token matching
    try:
        value = line.split(":", 1)[1]
    except Exception:
        value = line
    value = re.sub(r"mailto:", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", " ", value)
    value_tokens = _tokenize(value)
    tokens = set(value_tokens)
    last_token = value_tokens[-1] if value_tokens else ""
    if not tokens:
        return False
    for email in ess_team:
        local_tokens = set(_tokenize(_email_local(email)))
        if not local_tokens:
            continue
        overlap = tokens & local_tokens
        if len(overlap) >= 2:
            return True
        if overlap and last_token and last_token in local_tokens:
            return True
    return False


def _is_system_from_line(line: str) -> bool:
    if not line:
        return False
    value = line.lower()
    tokens = [
        "eai-system-notification",
        "system-notification",
        "system notification",
        "no-reply",
        "noreply",
        "do-not-reply",
        "donotreply",
    ]
    return any(t in value for t in tokens)


def _to_line_has_ess_dl(line: str) -> bool:
    if not line:
        return False
    value = line.lower()
    ess_dls = [
        "enterprise-services-support@umusic.com",
        "enterprise-services-support@inveniolsi.com",
        "enterprise-services-support@invenio-solutions.com",
    ]
    return any(dl in value for dl in ess_dls)


def _block_has_system_token(block) -> bool:
    tokens = [
        "eai-system-notification",
        "system-notification",
        "system notification",
        "no-reply",
        "noreply",
        "do-not-reply",
        "donotreply",
    ]
    for line in block:
        lower = (line or "").lower()
        if any(t in lower for t in tokens):
            return True
    return False


def _has_internal_marker(emails) -> bool:
    for e in emails:
        subj = (e.subject or "")
        body = (e.body or "")
        if INTERNAL_MARKER_RE.search(subj) or INTERNAL_MARKER_RE.search(body):
            return True
    return False



def _clean_quote_line(line: str) -> str:
    if not line:
        return ""
    # Strip HTML tags and decode entities (helps extract Sent/From/To in HTML bodies)
    line = re.sub(r"<[^>]+>", " ", line)
    line = html.unescape(line)
    # Remove common quote markers
    line = re.sub(r"^(>+\s*)", "", line)
    # Collapse whitespace
    line = " ".join(line.split())
    return line.strip()


def _body_to_clean_lines(body: str) -> list[str]:
    if not body:
        return []
    lines = []
    if re.search(r"(?is)<(html|body|div|p|br|table|tr|td|th|li|span)\b", body or ""):
        lines = _bs4_visible_lines_from_text(body)
    if not lines:
        lines = _fallback_visible_lines_from_text(body)
    return [_clean_quote_line(line) for line in lines if line.strip()]


def _parse_datetime(value: str):
    value = value.strip()
    fmts = [
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M",
        "%d/%m/%Y %H:%M",
        "%d %B %Y %I:%M %p",
        "%d-%b-%Y %I:%M %p",
        "%d-%b-%Y %H:%M",
        "%A, %B %d, %Y %I:%M %p",
        "%a, %b %d, %Y %I:%M %p",
        "%A, %d %B, %Y %H:%M",
        "%A, %d %B, %Y %I:%M %p",
        "%A, %d %B %Y %H:%M",
        "%A, %d %B %Y %I:%M %p",
        "%d %B %Y %H:%M",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            continue
    return None
