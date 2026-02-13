from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import html
import re
from zoneinfo import ZoneInfo
import unicodedata

from src.rules.subject_normalizer import normalize_subject, normalize_subject_for_match


ACK_PHRASES = [
    "sure, we will process the file to",
    "we will check",
    "update you",
    "let you know",
    "we will process the file to",
    "we will do the same",
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
    "could you update us on the below",
    "please provide us an update on the below",
    "please provide us an update",
    "please provide an update on the below",
    "please provide an update",
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

    if len(overlap) / max(1, len(r_tokens)) >= 0.6:
        return True
    if len(overlap) / max(1, len(s_tokens)) >= 0.6:
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
    body = (email_record.body or "").lower()
    return (
        "thanks for the information" in body
        or "thanks for the confirmation" in body
        or "we will ignore" in body
        or "could you please provide us an update on the below" in body
        or "could you please provide us an update regarding the below" in body
        or "could you update us on the below" in body
    )


def _is_ack_body(body: str) -> bool:
    if not body:
        return False
    b = body.lower()
    if any(p in b for p in NON_ACK_PHRASES):
        return False
    return any(p in b for p in ACK_PHRASES)


def _is_ack_like_reply(email_record) -> bool:
    """
    Ack-like replies should not be treated as resolved-time candidates.
    Keep direct-resolution mails out of this bucket.
    """
    body = (email_record.body or "").lower()
    if not body:
        return False
    if any(p in body for p in DIRECT_RESOLUTION_PHRASES):
        return False
    return _is_ack_body(body)


IST = ZoneInfo("Asia/Kolkata")


def _to_ist(dt: datetime) -> datetime:
    if not isinstance(dt, datetime):
        return datetime.min.replace(tzinfo=IST)
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


def _format_time(dt: datetime) -> str:
    if not isinstance(dt, datetime):
        return ""
    if dt.year <= 1901:
        return ""
    try:
        dt = _to_ist(dt)
    except Exception:
        pass
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
    resolved_mail = requester_emails[-1] if requester_emails else None
    if resolved_mail and _is_thanks_info_reply(resolved_mail):
        for e in reversed(requester_emails):
            if not _is_thanks_info_reply(e):
                resolved_mail = e
                break
    if failed_subject and requester_emails:
        for e in reversed(requester_emails):
            if not _is_thanks_info_reply(e):
                resolved_mail = e
                break
    requester_first = requester_emails[0] if requester_emails else None
    requester_is_ess = bool(requester_candidates) or any(
        _is_ess_sender(e, ess_team) for e in requester_emails
    )

    # Force PROD subjects: all three times are the first email time
    for e in ordered:
        subj = (e.subject or "").lower()
        if any(p in subj for p in FORCE_PROD_SAME_TIME_PHRASES):
            t = _format_time(first.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(
                    first.sender_email,
                    first.sender_email,
                    first.sender_email,
                    "Force PROD subject; all times same",
                ),
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
        parsed_any = _find_latest_quoted_request_time(
            ordered,
            ess_team,
            max_dt=first.sent_time,
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
                    if not body:
                        continue
                    if subj_alt.lower() not in body.lower():
                        continue
                    parsed = _extract_request_time_from_body(
                        body,
                        ess_team,
                        max_dt=first.sent_time,
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
                created_t = _format_time(ess_emails[0].sent_time)
                non_ack_ess = [e for e in ess_emails if not _is_ack_like_reply(e)]
                last_ess = max(non_ack_ess, key=lambda e: e.sent_time) if non_ack_ess else max(ess_emails, key=lambda e: e.sent_time)

                # Prefer requester-owned reply for resolved in ESS-only span threads.
                requester_non_ack = [e for e in requester_emails if not _is_ack_like_reply(e)]
                if requester_non_ack:
                    resolved_mail_pick = max(requester_non_ack, key=lambda e: e.sent_time)
                    span_note = "ESS-only; no non-ESS request; requester span"
                elif requester_emails:
                    # Keep consultant ownership stable even when requester replies
                    # are mostly ack-like.
                    resolved_mail_pick = max(requester_emails, key=lambda e: e.sent_time)
                    span_note = "ESS-only; no non-ESS request; requester span(ack-like)"
                else:
                    resolved_mail_pick = last_ess
                    span_note = "ESS-only; no non-ESS request; span"

                resolved_t = _format_time(resolved_mail_pick.sent_time) if resolved_mail_pick else created_t

                # Try to use an ack phrase if it exists; otherwise keep ACK NOT FOUND.
                ack_mail_pick = None
                for e in ess_emails:
                    body = (e.body or "").lower()
                    if _is_ack_body(body):
                        ack_mail_pick = e
                        break
                if ack_mail_pick:
                    ack_t = _format_time(ack_mail_pick.sent_time)
                    ack_src = ack_mail_pick.sender_email or ack_mail_pick.sender_name
                else:
                    ack_t = resolved_t
                    ack_src = "ACK NOT FOUND"
                return (
                    TimeResult(created_t, ack_t, resolved_t),
                    TimeDebug(
                        ess_emails[0].sender_email or ess_emails[0].sender_name,
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
        parsed = _extract_request_time_from_body(
            only.body or "",
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
        # No direct-resolution reply: if we have a request time, set all three same
        if req_time:
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
            parsed = _extract_request_time_from_body(
                effective_ack.body or "",
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
        parsed_from_ack = _extract_request_time_from_body(
            effective_ack.body or "",
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
        response_time = _format_time(effective_ack.sent_time)
        if ack_dt - created_dt <= timedelta(minutes=16):
            if ack_mail is None:
                notes = "Ack fallback OK"
            else:
                notes = "OK"
        else:
            notes = "Ack delayed (>16 min)"

        # If created came from quoted parsing, re-parse ack body to find the
        # closest request immediately below the ack (common Outlook chain).
        if created_src.startswith("PARSED_FROM_BODY"):
            parsed_from_ack = _extract_request_time_from_body(
                effective_ack.body or "",
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
                        # Recompute note based on new delta
                        if ack_dt - created_dt <= timedelta(minutes=16):
                            notes = "OK"
                        else:
                            notes = "Ack delayed (>16 min)"

        # If created_time equals ack_time, try to parse request time from ack body
        if created_dt == ack_dt:
            parsed_from_ack = _extract_request_time_from_body(
                effective_ack.body or "",
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
                parsed = _extract_request_time_from_body(
                    resolved_mail.body or "",
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
                parsed = _extract_request_time_from_body(
                    (ordered[-1].body or ""),
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
                response_time = "NA"
                notes = "No ack; requester reply >20 min"
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
                if _is_ack_like_reply(e):
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
            parsed_any = _find_latest_quoted_request_time(
                ordered,
                ess_team,
                max_dt=first.sent_time,
                subject_norm=subject_norm,
            )
            if parsed_any:
                created_time = parsed_any
            else:
                if not has_ack_phrase and not has_internal_marker:
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

    return (
        TimeResult(
            _format_time(created_time),
            response_time,
            resolved_time,
        ),
        TimeDebug(
            created_src or (created_mail.sender_email if created_mail else ""),
            ack_src,
            resolved_mail.sender_email if resolved_mail else (created_src or ""),
            notes,
        ),
    )


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
        parsed = _extract_request_time_from_body(
            e.body or "",
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


def _select_ack_and_request(emails, ess_team, max_time: datetime | None = None):
    candidates = []
    for e in emails:
        if not _is_ess_sender(e, ess_team):
            continue
        if max_time and _to_ist(e.sent_time) > _to_ist(max_time):
            continue
        req_time, req_src = _latest_request_time_before(emails, ess_team, e.sent_time)
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

    # If we already saw a real non-ESS request, do not let quoted blocks override it.
    if not found_non_ess:
        # Quoted request times in ESS emails (or any email bodies)
        for e in emails:
            parsed = _extract_request_time_from_body(
                e.body or "",
                ess_team,
                max_dt=max_dt or e.sent_time,
                subject_norm=subject_norm,
            )
            if not parsed:
                continue
            if max_dt and _to_ist(parsed) > _to_ist(max_dt):
                continue
            if latest_time is None or _to_ist(parsed) > _to_ist(latest_time):
                latest_time = parsed
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

    # Non-ESS emails OR internal requester emails (even if ESS)
    for e in emails:
        if _is_ess_sender(e, ess_team) and not _match_requester(e.sender_name, e.sender_email, requester_name):
            continue
        if max_dt and _to_ist(e.sent_time) > _to_ist(max_dt):
            continue
        if latest_time is None or _to_ist(e.sent_time) > _to_ist(latest_time):
            latest_time = e.sent_time
            latest_src = e.sender_email or e.sender_name

    # Quoted request times in ESS emails (or any email bodies)
    for e in emails:
        parsed = _extract_request_time_from_body(
            e.body or "",
            ess_team,
            max_dt=max_dt or e.sent_time,
            subject_norm=subject_norm,
        )
        if not parsed:
            continue
        if max_dt and _to_ist(parsed) > _to_ist(max_dt):
            continue
        if latest_time is None or _to_ist(parsed) > _to_ist(latest_time):
            latest_time = parsed
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

    name_tokens = set(_tokenize(sender_name))
    if not name_tokens:
        return False

    for email in ess_team:
        local_tokens = set(_tokenize(_email_local(email)))
        if not local_tokens:
            continue
        overlap = name_tokens & local_tokens
        if len(overlap) >= 2:
            return True
        if overlap and list(name_tokens)[-1] in local_tokens:
            return True
    return False


def _is_system_sender(email_record) -> bool:
    sender_email = (email_record.sender_email or "").lower()
    sender_name = (email_record.sender_name or "").lower()
    tokens = ["system-notification", "system notification", "no-reply", "noreply", "do-not-reply", "donotreply"]
    return any(t in sender_email for t in tokens) or any(t in sender_name for t in tokens)


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
    lines = [_clean_quote_line(line) for line in body.splitlines() if line.strip()]
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


def _extract_outlook_block_sent(
    lines,
    max_dt: datetime | None = None,
    ess_team=None,
    require_non_ess: bool = False,
    exclude_system: bool = False,
    subject_norm: str | None = None,
):
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


def _extract_email_from_from_line(line: str) -> str:
    if not line:
        return ""
    try:
        value = line.split(":", 1)[1].strip()
    except Exception:
        value = line
    m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}", value)
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
    value = re.sub(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}", " ", value)
    tokens = set(_tokenize(value))
    if not tokens:
        return False
    for email in ess_team:
        local_tokens = set(_tokenize(_email_local(email)))
        if not local_tokens:
            continue
        overlap = tokens & local_tokens
        if len(overlap) >= 2:
            return True
        if overlap and list(tokens)[-1] in local_tokens:
            return True
    return False


def _is_system_from_line(line: str) -> bool:
    if not line:
        return False
    value = line.lower()
    tokens = ["system-notification", "system notification", "no-reply", "noreply", "do-not-reply", "donotreply"]
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
    tokens = ["system-notification", "system notification", "no-reply", "noreply", "do-not-reply", "donotreply"]
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
    line = re.sub(r"^(>+\\s*)", "", line)
    # Collapse whitespace
    line = " ".join(line.split())
    return line.strip()


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
