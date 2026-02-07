from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import html
import re
from zoneinfo import ZoneInfo


ACK_PHRASES = [
    "we will check",
    "update you",
    "let you know",
    "we will do the same",
    "we will do the same",
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
    return [t for t in _normalize_name(name).split() if t]


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


IST = ZoneInfo("Asia/Kolkata")


def _to_ist(dt: datetime) -> datetime:
    if not isinstance(dt, datetime):
        return datetime.min.replace(tzinfo=IST)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    return dt.astimezone(IST)


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


def resolve_times_with_debug(thread, requester_name, ess_team):
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

    # ESS-only thread with no non-ESS request (even quoted): all three same
    if not non_ess_emails:
        parsed_any = _find_latest_quoted_request_time(ordered, ess_team)
        if not parsed_any:
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
        parsed = _extract_request_time_from_body(only.body or "", ess_team, max_dt=only.sent_time)
        if not parsed:
            t = _format_time(only.sent_time)
            return (
                TimeResult(t, t, t),
                TimeDebug(only.sender_email, only.sender_email, only.sender_email, "Single ESS; no request found"),
            )

    # Failed/skip subjects with no ack phrase: handle safely without treating late ESS replies as ack
    if failed_subject and not has_ack_phrase:
        req_time, req_src = _latest_request_time_before(ordered, ess_team)
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
            if any(p in body for p in ACK_PHRASES):
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
        req_time, req_src = _latest_request_time_before(ordered, ess_team, resolved_mail.sent_time)
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
        req_time, req_src = _latest_request_time_before(ordered, ess_team)
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
        latest_req, latest_req_src = _latest_request_time_before(ordered, ess_team)
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
                ordered, ess_team, effective_ack.sent_time
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
            if not any(p in body for p in ACK_PHRASES):
                continue
            if created_time and _to_ist(e.sent_time) < _to_ist(created_time):
                continue
            req_time, req_src = _latest_request_time_before(ordered, ess_team, e.sent_time)
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
        req_time, req_src = _latest_request_time_before(ordered, ess_team, effective_ack.sent_time)
        if req_time:
            created_time = req_time
            created_src = req_src
        else:
            parsed = _extract_request_time_from_body(
                effective_ack.body or "", ess_team, max_dt=effective_ack.sent_time
            )
            if parsed:
                created_time = parsed
                created_src = "PARSED_FROM_BODY"
                if ack_mail is None:
                    notes = "Ack fallback; created parsed from body"
                else:
                    notes = "Ack found; created parsed from body"
            else:
                parsed_any = _find_latest_quoted_request_time(ordered, ess_team, max_dt=effective_ack.sent_time)
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
            effective_ack.body or "", ess_team, max_dt=effective_ack.sent_time
        )
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
                if not any(p in body for p in ACK_PHRASES):
                    continue
                if _to_ist(e.sent_time) > ack_dt:
                    break
                req_time, req_src = _latest_request_time_before(ordered, ess_team, e.sent_time)
                if req_time:
                    effective_ack = e
                    ack_dt = _to_ist(e.sent_time)
                    created_time = req_time
                    created_src = req_src
                    created_dt = _to_ist(created_time)
                    notes = "Ack phrase preferred (created<=ack)"
                    break
            # Ensure created time is not after ack; recompute using ack time as cap
            req_time, req_src = _latest_request_time_before(ordered, ess_team, effective_ack.sent_time)
            if req_time:
                created_time = req_time
                created_src = req_src
                created_dt = _to_ist(created_time)
        response_time = _format_time(effective_ack.sent_time)
        if ack_dt - created_dt <= timedelta(minutes=17):
            if ack_mail is None:
                notes = "Ack fallback OK"
            else:
                notes = "OK"
        else:
            notes = "Ack delayed (>17 min)"

        # If created_time equals ack_time, try to parse request time from ack body
        if created_dt == ack_dt:
            parsed_from_ack = _extract_request_time_from_body(
                effective_ack.body or "", ess_team, max_dt=effective_ack.sent_time
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
                    resolved_mail.body or "", ess_team, max_dt=resolved_mail.sent_time
                )
                if parsed:
                    created_time = parsed
                    created_src = "PARSED_FROM_BODY"
                    notes = "Ack missing; created parsed from body"
                else:
                    parsed_any = _find_latest_quoted_request_time(ordered, ess_team, max_dt=resolved_mail.sent_time)
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
                parsed = _extract_request_time_from_body((ordered[-1].body or ""), ess_team)
                if parsed:
                    created_time = parsed
                    created_src = "PARSED_FROM_BODY"
                    notes = "Ack missing; created parsed from body"
                else:
                    parsed_any = _find_latest_quoted_request_time(ordered, ess_team, max_dt=ordered[-1].sent_time)
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

    if not created_time:
        # ESS-only threads with no parsed request: fallback to first email time
        if non_ess_emails == [] and ess_emails:
            parsed_any = _find_latest_quoted_request_time(ordered, ess_team, max_dt=first.sent_time)
            if parsed_any:
                created_time = parsed_any
            else:
                if not has_ack_phrase and not has_internal_marker:
                    t = _format_time(first.sent_time)
                    return (
                        TimeResult(t, t, t),
                        TimeDebug(first.sender_email, first.sender_email, first.sender_email, "ESS only; no request found"),
                    )
        return (
            TimeResult("", "NA", _format_time(resolved_mail.sent_time) if resolved_mail else ""),
            TimeDebug(created_src, ack_src, resolved_mail.sender_email if resolved_mail else "", notes),
        )

    return (
        TimeResult(
            _format_time(created_time),
            response_time,
            _format_time(resolved_mail.sent_time) if resolved_mail else (_format_time(created_time) if created_time else ""),
        ),
        TimeDebug(
            created_src or (created_mail.sender_email if created_mail else ""),
            ack_src,
            resolved_mail.sender_email if resolved_mail else (created_src or ""),
            notes if resolved_mail else notes + "; No requester match",
        ),
    )


def _has_ack_phrase(ordered, ess_team) -> bool:
    for e in ordered:
        if _is_ess_sender(e, ess_team):
            body = (e.body or "").lower()
            if any(p in body for p in ACK_PHRASES):
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


def _find_latest_quoted_request_time(emails, ess_team, max_dt: datetime | None = None):
    latest = None
    for e in emails:
        parsed = _extract_request_time_from_body(e.body or "", ess_team, max_dt=max_dt or e.sent_time)
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
        has_phrase = any(p in body for p in ACK_PHRASES)
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
        has_phrase = any(p in body for p in ACK_PHRASES)
        candidates.append((has_phrase, e))

    if not candidates:
        return None

    # Prefer any ESS reply within 17 minutes of the request (even without ack phrase).
    window = timedelta(minutes=17)
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


def _latest_request_time_before(emails, ess_team, max_dt: datetime | None = None):
    latest_time = None
    latest_src = ""

    # Non-ESS emails
    for e in emails:
        if _is_ess_sender(e, ess_team):
            continue
        if max_dt and _to_ist(e.sent_time) > _to_ist(max_dt):
            continue
        if latest_time is None or _to_ist(e.sent_time) > _to_ist(latest_time):
            latest_time = e.sent_time
            latest_src = e.sender_email or e.sender_name

    # Quoted request times in ESS emails (or any email bodies)
    for e in emails:
        parsed = _extract_request_time_from_body(e.body or "", ess_team, max_dt=max_dt or e.sent_time)
        if not parsed:
            continue
        if max_dt and _to_ist(parsed) > _to_ist(max_dt):
            continue
        if latest_time is None or _to_ist(parsed) > _to_ist(latest_time):
            latest_time = parsed
            latest_src = "PARSED_FROM_BODY"

    return latest_time, latest_src


def _latest_request_time_before_internal(emails, ess_team, requester_name: str, max_dt: datetime | None = None):
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
        parsed = _extract_request_time_from_body(e.body or "", ess_team, max_dt=max_dt or e.sent_time)
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
    sent_lines = [l for l in lines if l.lower().startswith("sent:")]
    if not sent_lines:
        # common Outlook block: "Sent: Tuesday, January 2, 2026 3:45 PM"
        sent_lines = [l for l in lines if "sent:" in l.lower()]

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


def _extract_request_time_from_body(body: str, ess_team, max_dt: datetime | None = None):
    if not body:
        return None
    lines = [_clean_quote_line(line) for line in body.splitlines() if line.strip()]
    block_dt = _extract_outlook_block_sent(
        lines,
        max_dt,
        ess_team=ess_team,
        require_non_ess=True,
        exclude_system=True,
    )
    if block_dt:
        return block_dt
    return _extract_sent_from_context(lines, ess_team, max_dt)


def _extract_outlook_block_sent(
    lines,
    max_dt: datetime | None = None,
    ess_team=None,
    require_non_ess: bool = False,
    exclude_system: bool = False,
):
    # Look for blocks containing From/Sent/To/Subject
    block = []
    blocks = []
    has_any_from = any((line or "").lower().startswith("from:") for line in lines)
    for line in lines:
        if line.lower().startswith("from:"):
            if block:
                blocks.append(block)
                block = []
            block.append(line)
            continue
        if block:
            block.append(line)
            if line.lower().startswith("subject:"):
                blocks.append(block)
                block = []
    if block:
        blocks.append(block)

    candidates = []
    for block in blocks:
        from_line = next((l for l in block if l.lower().startswith("from:")), "")
        to_line = next((l for l in block if l.lower().startswith("to:")), "")
        subject_line = next((l for l in block if l.lower().startswith("subject:")), "")
        if _is_empty_header_value(from_line):
            from_line = ""
        inferred_from = ""
        if not from_line:
            inferred_from = _infer_from_line_from_block(block)

        effective_from = from_line or inferred_from
        missing_from = not effective_from
        allow_missing_from = False
        if missing_from and (require_non_ess or exclude_system):
            # Allow missing-From blocks only when they clearly target ESS and include a subject.
            if _to_line_has_ess_dl(to_line) and subject_line and not _block_has_system_token(block):
                allow_missing_from = True
            else:
                continue

        sent_line = next((l for l in block if l.lower().startswith("sent:")), None)
        if not sent_line:
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


def _extract_sent_from_context(lines, ess_team, max_dt: datetime | None = None):
    candidates = []
    for idx, line in enumerate(lines):
        lower = line.lower()
        if "sent:" not in lower:
            continue
        # Context window
        start = max(0, idx - 3)
        end = min(len(lines), idx + 4)
        context = lines[start:end]
        from_line = next((l for l in context if l.lower().startswith("from:")), "")
        to_line = next((l for l in context if l.lower().startswith("to:")), "")
        subject_line = next((l for l in context if l.lower().startswith("subject:")), "")
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
    if not line.lower().startswith("from:"):
        return False
    return line.split(":", 1)[1].strip() == ""


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
        "%A, %d %B %Y %H:%M",
        "%d %B %Y %H:%M",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            continue
    return None
