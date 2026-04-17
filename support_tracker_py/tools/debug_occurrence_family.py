import argparse
import csv
import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Iterable

from src.rules.subject_normalizer import extract_subject_from_description, normalize_subject
from src.rules.time_resolver import (
    _classify_reply_kind,
    _is_ess_sender,
    _match_requester,
    _to_ist,
)


@dataclass
class DebugEmail:
    subject: str
    sender_name: str
    sender_email: str
    sent_time: datetime | None
    body: str
    body_html: str
    path: Path


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


def _parse_datetime(value: str) -> datetime | None:
    if not value:
        return None
    for fmt in (
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return datetime.strptime(value.strip(), fmt)
        except Exception:
            continue
    return None


def _fmt(dt: datetime | None) -> str:
    if not dt:
        return ""
    ist = _to_ist(dt)
    return ist.strftime("%d-%m-%Y %H:%M") if ist else ""


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


def _subject_match(subject_norm: str, email_subject: str) -> bool:
    e_norm = normalize_subject(email_subject or "")
    if not subject_norm or not e_norm:
        return False
    if subject_norm == e_norm:
        return True
    return subject_norm in e_norm or e_norm in subject_norm


def _cluster_family_rows(rows: list[dict]) -> dict[str, list[dict]]:
    families: dict[str, list[dict]] = {}
    for row in rows:
        row_norm = _family_subject_norm(row)
        if not row_norm:
            continue
        chosen_key = None
        for family_key in families:
            if _subject_match(family_key, row_norm) or _subject_match(row_norm, family_key):
                chosen_key = family_key
                break
        if chosen_key is None:
            chosen_key = row_norm
            families[chosen_key] = []
        families[chosen_key].append(row)
    return families


def _minute_dedupe(emails: list[DebugEmail]) -> list[DebugEmail]:
    seen = set()
    out = []
    for email in sorted(emails, key=lambda x: x.sent_time or datetime.max):
        if not email.sent_time:
            continue
        minute_key = _to_ist(email.sent_time).replace(second=0, microsecond=0)
        if minute_key in seen:
            continue
        seen.add(minute_key)
        out.append(email)
    return out


def _print_row_block(title: str, rows: list[dict]) -> None:
    print(title)
    if not rows:
        print("  <none>")
        return
    for row in rows:
        desc = _get_col(row, "Description")
        req = _get_col(row, "Requester", "Consultant")
        created = _get_col(row, "Created Date & Time")
        ack = _get_col(row, "Actual Response Date & Time")
        resolved = _get_col(row, "Actual Resolved Date & Time")
        notes = _get_col(row, "Notes")
        line = row.get("_line", "")
        print(f"  line={line} | requester={req} | created={created} | ack={ack} | resolved={resolved}")
        print(f"    desc={desc}")
        if notes:
            print(f"    notes={notes}")


def _requester_match_any(email: DebugEmail, requesters: list[str]) -> bool:
    return any(_match_requester(email.sender_name, email.sender_email, req) for req in requesters if req)


def _print_pool(title: str, pool: list[DebugEmail], ess_team: list[str], requesters: list[str]) -> None:
    print(title)
    if not pool:
        print("  <empty>")
        return
    for idx, email in enumerate(pool, start=1):
        cls = _classify_reply_kind(email)
        kind = cls.get("kind")
        is_ess = _is_ess_sender(email, ess_team)
        req_match = _requester_match_any(email, requesters)
        print(
            f"  {idx}. {_fmt(email.sent_time)} | kind={kind} | ess={is_ess} | requester_match={req_match} | "
            f"from={email.sender_email or email.sender_name} | subj={email.subject}"
        )


def _fmt_triplet(row: dict) -> str:
    created = _get_col(row, "Created Date & Time")
    ack = _get_col(row, "Actual Response Date & Time")
    resolved = _get_col(row, "Actual Resolved Date & Time")
    return f"{created or '-'} / {ack or '-'} / {resolved or '-'}"


def _row_is_all_ack_to_ess(row: dict) -> bool:
    notes = _get_col(row, "Notes").lower()
    return "requester span(all-ack->ess)" in notes and "ess-only; no non-ess request" in notes


def _row_is_occurrence_managed(row: dict) -> bool:
    notes = _get_col(row, "Notes").lower()
    return any(
        token in notes
        for token in (
            "dateanchoroccurrence",
            "ess-only; no non-ess request",
            "requester follow-up",
            "esscontinuationguard[",
            "quotedrequestonlynopair",
        )
    )


def _row_shape_hint(row: dict) -> tuple[str, str]:
    notes_l = _get_col(row, "Notes").lower()
    created_src_l = _get_col(row, "CreatedSource", "Created Source").lower()
    created = _parse_datetime(_get_col(row, "Created Date & Time"))
    ack = _parse_datetime(_get_col(row, "Actual Response Date & Time"))
    resolved = _parse_datetime(_get_col(row, "Actual Resolved Date & Time"))

    if "lanelocalinitialepisode[all-three-same]" in notes_l:
        return "all_three_same", "seeded_all_three_same"
    if "esscontinuationguard[allthreestrictessonly]" in notes_l:
        return "all_three_same", "legacy_occurrence_same_time"
    if "occurrenceslotshape[allthreesame]" in notes_l:
        return "all_three_same", "slot_shape_same_time"
    if created and ack and resolved and created == ack == resolved:
        if "ess-only; no non-ess request" in notes_l or "ess initiated; no ack; no consultant reply after request" in notes_l:
            return "all_three_same", "existing_same_time_triplet"
    if "lanelocalinitialepisode[req-ack-reply]" in notes_l:
        return "hybrid", "seeded_req_ack_reply"
    if "lanelocalinitialepisode[direct-reply]" in notes_l or created_src_l == "parsed_from_quoted_request":
        return "triplet", "seeded_direct_reply"
    return "", ""


def _family_slot_strategy(rows: list[dict]) -> str:
    requesters = []
    for row in rows:
        req = _get_col(row, "Requester", "Consultant")
        if req:
            requesters.append(req.strip().lower())
    distinct = {req for req in requesters if req}
    if len(distinct) <= 1:
        return "same_requester_chronological"
    return "latest_live_reply_ownership"


def _substantive_requester_pool(matched_emails: list[DebugEmail], requesters: list[str]) -> list[DebugEmail]:
    substantive = []
    acky = []
    fallback = []
    for email in matched_emails:
        if not email.sent_time:
            continue
        if not _requester_match_any(email, requesters):
            continue
        cls = _classify_reply_kind(email)
        if cls.get("thanks_info") or cls.get("nonfinal_followup"):
            continue
        if cls.get("real_reply") or cls.get("direct_resolution"):
            substantive.append(email)
        elif cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack"):
            acky.append(email)
        else:
            fallback.append(email)
    pool = substantive or acky or fallback
    return _minute_dedupe(pool)


def _latest_owner_reply_for_row(row: dict, matched_emails: list[DebugEmail]) -> DebugEmail | None:
    requester = _get_col(row, "Requester", "Consultant")
    if not requester:
        return None
    pool = _substantive_requester_pool(matched_emails, [requester])
    if not pool:
        return None
    return max(
        pool,
        key=lambda email: (_to_ist(email.sent_time).timestamp() if _to_ist(email.sent_time) else float("-inf")),
    )


def _row_anchor_ist(row: dict, output_by_line: dict[int, dict]) -> datetime | None:
    try:
        line = int(row.get("_line", 0))
    except Exception:
        return None
    out_row = output_by_line.get(line) or {}
    for col in ("Actual Resolved Date & Time", "Actual Response Date & Time", "Created Date & Time"):
        parsed = _parse_datetime(_get_col(out_row, col))
        if parsed:
            return _to_ist(parsed)
    return None


def _family_anchor_ists(family_rows: list[dict], output_by_line: dict[int, dict]) -> list[datetime]:
    anchors = []
    seen = set()
    for row in family_rows:
        anchor = _row_anchor_ist(row, output_by_line)
        if anchor is None:
            continue
        minute_anchor = anchor.replace(second=0, microsecond=0)
        if minute_anchor in seen:
            continue
        seen.add(minute_anchor)
        anchors.append(anchor)
    anchors.sort()
    return anchors


def _fmt_time_list(times: list[datetime | None]) -> str:
    rendered = [_fmt(item) for item in times if item is not None]
    return ", ".join(rendered) if rendered else "-"


def _distinct_requesters(rows: list[dict]) -> list[str]:
    requesters = []
    seen = set()
    for row in rows:
        req = _get_col(row, "Requester", "Consultant")
        req_key = req.strip().lower()
        if req and req_key not in seen:
            seen.add(req_key)
            requesters.append(req)
    return requesters


def _select_pool_window(
    full_pool: list[DebugEmail],
    family_anchor_ists: list[datetime],
    rows_needed: int,
) -> tuple[list[DebugEmail], str, list[datetime | None]]:
    family_start = family_anchor_ists[0] if family_anchor_ists else None
    family_end = family_anchor_ists[-1] if family_anchor_ists else None
    local_pool = []
    if family_start and family_end:
        lower = family_start - timedelta(hours=48)
        upper = family_end + timedelta(hours=24)
        local_pool = [
            email
            for email in full_pool
            if email.sent_time and lower <= (_to_ist(email.sent_time) or lower) <= upper
        ]

    if local_pool:
        pool = local_pool
        label = "family_anchor_window" if len(local_pool) >= rows_needed else "family_anchor_window_partial"
    elif len(full_pool) >= rows_needed:
        pool = full_pool[-rows_needed:]
        label = "latest_n_requester_pool"
    else:
        pool = full_pool
        label = "full_requester_pool"

    pool_times = [_to_ist(email.sent_time) if email.sent_time else None for email in pool]
    return pool, label, pool_times


def _assign_requester_rows(
    requester_rows: list[dict],
    matched_emails: list[DebugEmail],
    output_by_line: dict[int, dict],
    family_anchor_ists: list[datetime],
) -> list[dict]:
    sorted_rows = sorted(requester_rows, key=lambda r: int(r.get("_line", 10**9)))
    requester = _get_col(sorted_rows[0], "Requester", "Consultant") if sorted_rows else ""
    full_pool = _substantive_requester_pool(matched_emails, [requester]) if requester else []
    if not full_pool:
        return [
            {
                "row": row,
                "owner_reply": None,
                "owner_reply_ist": None,
                "row_anchor_ist": _row_anchor_ist(row, output_by_line),
                "assign_mode": "no_requester_pool",
                "candidate_pool_label": "empty",
                "candidate_pool_times": [],
            }
            for row in sorted_rows
        ]

    pool, pool_label, pool_time_list = _select_pool_window(full_pool, family_anchor_ists, len(sorted_rows))

    if len(sorted_rows) == 1:
        owner_reply = pool[-1]
        owner_reply_ist = _to_ist(owner_reply.sent_time) if owner_reply and owner_reply.sent_time else None
        return [
            {
                "row": sorted_rows[0],
                "owner_reply": owner_reply,
                "owner_reply_ist": owner_reply_ist,
                "row_anchor_ist": _row_anchor_ist(sorted_rows[0], output_by_line),
                "assign_mode": "single_latest",
                "candidate_pool_label": pool_label,
                "candidate_pool_times": pool_time_list,
            }
        ]

    anchor_rows = []
    for row in sorted_rows:
        anchor_ist = _row_anchor_ist(row, output_by_line)
        anchor_rows.append((row, anchor_ist))

    distinct_anchors = {
        anchor.replace(second=0, microsecond=0)
        for _row, anchor in anchor_rows
        if anchor is not None
    }
    use_anchor_matching = len(distinct_anchors) >= len(sorted_rows)

    assignments = []
    available = list(pool)
    if use_anchor_matching and len(pool) >= len(sorted_rows):
        for row, anchor_ist in sorted(anchor_rows, key=lambda item: (item[1] is None, item[1] or datetime.max)):
            if not available:
                assignments.append(
                    {
                        "row": row,
                        "owner_reply": None,
                        "owner_reply_ist": None,
                        "row_anchor_ist": anchor_ist,
                        "assign_mode": "anchor_match_exhausted",
                        "candidate_pool_label": pool_label,
                        "candidate_pool_times": pool_time_list,
                    }
                )
                continue
            if anchor_ist is None:
                chosen = available.pop(0)
            else:
                chosen_idx = min(
                    range(len(available)),
                    key=lambda idx: abs(((_to_ist(available[idx].sent_time) or anchor_ist) - anchor_ist).total_seconds()),
                )
                chosen = available.pop(chosen_idx)
            assignments.append(
                {
                    "row": row,
                    "owner_reply": chosen,
                    "owner_reply_ist": _to_ist(chosen.sent_time) if chosen and chosen.sent_time else None,
                    "row_anchor_ist": anchor_ist,
                    "assign_mode": "anchor_match_unique",
                    "candidate_pool_label": pool_label,
                    "candidate_pool_times": pool_time_list,
                }
            )
        assignments.sort(key=lambda item: int(item["row"].get("_line", 10**9)))
        return assignments

    if pool_label.startswith("family_anchor_window"):
        # When anchors collapse to the same local minute, reuse the local cluster instead of
        # forcing stale historical replies just to make rows unique.
        for row, anchor_ist in anchor_rows:
            if anchor_ist is None:
                chosen = pool[-1]
            else:
                chosen = min(
                    pool,
                    key=lambda email: abs(((_to_ist(email.sent_time) or anchor_ist) - anchor_ist).total_seconds()),
                )
            assignments.append(
                {
                    "row": row,
                    "owner_reply": chosen,
                    "owner_reply_ist": _to_ist(chosen.sent_time) if chosen and chosen.sent_time else None,
                    "row_anchor_ist": anchor_ist,
                    "assign_mode": "anchor_match_reuse_local",
                    "candidate_pool_label": pool_label,
                    "candidate_pool_times": pool_time_list,
                }
            )
        return assignments

    # Same-requester repeated rows: keep lanes distinct, but prefer recent requester history only.
    for idx, (row, anchor_ist) in enumerate(anchor_rows):
        chosen = pool[idx] if idx < len(pool) else pool[-1]
        assignments.append(
            {
                "row": row,
                "owner_reply": chosen,
                "owner_reply_ist": _to_ist(chosen.sent_time) if chosen and chosen.sent_time else None,
                "row_anchor_ist": anchor_ist,
                "assign_mode": "chronological_selected_pool",
                "candidate_pool_label": pool_label,
                "candidate_pool_times": pool_time_list,
            }
        )
    return assignments


def _assign_subject_wide_rows(
    special_rows: list[dict],
    matched_emails: list[DebugEmail],
    output_by_line: dict[int, dict],
) -> list[dict]:
    sorted_rows = sorted(special_rows, key=lambda r: int(r.get("_line", 10**9)))
    requesters = _distinct_requesters(sorted_rows)
    full_pool = _substantive_requester_pool(matched_emails, requesters)
    family_anchor_ists = _family_anchor_ists(sorted_rows, output_by_line)
    if not full_pool:
        return [
            {
                "row": row,
                "owner_reply": None,
                "owner_reply_ist": None,
                "row_anchor_ist": _row_anchor_ist(row, output_by_line),
                "assign_mode": "subject_wide_no_pool",
                "candidate_pool_label": "empty",
                "candidate_pool_times": [],
                "subject_wide_special": True,
            }
            for row in sorted_rows
        ]

    pool, pool_label, pool_time_list = _select_pool_window(full_pool, family_anchor_ists, len(sorted_rows))
    anchor_rows = []
    for row in sorted_rows:
        anchor_rows.append((row, _row_anchor_ist(row, output_by_line)))

    distinct_anchors = {
        anchor.replace(second=0, microsecond=0)
        for _row, anchor in anchor_rows
        if anchor is not None
    }
    use_anchor_matching = len(distinct_anchors) >= len(sorted_rows) and len(pool) >= len(sorted_rows)

    assignments = []
    available = list(pool)
    if use_anchor_matching:
        for row, anchor_ist in sorted(anchor_rows, key=lambda item: (item[1] is None, item[1] or datetime.max)):
            if not available:
                chosen = None
            elif anchor_ist is None:
                chosen = available.pop(0)
            else:
                chosen_idx = min(
                    range(len(available)),
                    key=lambda idx: abs(((_to_ist(available[idx].sent_time) or anchor_ist) - anchor_ist).total_seconds()),
                )
                chosen = available.pop(chosen_idx)
            assignments.append(
                {
                    "row": row,
                    "owner_reply": chosen,
                    "owner_reply_ist": _to_ist(chosen.sent_time) if chosen and chosen.sent_time else None,
                    "row_anchor_ist": anchor_ist,
                    "assign_mode": "subject_wide_anchor_unique",
                    "candidate_pool_label": f"subject_wide:{pool_label}",
                    "candidate_pool_times": pool_time_list,
                    "subject_wide_special": True,
                }
            )
        assignments.sort(key=lambda item: int(item["row"].get("_line", 10**9)))
        return assignments

    for row, anchor_ist in anchor_rows:
        if anchor_ist is None:
            chosen = pool[-1]
        else:
            chosen = min(
                pool,
                key=lambda email: abs(((_to_ist(email.sent_time) or anchor_ist) - anchor_ist).total_seconds()),
            )
        assignments.append(
            {
                "row": row,
                "owner_reply": chosen,
                "owner_reply_ist": _to_ist(chosen.sent_time) if chosen and chosen.sent_time else None,
                "row_anchor_ist": anchor_ist,
                "assign_mode": "subject_wide_anchor_reuse",
                "candidate_pool_label": f"subject_wide:{pool_label}",
                "candidate_pool_times": pool_time_list,
                "subject_wide_special": True,
            }
        )
    return assignments


def _assigned_family_rows(
    family_rows: list[dict],
    matched_emails: list[DebugEmail],
    output_by_line: dict[int, dict],
) -> tuple[str, list[dict], dict[int, int]]:
    sorted_rows = sorted(family_rows, key=lambda r: int(r.get("_line", 10**9)))
    strategy = _family_slot_strategy(sorted_rows)
    assignments: list[dict] = []
    family_anchor_ists = _family_anchor_ists(sorted_rows, output_by_line)
    special_rows = [row for row in sorted_rows if _row_is_all_ack_to_ess(row)]
    special_lines = {int(row.get("_line", 0)) for row in special_rows}
    normal_rows = [row for row in sorted_rows if int(row.get("_line", 0)) not in special_lines]

    if special_rows:
        assignments.extend(_assign_subject_wide_rows(special_rows, matched_emails, output_by_line))

    if strategy == "same_requester_chronological":
        if normal_rows:
            assignments.extend(_assign_requester_rows(normal_rows, matched_emails, output_by_line, family_anchor_ists))
    else:
        by_requester: dict[str, list[dict]] = {}
        requester_order: list[str] = []
        for row in normal_rows:
            requester = _get_col(row, "Requester", "Consultant").strip() or "<blank>"
            if requester not in by_requester:
                by_requester[requester] = []
                requester_order.append(requester)
            by_requester[requester].append(row)
        for requester in requester_order:
            assignments.extend(
                _assign_requester_rows(by_requester[requester], matched_emails, output_by_line, family_anchor_ists)
            )
    assignments.sort(
        key=lambda item: (
            item["owner_reply_ist"] is None,
            item["owner_reply_ist"].timestamp() if item["owner_reply_ist"] else float("inf"),
            int(item["row"].get("_line", 10**9)),
        )
    )

    slot_by_line: dict[int, int] = {}
    for idx, item in enumerate(assignments):
        line = int(item["row"].get("_line", 0))
        slot_by_line[line] = idx
        item["slot_index"] = idx
    return strategy, assignments, slot_by_line


def _suggest_slot_shape(
    row: dict,
    slot_idx: int,
    consultant_ess_pool: list[DebugEmail],
    ess_pool: list[DebugEmail],
    reply_pool: list[DebugEmail],
    ack_pool: list[DebugEmail],
    direct_pool: list[DebugEmail],
) -> tuple[str, str]:
    shape_hint, reason = _row_shape_hint(row)
    if shape_hint:
        return shape_hint, reason

    notes_l = _get_col(row, "Notes").lower()
    same_time_hint = any(
        token in notes_l
        for token in (
            "ess-only; no non-ess request",
            "ess initiated; no ack; no consultant reply after request",
            "failed subject; ess initiated; no ack phrase",
            "force prod subject; all times same",
        )
    )

    if slot_idx < len(direct_pool):
        return "triplet", "direct_pool_slot"
    if slot_idx < len(reply_pool) and slot_idx < len(ack_pool):
        return "hybrid", "reply_plus_ack_slot"
    if same_time_hint and (slot_idx < len(consultant_ess_pool) or slot_idx < len(ess_pool)):
        return "all_three_same", "ess_only_slot"
    if slot_idx < len(reply_pool):
        return "triplet", "reply_slot_only"
    if slot_idx < len(consultant_ess_pool) or slot_idx < len(ess_pool):
        return "all_three_same_candidate", "ess_pool_fallback"
    return "unknown", "no_usable_slot_shape"


def _same_minute(left: datetime | None, right: datetime | None) -> bool:
    if left is None or right is None:
        return False
    return left.replace(second=0, microsecond=0) == right.replace(second=0, microsecond=0)


def _preserve_same_time_decision(
    row: dict,
    assignment_info: dict,
    out_row: dict,
    suggested_shape: str,
) -> tuple[bool, str, str]:
    notes_l = _get_col(row, "Notes").lower()
    row_anchor = assignment_info.get("row_anchor_ist")
    owner_reply = assignment_info.get("owner_reply")
    owner_ist = _to_ist(owner_reply.sent_time) if owner_reply and owner_reply.sent_time else None
    created = _parse_datetime(_get_col(out_row, "Created Date & Time"))
    ack = _parse_datetime(_get_col(out_row, "Actual Response Date & Time"))
    resolved = _parse_datetime(_get_col(out_row, "Actual Resolved Date & Time"))
    current_all_same = bool(created and ack and resolved and created == ack == resolved)
    has_same_time_note = any(
        token in notes_l
        for token in (
            "occurrenceslotshape[allthreesame]",
            "esscontinuationguard[allthreestrictessonly]",
            "lanelocalinitialepisode[all-three-same]",
        )
    )
    is_subject_wide = bool(assignment_info.get("subject_wide_special"))

    if owner_ist is None or row_anchor is None:
        return False, "", "missing_owner_or_anchor"
    if not _same_minute(owner_ist, row_anchor):
        return False, "", "owner_anchor_mismatch"
    if suggested_shape in ("triplet", "hybrid") and not has_same_time_note:
        return False, "", f"shape_override_{suggested_shape}"
    if is_subject_wide and has_same_time_note:
        if "esscontinuationguard[allthreestrictessonly]" in notes_l:
            return True, "existing_strict_same_time_guard", ""
        return True, "subject_wide_all_three_same_aligned", ""
    if has_same_time_note and current_all_same:
        return True, "existing_all_three_same_output", ""
    return False, "", "no_preserve_rule_matched"


def _print_slot_rows(assignments: list[dict]) -> None:
    print("family rows by slot:")
    for item in assignments:
        row = item["row"]
        slot = int(item.get("slot_index", 0)) + 1
        req = _get_col(row, "Requester", "Consultant")
        desc = _get_col(row, "Description")
        notes = _get_col(row, "Notes")
        shape_hint, reason = _row_shape_hint(row)
        print(
            f"  slot={slot} | line={row.get('_line')} | requester={req} | triplet={_fmt_triplet(row)}"
        )
        owner_reply = item.get("owner_reply")
        if owner_reply:
            print(
                f"    owner_reply={_fmt(owner_reply.sent_time)} | owner_from={owner_reply.sender_email or owner_reply.sender_name}"
            )
        row_anchor = item.get("row_anchor_ist")
        if row_anchor is not None:
            print(f"    row_anchor={_fmt(row_anchor)}")
        assign_mode = item.get("assign_mode")
        if assign_mode:
            print(f"    assign_mode={assign_mode}")
        if item.get("subject_wide_special"):
            print("    assignment_scope=subject_wide_ess")
        candidate_pool_label = item.get("candidate_pool_label")
        candidate_pool_times = item.get("candidate_pool_times") or []
        if candidate_pool_label:
            print(f"    candidate_pool={candidate_pool_label} | times={_fmt_time_list(candidate_pool_times)}")
        if shape_hint:
            print(f"    shape_hint={shape_hint} | reason={reason}")
        print(f"    desc={desc}")
        if notes:
            print(f"    notes={notes}")


def _print_assignment_trace(assignments: list[dict]) -> None:
    print("ASSIGNMENT TRACE")
    for item in assignments:
        row = item["row"]
        print(
            f"  line={row.get('_line')} | requester={_get_col(row, 'Requester', 'Consultant')} | "
            f"row_anchor={_fmt(item.get('row_anchor_ist')) or '-'} | "
            f"owner_reply={_fmt(item.get('owner_reply').sent_time) if item.get('owner_reply') else '-'} | "
            f"assign_mode={item.get('assign_mode') or '-'} | "
            f"candidate_pool={item.get('candidate_pool_label') or '-'}"
        )
        if item.get("subject_wide_special"):
            print("    assignment_scope=subject_wide_ess")
        pool_times = item.get("candidate_pool_times") or []
        print(f"    pool_times={_fmt_time_list(pool_times)}")


def _trace_row_flow(
    family_rows: list[dict],
    assignments: list[dict],
    slot_by_line: dict[int, int],
    family_slot_strategy: str,
    output_by_line: dict[int, dict],
    consultant_ess_pool: list[DebugEmail],
    ess_pool: list[DebugEmail],
    reply_pool: list[DebugEmail],
    ack_pool: list[DebugEmail],
    direct_pool: list[DebugEmail],
) -> None:
    print("TRACE FLOW")
    sorted_rows = sorted(family_rows, key=lambda r: int(r.get("_line", 10**9)))
    print(f"  family_slot_strategy={family_slot_strategy}")
    assignment_by_line = {int(item["row"].get("_line", 0)): item for item in assignments}
    special_assignments = [item for item in assignments if _row_is_all_ack_to_ess(item["row"])]
    special_slot_by_line = {int(item["row"]["_line"]): idx for idx, item in enumerate(special_assignments)}

    usable_non_ess_lane = (len(reply_pool) >= len(sorted_rows) and len(ack_pool) >= len(sorted_rows)) or (len(direct_pool) >= len(sorted_rows))
    usable_consultant_ess_lane_special = len(consultant_ess_pool) >= len(special_assignments) if special_assignments else False

    for row in sorted_rows:
        line = int(row.get("_line", 0))
        req = _get_col(row, "Requester", "Consultant")
        notes = _get_col(row, "Notes")
        notes_l = notes.lower()
        out_row = output_by_line.get(line, {})
        current_triplet = _fmt_triplet(out_row) if out_row else "- / - / -"
        is_all_ack = _row_is_all_ack_to_ess(row)
        is_occ = _row_is_occurrence_managed(row)
        scope = "requester"
        slot_idx = slot_by_line.get(line, 0)
        expected_kind = "reply"
        expected_pick = None
        blocker = ""

        if is_all_ack and special_assignments:
            scope = "subject_wide_ess"
            slot_idx = special_slot_by_line.get(line, 0)
            assigned_owner = assignment_by_line.get(line, {}).get("owner_reply")
            if assigned_owner is not None:
                expected_kind = "ess_over_ess"
                expected_pick = assigned_owner
            elif usable_non_ess_lane:
                expected_kind = "reply_or_direct"
                if slot_idx < len(direct_pool):
                    expected_pick = direct_pool[slot_idx]
                elif slot_idx < len(reply_pool) and slot_idx < len(ack_pool):
                    expected_pick = reply_pool[slot_idx]
                else:
                    blocker = "non_ess_lane_not_slot_usable"
            elif usable_consultant_ess_lane_special:
                expected_kind = "ess_over_ess"
                if slot_idx < len(consultant_ess_pool):
                    expected_pick = consultant_ess_pool[slot_idx]
                else:
                    blocker = "consultant_ess_slot_missing"
            elif slot_idx < len(ess_pool):
                expected_kind = "ess_over_ess_generic"
                expected_pick = ess_pool[slot_idx]
            else:
                blocker = "no_usable_ess_lane"
        else:
            assigned_owner = assignment_by_line.get(line, {}).get("owner_reply")
            if assigned_owner is not None:
                expected_pick = assigned_owner
            if slot_idx < len(reply_pool):
                expected_kind = "reply"
                expected_pick = expected_pick or reply_pool[slot_idx]
            elif slot_idx < len(direct_pool):
                expected_kind = "direct_resolution"
                expected_pick = expected_pick or direct_pool[slot_idx]
            elif slot_idx < len(ack_pool):
                expected_kind = "ess_acky"
                expected_pick = expected_pick or ack_pool[slot_idx]
            elif slot_idx < len(ess_pool):
                expected_kind = "ess_over_ess_fallback"
                expected_pick = expected_pick or ess_pool[slot_idx]
            else:
                blocker = "no_slot_pick"

        expected_when = _fmt(expected_pick.sent_time) if expected_pick else "-"
        expected_triplet = (
            f"{expected_when} / {expected_when} / {expected_when}"
            if expected_pick and expected_kind.startswith("ess_over_ess")
            else "-"
        )
        suggested_shape, suggested_reason = _suggest_slot_shape(
            row,
            slot_idx,
            consultant_ess_pool,
            ess_pool,
            reply_pool,
            ack_pool,
            direct_pool,
        )
        assignment_info = assignment_by_line.get(line, {})
        preserve_same_time, preserve_reason, override_reason = _preserve_same_time_decision(
            row,
            assignment_info,
            out_row,
            suggested_shape,
        )
        has_final_guard = "esscontinuationguard[allthreestrictessonly]" in notes_l
        likely_drop = ""
        if preserve_same_time and not has_final_guard:
            likely_drop = "final_same_time_preserve_needed"
        elif expected_pick and expected_kind.startswith("ess_over_ess") and not has_final_guard:
            likely_drop = "final_ess_guard_not_applied"
        elif blocker:
            likely_drop = blocker

        print(
            f"  line={line} | requester={req} | occ={is_occ} | all_ack_to_ess={is_all_ack} | "
            f"scope={scope} | slot={slot_idx + 1} | expected_kind={expected_kind}"
        )
        print(f"    current_triplet={current_triplet}")
        if assignment_info.get("row_anchor_ist") is not None:
            print(f"    row_anchor={_fmt(assignment_info.get('row_anchor_ist'))}")
        if assignment_info.get("owner_reply") is not None:
            owner_reply = assignment_info["owner_reply"]
            print(f"    owner_reply={_fmt(owner_reply.sent_time)}")
        if assignment_info.get("assign_mode"):
            print(f"    assign_mode={assignment_info.get('assign_mode')}")
        if assignment_info.get("subject_wide_special"):
            print("    assignment_scope=subject_wide_ess")
        if assignment_info.get("candidate_pool_label"):
            print(
                f"    candidate_pool={assignment_info.get('candidate_pool_label')} | "
                f"times={_fmt_time_list(assignment_info.get('candidate_pool_times') or [])}"
            )
        print(f"    suggested_shape={suggested_shape} | reason={suggested_reason}")
        print(f"    preserve_same_time={preserve_same_time}")
        if preserve_reason:
            print(f"    preserve_reason={preserve_reason}")
        if override_reason:
            print(f"    override_reason={override_reason}")
        print(f"    expected_pick={expected_when}")
        if expected_triplet != "-":
            print(f"    expected_triplet={expected_triplet}")
        print(f"    final_ess_guard_note={has_final_guard}")
        if likely_drop:
            print(f"    likely_drop={likely_drop}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Debug repeated occurrence family decisions")
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

    _print_row_block("DEBUG ROWS", debug_rows)
    print("-" * 100)
    _print_row_block("OUTPUT ROWS", output_rows)
    print("-" * 100)

    if not debug_rows:
        print("No matching debug rows found.")
        return

    families = _cluster_family_rows(debug_rows)

    all_emails = list(_iter_emails(Path(args.eml_dir)))
    for subject_norm, family_rows in families.items():
        print("=" * 100)
        requesters = []
        seen_requesters = set()
        for row in family_rows:
            req = _get_col(row, "Requester", "Consultant")
            req_key = req.lower()
            if req and req_key not in seen_requesters:
                seen_requesters.add(req_key)
                requesters.append(req)
        family_anchors = _family_anchor_ists(family_rows, output_by_line)
        print(f"family_subject_norm={subject_norm}")
        print(f"family_requesters={', '.join(requesters) if requesters else '<none>'}")
        print(f"family_anchors={_fmt_time_list(family_anchors)}")

        matched_emails = []
        for email in all_emails:
            if not email.sent_time:
                continue
            if not _subject_match(subject_norm, email.subject):
                continue
            matched_emails.append(email)
        matched_emails.sort(key=lambda x: x.sent_time or datetime.max)
        family_slot_strategy, assignments, slot_by_line = _assigned_family_rows(family_rows, matched_emails, output_by_line)
        print("-" * 100)
        _print_slot_rows(assignments)
        print("-" * 100)
        _print_assignment_trace(assignments)
        print("-" * 100)
        print(f"matched_emails={len(matched_emails)}")

        reply_pool = []
        ack_pool = []
        direct_pool = []
        ess_pool = []
        consultant_ess_pool = []
        for email in matched_emails:
            cls = _classify_reply_kind(email)
            is_ess = _is_ess_sender(email, ess_team)
            req_match = _requester_match_any(email, requesters)
            if is_ess and not (cls.get("thanks_info") or cls.get("nonfinal_followup")):
                ess_pool.append(email)
            if is_ess and req_match and not (cls.get("thanks_info") or cls.get("nonfinal_followup")):
                consultant_ess_pool.append(email)
            if req_match and cls.get("real_reply"):
                reply_pool.append(email)
            if req_match and (cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack")):
                ack_pool.append(email)
            if req_match and cls.get("direct_resolution"):
                direct_pool.append(email)

        reply_pool = _minute_dedupe(reply_pool)
        ack_pool = _minute_dedupe(ack_pool)
        direct_pool = _minute_dedupe(direct_pool)
        ess_pool = _minute_dedupe(ess_pool)
        consultant_ess_pool = _minute_dedupe(consultant_ess_pool)

        requester_counts: dict[str, int] = {}
        for row in family_rows:
            req = _get_col(row, "Requester", "Consultant") or "<blank>"
            requester_counts[req] = requester_counts.get(req, 0) + 1
        all_ack_to_ess_rows = sum(1 for row in family_rows if _row_is_all_ack_to_ess(row))
        occurrence_rows = sum(1 for row in family_rows if _row_is_occurrence_managed(row))
        print(f"requester_row_counts={requester_counts}")
        print(f"occurrence_managed_rows={occurrence_rows}/{len(family_rows)}")
        print(f"all_ack_to_ess_rows={all_ack_to_ess_rows}/{len(family_rows)}")
        print(
            "subject_wide_ess_candidate="
            f"{all_ack_to_ess_rows >= 1 and occurrence_rows >= 2}"
        )
        print("-" * 100)

        _print_pool("REPLY POOL", reply_pool, ess_team, requesters)
        print("-" * 100)
        _print_pool("ACK POOL", ack_pool, ess_team, requesters)
        print("-" * 100)
        _print_pool("DIRECT POOL", direct_pool, ess_team, requesters)
        print("-" * 100)
        _print_pool("CONSULTANT ESS POOL", consultant_ess_pool, ess_team, requesters)
        print("-" * 100)
        _print_pool("ESS POOL", ess_pool, ess_team, requesters)
        print("-" * 100)

        group_size = len(family_rows)
        print(f"group_size={group_size}")
        print(f"usable_non_ess_lane={(len(reply_pool) >= group_size and len(ack_pool) >= group_size) or (len(direct_pool) >= group_size)}")
        print(f"usable_consultant_ess_lane={len(consultant_ess_pool) >= group_size}")
        print(f"usable_ess_lane={len(ess_pool) >= group_size}")
        print("slot assignments:")
        for slot in range(group_size):
            reply_pick = _fmt(reply_pool[slot].sent_time) if slot < len(reply_pool) else ""
            ack_pick = _fmt(ack_pool[slot].sent_time) if slot < len(ack_pool) else ""
            direct_pick = _fmt(direct_pool[slot].sent_time) if slot < len(direct_pool) else ""
            consultant_ess_pick = _fmt(consultant_ess_pool[slot].sent_time) if slot < len(consultant_ess_pool) else ""
            ess_pick = _fmt(ess_pool[slot].sent_time) if slot < len(ess_pool) else ""
            print(
                f"  slot={slot + 1} | reply={reply_pick or '-'} | ack={ack_pick or '-'} | "
                f"direct={direct_pick or '-'} | consultant_ess={consultant_ess_pick or '-'} | ess={ess_pick or '-'}"
            )
        print("-" * 100)
        _trace_row_flow(
            family_rows,
            assignments,
            slot_by_line,
            family_slot_strategy,
            output_by_line,
            consultant_ess_pool,
            ess_pool,
            reply_pool,
            ack_pool,
            direct_pool,
        )


if __name__ == "__main__":
    main()
