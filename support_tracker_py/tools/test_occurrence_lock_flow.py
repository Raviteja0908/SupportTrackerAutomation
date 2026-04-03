import argparse
import csv
import json
from dataclasses import dataclass
from datetime import datetime
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


def _main_line_refs() -> dict[str, int | None]:
    main_path = Path(__file__).resolve().parents[1] / "main.py"
    refs: dict[str, int | None] = {
        "override_entry_guard": None,
        "override_notes_gate": None,
        "override_group_size_gate": None,
        "override_slot_gate": None,
        "override_pool_gate": None,
        "override_pick_gate": None,
        "apply_shared_occurrence_triplet_def": None,
        "ess_only_strict_shared_triplet_call": None,
        "ess_only_strict_authoritative_call": None,
        "ess_only_strict_manual_occ_pick_branch": None,
        "ess_only_strict_manual_notes_write": None,
        "final_occurrence_workbook_if": None,
        "final_occurrence_notes_gate": None,
        "final_occurrence_plan_gate": None,
        "final_occurrence_pick_gate": None,
        "final_occurrence_triplet_gate": None,
        "final_occurrence_branch_if": None,
        "final_occurrence_apply_call": None,
        "authoritative_created_write": None,
        "authoritative_ack_write": None,
        "authoritative_resolved_write": None,
        "lock_call": None,
    }
    try:
        for idx, line in enumerate(main_path.read_text(encoding="utf-8").splitlines(), start=1):
            if refs["final_occurrence_workbook_if"] is None and 'if workbook_kind == "incident_business":' in line:
                refs["final_occurrence_workbook_if"] = idx
            if refs["override_entry_guard"] is None and "if list_index is None or not requester or not subject_norm_value:" in line:
                refs["override_entry_guard"] = idx
            if refs["override_notes_gate"] is None and "if not _is_all_ack_to_ess_notes(notes_l):" in line:
                refs["override_notes_gate"] = idx
            if refs["override_group_size_gate"] is None and "if len(group_sorted) < 2:" in line:
                refs["override_group_size_gate"] = idx
            if refs["override_slot_gate"] is None and "if slot_index is None:" in line:
                refs["override_slot_gate"] = idx
            if refs["override_pool_gate"] is None and (
                "if len(pool) < len(group_sorted):" in line
                or "if not pool or len(pool) < len(group_sorted):" in line
            ):
                refs["override_pool_gate"] = idx
            if refs["override_pick_gate"] is None and "if not pick_ist:" in line:
                refs["override_pick_gate"] = idx
            if refs["apply_shared_occurrence_triplet_def"] is None and "def _apply_shared_occurrence_triplet(" in line:
                refs["apply_shared_occurrence_triplet_def"] = idx
            if refs["ess_only_strict_shared_triplet_call"] is None and "_apply_shared_occurrence_triplet(" in line and "def _apply_shared_occurrence_triplet(" not in line:
                refs["ess_only_strict_shared_triplet_call"] = idx
            if refs["ess_only_strict_authoritative_call"] is None and idx > 11000 and "and _apply_occurrence_plan_authoritatively(" in line:
                refs["ess_only_strict_authoritative_call"] = idx
            if (
                refs["ess_only_strict_manual_occ_pick_branch"] is None
                and refs["ess_only_strict_authoritative_call"] is not None
                and idx > refs["ess_only_strict_authoritative_call"]
                and 'if shared_occ_plan and shared_occ_plan.get("pick") is not None:' in line
            ):
                refs["ess_only_strict_manual_occ_pick_branch"] = idx
            if (
                refs["ess_only_strict_manual_notes_write"] is None
                and refs["ess_only_strict_manual_occ_pick_branch"] is not None
                and idx > refs["ess_only_strict_manual_occ_pick_branch"]
                and 'ESSContinuationGuard[AllThreeStrictEssOnly]' in line
                and 'debug_rows[list_index]["Notes"]' in line
            ):
                refs["ess_only_strict_manual_notes_write"] = idx
            if refs["final_occurrence_notes_gate"] is None and "if not _is_all_ack_to_ess_notes(notes_l):" in line:
                refs["final_occurrence_notes_gate"] = idx
            if refs["final_occurrence_plan_gate"] is None and (
                'if not shared_occ_plan or (shared_occ_plan.get("lane_kind") or "") != "ess_over_ess":' in line
                or 'if not shared_occ_plan or not _is_authoritative_occurrence_lane(shared_occ_plan.get("lane_kind") or ""):' in line
            ):
                refs["final_occurrence_plan_gate"] = idx
            if refs["final_occurrence_pick_gate"] is None and "if not pick_when:" in line:
                refs["final_occurrence_pick_gate"] = idx
            if refs["final_occurrence_triplet_gate"] is None and "if current_triplet != target_triplet:" in line:
                refs["final_occurrence_triplet_gate"] = idx
            if refs["final_occurrence_branch_if"] is None and "if current_triplet != target_triplet:" in line:
                refs["final_occurrence_branch_if"] = idx
            if refs["final_occurrence_apply_call"] is None and "applied = _apply_occurrence_plan_authoritatively(" in line:
                refs["final_occurrence_apply_call"] = idx
            if refs["authoritative_created_write"] is None and 'row_vals["Created Date & Time"] = t' in line:
                refs["authoritative_created_write"] = idx
            if refs["authoritative_ack_write"] is None and 'row_vals["Actual Response Date & Time"] = t' in line:
                refs["authoritative_ack_write"] = idx
            if refs["authoritative_resolved_write"] is None and 'row_vals["Actual Resolved Date & Time"] = t' in line:
                refs["authoritative_resolved_write"] = idx
            if refs["lock_call"] is None and "if _lock_occurrence_row(" in line:
                refs["lock_call"] = idx
    except Exception:
        pass
    return refs


MAIN_LINE_REFS = _main_line_refs()


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


def _fmt(dt: datetime | None) -> str:
    if not dt:
        return "-"
    ist = _to_ist(dt)
    return ist.strftime("%d-%m-%Y %H:%M") if ist else "-"


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


def _row_triplet(row: dict) -> tuple[datetime | None, datetime | None, datetime | None]:
    return (
        _parse_datetime(_get_col(row, "Created Date & Time")),
        _parse_datetime(_get_col(row, "Actual Response Date & Time")),
        _parse_datetime(_get_col(row, "Actual Resolved Date & Time")),
    )


def _fmt_triplet(row: dict) -> str:
    c_dt, a_dt, r_dt = _row_triplet(row)
    return f"{_fmt(c_dt)} / {_fmt(a_dt)} / {_fmt(r_dt)}"


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


def _row_is_all_ack_to_ess(row: dict) -> bool:
    notes = _get_col(row, "Notes").lower()
    return "requester span(all-ack->ess)" in notes and "ess-only; no non-ess request" in notes


def _requester_match_any(email: DebugEmail, requesters: list[str]) -> bool:
    return any(_match_requester(email.sender_name, email.sender_email, req) for req in requesters if req)


def _row_is_occurrence_managed(row: dict) -> bool:
    notes = _get_col(row, "Notes").lower()
    return (
        "dateanchoroccurrence" in notes
        or "ess-only; no non-ess request" in notes
        or "requester follow-up" in notes
        or "esscontinuationguard[" in notes
        or "quotedrequestonlynopair" in notes
    )


def _collect_family_pools(
    family_subject: str,
    family_rows: list[dict],
    all_emails: list[DebugEmail],
    ess_team: list[str],
) -> tuple[list[DebugEmail], list[DebugEmail], list[DebugEmail], list[DebugEmail], list[DebugEmail], list[DebugEmail]]:
    requesters = []
    seen = set()
    for row in family_rows:
        req = _get_col(row, "Requester", "Consultant")
        req_key = req.lower()
        if req and req_key not in seen:
            seen.add(req_key)
            requesters.append(req)

    matched = []
    for email in all_emails:
        if not email.sent_time:
            continue
        if _subject_match(family_subject, email.subject):
            matched.append(email)

    reply_pool = []
    ack_pool = []
    direct_pool = []
    consultant_ess_pool = []
    ess_pool = []
    occurrence_acky_pool = []
    for email in matched:
        cls = _classify_reply_kind(email)
        is_ess = _is_ess_sender(email, ess_team)
        req_match = _requester_match_any(email, requesters)
        if is_ess and _is_real_ess_progression(email, ess_team):
            ess_pool.append(email)
        if is_ess and req_match and _is_real_ess_progression(email, ess_team):
            consultant_ess_pool.append(email)
        if req_match and cls.get("real_reply"):
            reply_pool.append(email)
        if req_match and (cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack")):
            ack_pool.append(email)
        if req_match and cls.get("direct_resolution"):
            direct_pool.append(email)
        if _is_occurrence_acky_candidate(email, requesters):
            occurrence_acky_pool.append(email)

    return (
        _minute_dedupe(reply_pool),
        _minute_dedupe(ack_pool),
        _minute_dedupe(direct_pool),
        _minute_dedupe(consultant_ess_pool),
        _minute_dedupe(ess_pool),
        _minute_dedupe(occurrence_acky_pool),
    )


def _is_ack_like_only(email: DebugEmail) -> bool:
    cls = _classify_reply_kind(email)
    is_ack = bool(cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack"))
    is_substantive = bool(cls.get("real_reply") or cls.get("direct_resolution"))
    return is_ack and not is_substantive


def _is_real_ess_progression(email: DebugEmail, ess_team: list[str]) -> bool:
    cls = _classify_reply_kind(email)
    if not _is_ess_sender(email, ess_team):
        return False
    if cls.get("thanks_info") or cls.get("nonfinal_followup"):
        return False
    if cls.get("direct_resolution") or cls.get("real_reply"):
        return True
    return not _is_ack_like_only(email)


def _is_occurrence_acky_candidate(email: DebugEmail, requesters: list[str]) -> bool:
    if not email.sent_time:
        return False
    if not _requester_match_any(email, requesters):
        return False
    cls = _classify_reply_kind(email)
    if cls.get("thanks_info") or cls.get("nonfinal_followup"):
        return False
    return bool(
        cls.get("real_reply")
        or cls.get("direct_resolution")
        or cls.get("ack_like")
        or cls.get("explicit_ack")
        or cls.get("short_ess_ack")
    )


def _lane_is_authoritative_occurrence(lane_kind: str | None) -> bool:
    return (lane_kind or "") in {"ess_over_ess", "ess_acky_sequence"}


def _dedupe_reply_minutes_prefer_consultant(items: list[DebugEmail], requesters: list[str]) -> list[DebugEmail]:
    buckets: dict[datetime, list[DebugEmail]] = {}
    for item in items:
        if not item.sent_time:
            continue
        minute_key = _to_ist(item.sent_time).replace(second=0, microsecond=0)
        buckets.setdefault(minute_key, []).append(item)

    out = []
    for minute_key in sorted(buckets):
        bucket = buckets[minute_key]
        bucket.sort(
            key=lambda email: (
                0 if _requester_match_any(email, requesters) else 1,
                0 if (_classify_reply_kind(email).get("direct_resolution")) else 1,
                0 if (_classify_reply_kind(email).get("real_reply")) else 1,
                1 if _is_ack_like_only(email) else 0,
                _to_ist(email.sent_time),
            )
        )
        out.append(bucket[0])
    return out


def _exact_shared_occurrence_plan(
    target_row: dict,
    family_rows: list[dict],
    output_by_line: dict[int, dict],
    all_emails: list[DebugEmail],
    ess_team: list[str],
):
    notes_l = _get_col(target_row, "Notes").lower()
    requester = _get_col(target_row, "Requester", "Consultant")
    subject_norm_value = _family_subject_norm(target_row)
    if not requester or not subject_norm_value:
        return None

    current_is_all_ack = _row_is_all_ack_to_ess(target_row)

    requester_group = []
    requester_group_requesters = set()
    requester_allow_acky = False
    subject_group = []
    subject_group_requesters = set()
    subject_allow_acky = False

    for row in family_rows:
        other_notes_l = _get_col(row, "Notes").lower()
        if not _row_is_occurrence_managed(row):
            continue
        if _family_subject_norm(row) == subject_norm_value and _get_col(row, "Requester", "Consultant") == requester:
            requester_group.append(row)
            req = _get_col(row, "Requester", "Consultant")
            if req:
                requester_group_requesters.add(req)
            requester_allow_acky = requester_allow_acky or ("requester span(all-ack->ess)" in other_notes_l)
        if current_is_all_ack:
            if not _row_is_all_ack_to_ess(row):
                continue
            if not _subject_match(subject_norm_value, _family_subject_norm(row)):
                continue
            subject_group.append(row)
            req = _get_col(row, "Requester", "Consultant")
            if req:
                subject_group_requesters.add(req)
            subject_allow_acky = subject_allow_acky or ("requester span(all-ack->ess)" in other_notes_l)

    requester_group.sort(key=lambda r: int(r.get("_line", 10**9)))
    subject_group.sort(key=lambda r: int(r.get("_line", 10**9)))
    if len(subject_group) >= 2 and len(subject_group) > len(requester_group):
        group_sorted = subject_group
        group_requesters = sorted(subject_group_requesters)
        allow_acky = subject_allow_acky
        selected_scope = "subject_wide_ess"
    else:
        group_sorted = requester_group
        group_requesters = sorted(requester_group_requesters)
        allow_acky = requester_allow_acky
        selected_scope = "requester"

    if len(group_sorted) < 2:
        return None

    slot_index = None
    target_line = int(target_row["_line"])
    for idx, row in enumerate(group_sorted):
        if int(row["_line"]) == target_line:
            slot_index = idx
            break
    if slot_index is None:
        return None

    merged = []
    for email in all_emails:
        if email.sent_time and _subject_match(subject_norm_value, email.subject):
            merged.append(email)
    merged.sort(key=lambda e: e.sent_time or datetime.max)
    if not merged:
        return None

    def _group_req_match(email: DebugEmail) -> bool:
        return _requester_match_any(email, group_requesters)

    def _collect_pool(allow_acky_local: bool, use_ess_pool: bool) -> list[DebugEmail]:
        out = []
        for email in merged:
            if not email.sent_time:
                continue
            cls = _classify_reply_kind(email)
            if use_ess_pool:
                if not _is_real_ess_progression(email, ess_team):
                    continue
            else:
                if not _group_req_match(email):
                    continue
                if not allow_acky_local and not cls.get("real_reply"):
                    continue
            if not _subject_match(subject_norm_value, email.subject):
                continue
            out.append(email)
        return _dedupe_reply_minutes_prefer_consultant(out, group_requesters)

    def _collect_non_ess_ack_pool() -> list[DebugEmail]:
        out = []
        for email in merged:
            if not email.sent_time or not _group_req_match(email):
                continue
            cls = _classify_reply_kind(email)
            if cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack"):
                out.append(email)
        return _dedupe_reply_minutes_prefer_consultant(out, group_requesters)

    def _collect_direct_pool() -> list[DebugEmail]:
        out = []
        for email in merged:
            if not email.sent_time or not _group_req_match(email):
                continue
            if _classify_reply_kind(email).get("direct_resolution"):
                out.append(email)
        return _dedupe_reply_minutes_prefer_consultant(out, group_requesters)

    def _collect_consultant_ess_pool() -> list[DebugEmail]:
        out = []
        for email in merged:
            if not email.sent_time:
                continue
            if not _group_req_match(email):
                continue
            if not _is_real_ess_progression(email, ess_team):
                continue
            out.append(email)
        return _dedupe_reply_minutes_prefer_consultant(out, group_requesters)

    def _collect_occurrence_acky_pool() -> list[DebugEmail]:
        out = []
        for email in merged:
            if not _is_occurrence_acky_candidate(email, group_requesters):
                continue
            out.append(email)
        return _dedupe_reply_minutes_prefer_consultant(out, group_requesters)

    reply_pool = _collect_pool(False, False)
    acky_pool = _collect_pool(True, False) if allow_acky else list(reply_pool)
    ess_pool = _collect_pool(True, True)
    ack_pool = _collect_non_ess_ack_pool()
    direct_pool = _collect_direct_pool()
    consultant_ess_pool = _collect_consultant_ess_pool()
    occurrence_acky_pool = _collect_occurrence_acky_pool()

    anchor_row = output_by_line.get(int(group_sorted[0]["_line"]), {})
    anchor_ack = _parse_datetime(_get_col(anchor_row, "Actual Response Date & Time"))
    anchor_ack_ist = _to_ist(anchor_ack) if anchor_ack else None

    def _same_month_pool(pool_in: list[DebugEmail]) -> list[DebugEmail]:
        if not anchor_ack_ist:
            return list(pool_in)
        same_month = []
        for email in pool_in:
            e_ist = _to_ist(email.sent_time) if email.sent_time else None
            if e_ist and e_ist.year == anchor_ack_ist.year and e_ist.month == anchor_ack_ist.month:
                same_month.append(email)
        if len(same_month) >= len(group_sorted):
            return same_month
        return list(pool_in)

    reply_pool = _same_month_pool(reply_pool)
    acky_pool = _same_month_pool(acky_pool)
    ess_pool = _same_month_pool(ess_pool)
    ack_pool = _same_month_pool(ack_pool)
    direct_pool = _same_month_pool(direct_pool)
    consultant_ess_pool = _same_month_pool(consultant_ess_pool)
    occurrence_acky_pool = _same_month_pool(occurrence_acky_pool)

    default_pool = reply_pool or acky_pool or consultant_ess_pool or ess_pool
    if len(default_pool) < len(group_sorted):
        return {
            "exists": False,
            "reason": "default_pool_too_short",
            "scope": selected_scope,
            "group_size": len(group_sorted),
            "slot_index": slot_index,
            "reply_pool": reply_pool,
            "acky_pool": acky_pool,
            "consultant_ess_pool": consultant_ess_pool,
            "ess_pool": ess_pool,
            "occurrence_acky_pool": occurrence_acky_pool,
            "ack_pool": ack_pool,
            "direct_pool": direct_pool,
        }

    def _pick_from(pool_in: list[DebugEmail], lane_kind: str):
        if not pool_in or len(pool_in) < len(group_sorted):
            return None
        email = pool_in[min(slot_index, len(pool_in) - 1)]
        return {
            "lane_kind": lane_kind,
            "pick": email,
            "pick_when": _to_ist(email.sent_time).replace(second=0, microsecond=0),
        }

    strong_non_ess_live = False
    if current_is_all_ack:
        if slot_index < len(direct_pool):
            strong_non_ess_live = True
        elif (
            len(reply_pool) >= max(1, len(group_sorted))
            and slot_index < len(reply_pool)
            and slot_index < len(ack_pool)
        ):
            strong_non_ess_live = True

        if not strong_non_ess_live:
            plan = _pick_from(consultant_ess_pool, "ess_over_ess") or _pick_from(ess_pool, "ess_over_ess")
            if plan:
                return {
                    "exists": True,
                    "scope": selected_scope,
                    "group_size": len(group_sorted),
                    "slot_index": slot_index,
                    "allow_acky": allow_acky,
                    "reply_pool": reply_pool,
                    "acky_pool": acky_pool,
                    "consultant_ess_pool": consultant_ess_pool,
                    "ess_pool": ess_pool,
                    "occurrence_acky_pool": occurrence_acky_pool,
                    "ack_pool": ack_pool,
                    "direct_pool": direct_pool,
                    "strong_non_ess_live": strong_non_ess_live,
                    **plan,
                }
            plan = _pick_from(occurrence_acky_pool, "ess_acky_sequence")
            if plan:
                return {
                    "exists": True,
                    "scope": selected_scope,
                    "group_size": len(group_sorted),
                    "slot_index": slot_index,
                    "allow_acky": allow_acky,
                    "reply_pool": reply_pool,
                    "acky_pool": acky_pool,
                    "consultant_ess_pool": consultant_ess_pool,
                    "ess_pool": ess_pool,
                    "occurrence_acky_pool": occurrence_acky_pool,
                    "ack_pool": ack_pool,
                    "direct_pool": direct_pool,
                    "strong_non_ess_live": strong_non_ess_live,
                    **plan,
                }

    plan = (
        _pick_from(reply_pool, "reply")
        or _pick_from(acky_pool, "ess_acky")
        or _pick_from(consultant_ess_pool, "ess_over_ess")
        or _pick_from(ess_pool, "ess_over_ess")
    )
    if not plan:
        return {
            "exists": False,
            "reason": "no_pick_after_fill_plan",
            "scope": selected_scope,
            "group_size": len(group_sorted),
            "slot_index": slot_index,
            "reply_pool": reply_pool,
            "acky_pool": acky_pool,
            "consultant_ess_pool": consultant_ess_pool,
            "ess_pool": ess_pool,
            "ack_pool": ack_pool,
            "direct_pool": direct_pool,
        }
    return {
        "exists": True,
        "scope": selected_scope,
        "group_size": len(group_sorted),
        "slot_index": slot_index,
        "allow_acky": allow_acky,
        "reply_pool": reply_pool,
        "acky_pool": acky_pool,
        "consultant_ess_pool": consultant_ess_pool,
        "ess_pool": ess_pool,
        "occurrence_acky_pool": occurrence_acky_pool,
        "ack_pool": ack_pool,
        "direct_pool": direct_pool,
        "strong_non_ess_live": strong_non_ess_live,
        **plan,
    }


def _proposed_locked_triplet(
    row: dict,
    output_by_line: dict[int, dict],
    exact_plan: dict | None,
) -> str:
    notes_l = _get_col(row, "Notes").lower()
    current_triplet = _fmt_triplet(output_by_line.get(int(row["_line"]), {}))
    if not exact_plan or not exact_plan.get("exists"):
        return current_triplet

    lane_kind = exact_plan.get("lane_kind") or ""
    pick = exact_plan.get("pick")
    if _lane_is_authoritative_occurrence(lane_kind) and "requester span(all-ack->ess)" in notes_l and pick is not None:
        t = _fmt(pick.sent_time)
        return f"{t} / {t} / {t}"

    return current_triplet


def _classify_row_stage(row: dict) -> list[str]:
    notes_l = _get_col(row, "Notes").lower()
    flags = []
    if "quotedrequestonly" in notes_l:
        flags.append("quoted_request_only")
    if "quotedrequestonlynopair" in notes_l:
        flags.append("quoted_no_pair")
    if "dateanchor" in notes_l:
        flags.append("date_anchor")
    if "dateanchoroccurrence" in notes_l:
        flags.append("date_anchor_occurrence")
    if "ackwindowguard" in notes_l:
        flags.append("ack_window_guard")
    if "bluequotedpairreanchor" in notes_l:
        flags.append("blue_quoted_reanchor")
    if "blueclearedstrict" in notes_l:
        flags.append("blue_cleared_strict")
    if "esscontinuationguard[" in notes_l:
        flags.append("ess_continuation_guard")
    if "occurrencelocked" in notes_l:
        flags.append("occurrence_locked")
    return flags


def _build_row_flow_summary(
    row: dict,
    output_by_line: dict[int, dict],
    exact_plan: dict | None,
) -> dict:
    notes_l = _get_col(row, "Notes").lower()
    current_triplet = _fmt_triplet(output_by_line.get(int(row["_line"]), {}))
    proposed_triplet = _proposed_locked_triplet(row, output_by_line, exact_plan)
    lane_kind = (exact_plan or {}).get("lane_kind") if exact_plan else None
    pick = (exact_plan or {}).get("pick") if exact_plan else None
    pick_when = _fmt(pick.sent_time) if pick else "-"
    should_apply_occurrence = bool(
        exact_plan
        and exact_plan.get("exists")
        and _lane_is_authoritative_occurrence(lane_kind)
        and "requester span(all-ack->ess)" in notes_l
    )
    should_lock = bool(
        should_apply_occurrence
        or ("occurrencelocked" in notes_l)
        or (current_triplet == proposed_triplet and proposed_triplet != "- / - / -")
    )
    return {
        "line": int(row["_line"]),
        "requester": _get_col(row, "Requester", "Consultant"),
        "description": _get_col(row, "Description"),
        "notes_flags": _classify_row_stage(row),
        "current_triplet": current_triplet,
        "exact_plan_exists": bool(exact_plan and exact_plan.get("exists")),
        "exact_scope": (exact_plan or {}).get("scope"),
        "exact_group_size": (exact_plan or {}).get("group_size"),
        "exact_slot": ((exact_plan or {}).get("slot_index", -1) + 1) if exact_plan and exact_plan.get("exists") else None,
        "exact_lane_kind": lane_kind,
        "exact_pick": pick_when,
        "should_apply_occurrence": should_apply_occurrence,
        "should_lock": should_lock,
        "proposed_triplet": proposed_triplet,
        "changed": current_triplet != proposed_triplet,
    }


def _runtime_mirror_trace(row: dict, output_by_line: dict[int, dict], exact_plan: dict | None) -> list[dict]:
    notes_l = _get_col(row, "Notes").lower()
    out_row = output_by_line.get(int(row["_line"]), {})
    current_triplet = _fmt_triplet(out_row)
    trace = []

    def add(step: str, result, detail: str = ""):
        trace.append({
            "step": step,
            "result": result,
            "detail": detail,
        })

    add("row_has_list_index", bool(row.get("_line")), f"line={row.get('_line')}")
    add("row_not_deployment", True, "script rows exclude deployment handling")
    add("row_not_already_locked", "occurrencelocked" in notes_l is False, f"notes_has_occurrence_locked={'occurrencelocked' in notes_l}")
    add(
        "main_final_occurrence_apply_call_line",
        MAIN_LINE_REFS.get("final_occurrence_apply_call"),
        (
            f"workbook_if={MAIN_LINE_REFS.get('final_occurrence_workbook_if')}, "
            f"notes_gate={MAIN_LINE_REFS.get('final_occurrence_notes_gate')}, "
            f"plan_gate={MAIN_LINE_REFS.get('final_occurrence_plan_gate')}, "
            f"pick_gate={MAIN_LINE_REFS.get('final_occurrence_pick_gate')}, "
            f"triplet_gate={MAIN_LINE_REFS.get('final_occurrence_triplet_gate')}, "
            f"branch_if={MAIN_LINE_REFS.get('final_occurrence_branch_if')}, "
            f"apply_call={MAIN_LINE_REFS.get('final_occurrence_apply_call')}, "
            f"lock_call={MAIN_LINE_REFS.get('lock_call')}"
        ),
    )
    add(
        "main_authoritative_write_lines",
        True,
        (
            f"created={MAIN_LINE_REFS.get('authoritative_created_write')}, "
            f"ack={MAIN_LINE_REFS.get('authoritative_ack_write')}, "
            f"resolved={MAIN_LINE_REFS.get('authoritative_resolved_write')}"
        ),
    )

    add("notes_contains_all_ack_to_ess", _row_is_all_ack_to_ess(row), notes_l)

    shared_occ_plan_exists = bool(exact_plan and exact_plan.get("exists"))
    add("shared_occ_plan_exists", shared_occ_plan_exists, f"exact_plan_exists={shared_occ_plan_exists}")

    lane_kind = (exact_plan or {}).get("lane_kind") if exact_plan else None
    add("shared_occ_plan_lane_kind", lane_kind or "<none>", f"scope={(exact_plan or {}).get('scope')}")

    pick = (exact_plan or {}).get("pick") if exact_plan else None
    pick_when = _fmt(pick.sent_time) if pick else "-"
    add("shared_occ_plan_pick_when", bool(pick), f"pick_when={pick_when}")

    current_target_mismatch = False
    if pick is not None:
        current_target_mismatch = current_triplet != f"{pick_when} / {pick_when} / {pick_when}"
    add(
        "current_triplet_differs_from_target",
        current_target_mismatch,
        f"triplet_gate_line={MAIN_LINE_REFS.get('final_occurrence_triplet_gate')}",
    )

    override_needed = bool(
        _row_is_all_ack_to_ess(row)
        and (not shared_occ_plan_exists or not _lane_is_authoritative_occurrence(lane_kind))
    )
    add("override_plan_needed", override_needed, f"lane_kind={lane_kind}")

    override_available = bool(
        _row_is_all_ack_to_ess(row)
        and exact_plan
        and exact_plan.get("exists")
        and _lane_is_authoritative_occurrence(lane_kind)
    )
    add("override_plan_available", override_available, f"exact_lane_kind={lane_kind}")

    enter_occurrence_apply_branch = bool(
        _lane_is_authoritative_occurrence(lane_kind)
        and ("requester span(all-ack->ess)" in notes_l)
        and pick is not None
    )
    add(
        "enter_occurrence_apply_branch",
        enter_occurrence_apply_branch,
        f"lane_kind={lane_kind}, notes_all_ack={'requester span(all-ack->ess)' in notes_l}, pick_when={pick_when}",
    )

    if enter_occurrence_apply_branch:
        t = pick_when
        target_triplet = f"{t} / {t} / {t}"
        add("authoritative_apply_would_format_time", t != "-", f"formatted={t}")
        add("authoritative_apply_target_triplet", True, target_triplet)
        add(
            "exact_runtime_handoff_line",
            MAIN_LINE_REFS.get("final_occurrence_apply_call"),
            f"if runtime diverges, first handoff is main.py:{MAIN_LINE_REFS.get('final_occurrence_apply_call')}",
        )
        add("authoritative_apply_would_write_row", True, f"current={current_triplet} -> target={target_triplet}")
        add("lock_would_set_occurrence_locked", True, target_triplet)
        add(
            "runtime_matches_expected_after_apply",
            current_triplet == target_triplet,
            f"current={current_triplet}, target={target_triplet}",
        )
        first_divergence_line = MAIN_LINE_REFS.get("final_occurrence_notes_gate")
        add(
            "first_possible_runtime_divergence_line",
            first_divergence_line,
            (
                "all mirrored gates are satisfied in script; if runtime still leaves old values, "
                f"the earliest gate that must differ is main.py:{first_divergence_line}"
            ),
        )
    else:
        add("authoritative_apply_would_write_row", False, "branch not entered")
        add("lock_would_set_occurrence_locked", False, "branch not entered")

    return trace


def _override_function_trace(row: dict, exact_plan: dict | None) -> list[dict]:
    notes_l = _get_col(row, "Notes").lower()
    requester = _get_col(row, "Requester", "Consultant")
    subject_norm_value = _family_subject_norm(row)
    trace = []

    def add(step: str, result, detail: str = ""):
        trace.append({
            "step": step,
            "result": result,
            "detail": detail,
        })

    add(
        "override_entry_guard",
        bool(row.get("_line") and requester and subject_norm_value),
        f"line={row.get('_line')}, requester={bool(requester)}, subject={bool(subject_norm_value)} line={MAIN_LINE_REFS.get('override_entry_guard')}",
    )
    add(
        "override_notes_gate",
        _row_is_all_ack_to_ess(row),
        f"line={MAIN_LINE_REFS.get('override_notes_gate')}, notes={notes_l}",
    )

    if not exact_plan:
        add("override_exact_plan_available", False, "no exact plan object")
        return trace

    add(
        "override_exact_plan_available",
        bool(exact_plan.get("exists")),
        f"lane_kind={exact_plan.get('lane_kind')}, scope={exact_plan.get('scope')}",
    )
    add(
        "override_group_size_gate",
        (exact_plan.get("group_size") or 0) >= 2,
        f"line={MAIN_LINE_REFS.get('override_group_size_gate')}, group_size={exact_plan.get('group_size')}",
    )
    add(
        "override_slot_gate",
        exact_plan.get("slot_index") is not None,
        f"line={MAIN_LINE_REFS.get('override_slot_gate')}, slot_index={exact_plan.get('slot_index')}",
    )
    add(
        "override_consultant_ess_pool_len",
        len(exact_plan.get("consultant_ess_pool") or []),
        f"group_size={exact_plan.get('group_size')}",
    )
    add(
        "override_ess_pool_len",
        len(exact_plan.get("ess_pool") or []),
        f"group_size={exact_plan.get('group_size')}",
    )
    add(
        "override_occurrence_acky_pool_len",
        len(exact_plan.get("occurrence_acky_pool") or []),
        f"group_size={exact_plan.get('group_size')}",
    )
    pool_len = max(
        len(exact_plan.get("consultant_ess_pool") or []),
        len(exact_plan.get("ess_pool") or []),
        len(exact_plan.get("occurrence_acky_pool") or []),
    )
    add(
        "override_pool_gate",
        pool_len >= (exact_plan.get("group_size") or 0),
        f"line={MAIN_LINE_REFS.get('override_pool_gate')}, chosen_pool_len={pool_len}, group_size={exact_plan.get('group_size')}",
    )
    pick_obj = exact_plan.get("pick")
    add(
        "override_pick_gate",
        bool(pick_obj),
        f"line={MAIN_LINE_REFS.get('override_pick_gate')}, pick={_fmt(pick_obj.sent_time) if pick_obj else '-'}",
    )
    return trace


def _runtime_style_override_plan(
    target_row: dict,
    family_rows: list[dict],
    output_by_line: dict[int, dict],
    all_emails: list[DebugEmail],
    ess_team: list[str],
):
    requester = _get_col(target_row, "Requester", "Consultant")
    subject_norm_value = _family_subject_norm(target_row)
    if not requester or not subject_norm_value or not _row_is_all_ack_to_ess(target_row):
        return None

    group_sorted = []
    group_requesters = set()
    for row in family_rows:
        if not _row_is_all_ack_to_ess(row):
            continue
        if not _subject_match(subject_norm_value, _family_subject_norm(row)):
            continue
        group_sorted.append(row)
        req = _get_col(row, "Requester", "Consultant")
        if req:
            group_requesters.add(req)

    group_sorted.sort(key=lambda r: int(r.get("_line", 10**9)))
    if len(group_sorted) < 2:
        return None

    target_line = int(target_row["_line"])
    slot_index = None
    for idx, row in enumerate(group_sorted):
        if int(row["_line"]) == target_line:
            slot_index = idx
            break
    if slot_index is None:
        return None

    requester_names = tuple(sorted(group_requesters))

    def _group_req_match(email: DebugEmail) -> bool:
        return _requester_match_any(email, list(requester_names))

    def _subject_match_override(email_subject: str) -> bool:
        e_norm = normalize_subject(email_subject or "")
        if not subject_norm_value or not e_norm:
            return False
        if subject_norm_value == e_norm:
            return True
        return subject_norm_value in e_norm or e_norm in subject_norm_value

    def _dedupe_group_ess_minutes(items: list[DebugEmail]) -> list[DebugEmail]:
        buckets: dict[datetime, list[DebugEmail]] = {}
        for item in items:
            if not item.sent_time:
                continue
            minute_key = _to_ist(item.sent_time).replace(second=0, microsecond=0)
            buckets.setdefault(minute_key, []).append(item)

        out = []
        for minute_key in sorted(buckets):
            bucket = buckets[minute_key]
            bucket.sort(
                key=lambda email_obj: (
                    0 if _group_req_match(email_obj) else 1,
                    0 if (_classify_reply_kind(email_obj).get("direct_resolution")) else 1,
                    0 if (_classify_reply_kind(email_obj).get("real_reply")) else 1,
                    _to_ist(email_obj.sent_time),
                )
            )
            out.append(bucket[0])
        return out

    consultant_ess_pool = []
    ess_pool = []
    occurrence_acky_pool = []
    for email in all_emails:
        if not email.sent_time:
            continue
        if not _subject_match_override(email.subject):
            continue
        if _is_real_ess_progression(email, ess_team):
            ess_pool.append(email)
            if _group_req_match(email):
                consultant_ess_pool.append(email)
        if _is_occurrence_acky_candidate(email, list(requester_names)):
            occurrence_acky_pool.append(email)

    consultant_ess_pool = _dedupe_group_ess_minutes(consultant_ess_pool)
    ess_pool = _dedupe_group_ess_minutes(ess_pool)
    occurrence_acky_pool = _dedupe_group_ess_minutes(occurrence_acky_pool)

    anchor_row = output_by_line.get(int(group_sorted[0]["_line"]), {})
    anchor_ack = _parse_datetime(_get_col(anchor_row, "Actual Response Date & Time"))
    anchor_ack_ist = _to_ist(anchor_ack) if anchor_ack else None

    def _same_month_pool(pool_in: list[DebugEmail]) -> list[DebugEmail]:
        if not anchor_ack_ist:
            return list(pool_in)
        same_month = []
        for email in pool_in:
            e_ist = _to_ist(email.sent_time) if email.sent_time else None
            if e_ist and e_ist.year == anchor_ack_ist.year and e_ist.month == anchor_ack_ist.month:
                same_month.append(email)
        if len(same_month) >= len(group_sorted):
            return same_month
        return list(pool_in)

    consultant_ess_pool = _same_month_pool(consultant_ess_pool)
    ess_pool = _same_month_pool(ess_pool)
    occurrence_acky_pool = _same_month_pool(occurrence_acky_pool)

    pool = None
    pool_name = "<none>"
    if len(consultant_ess_pool) >= len(group_sorted):
        pool = consultant_ess_pool
        pool_name = "consultant_ess_pool"
    elif len(ess_pool) >= len(group_sorted):
        pool = ess_pool
        pool_name = "ess_pool"
    elif len(occurrence_acky_pool) >= len(group_sorted):
        pool = occurrence_acky_pool
        pool_name = "occurrence_acky_pool"
    if not pool or len(pool) < len(group_sorted):
        return None

    pick = pool[slot_index]
    return {
        "group_sorted": group_sorted,
        "group_size": len(group_sorted),
        "slot_index": slot_index,
        "pick": pick,
        "pick_when": _fmt(pick.sent_time),
        "pool_name": pool_name,
        "pool": pool,
        "consultant_ess_pool": consultant_ess_pool,
        "ess_pool": ess_pool,
        "occurrence_acky_pool": occurrence_acky_pool,
    }


def _fmt_email_minutes(items: list[DebugEmail]) -> list[str]:
    out = []
    for item in items:
        out.append(_fmt(item.sent_time))
    return out


def _pool_slot_for_minute(pool: list[DebugEmail], minute_text: str) -> int | None:
    if not minute_text or minute_text == "-":
        return None
    for idx, item in enumerate(pool, start=1):
        if _fmt(item.sent_time) == minute_text:
            return idx
    return None


def _runtime_style_override_trace(
    row: dict,
    runtime_plan: dict | None,
    exact_plan: dict | None,
    duplicate_pick_lines: dict[str, list[int]],
) -> list[dict]:
    trace = []

    def add(step: str, result, detail: str = ""):
        trace.append({
            "step": step,
            "result": result,
            "detail": detail,
        })

    add(
        "runtime_override_entry",
        _row_is_all_ack_to_ess(row),
        f"line={row.get('_line')}, requester={_get_col(row, 'Requester', 'Consultant')}",
    )
    if not runtime_plan:
        add("runtime_override_plan_exists", False, "no runtime-style override plan")
        return trace

    pick_when = runtime_plan.get("pick_when") or "-"
    exact_pick_when = _fmt((exact_plan or {}).get("pick").sent_time if (exact_plan or {}).get("pick") else None)
    pool_name = runtime_plan.get("pool_name") or "<none>"
    consultant_pool = _fmt_email_minutes(runtime_plan.get("consultant_ess_pool") or [])
    ess_pool = _fmt_email_minutes(runtime_plan.get("ess_pool") or [])
    occurrence_acky_pool = _fmt_email_minutes(runtime_plan.get("occurrence_acky_pool") or [])
    chosen_pool = _fmt_email_minutes(runtime_plan.get("pool") or [])
    duplicate_lines = duplicate_pick_lines.get(pick_when, [])

    add(
        "runtime_override_plan_exists",
        True,
        (
            f"group_size={runtime_plan.get('group_size')} "
            f"slot={int(runtime_plan.get('slot_index', -1)) + 1} "
            f"pool_name={pool_name}"
        ),
    )
    add(
        "runtime_override_group_lines",
        True,
        "lines=" + ",".join(str(int(group_row.get('_line', -1))) for group_row in runtime_plan.get("group_sorted") or []),
    )
    add(
        "runtime_override_consultant_pool",
        len(consultant_pool),
        "minutes=" + (" | ".join(consultant_pool) if consultant_pool else "-"),
    )
    add(
        "runtime_override_ess_pool",
        len(ess_pool),
        "minutes=" + (" | ".join(ess_pool) if ess_pool else "-"),
    )
    add(
        "runtime_override_occurrence_acky_pool",
        len(occurrence_acky_pool),
        "minutes=" + (" | ".join(occurrence_acky_pool) if occurrence_acky_pool else "-"),
    )
    add(
        "runtime_override_chosen_pool",
        len(chosen_pool),
        "minutes=" + (" | ".join(chosen_pool) if chosen_pool else "-"),
    )
    add(
        "runtime_override_pick",
        pick_when,
        f"exact_pick={exact_pick_when}",
    )
    add(
        "runtime_override_matches_exact",
        pick_when == exact_pick_when,
        f"runtime_pick={pick_when}, exact_pick={exact_pick_when}",
    )
    add(
        "runtime_override_unique_pick",
        len(duplicate_lines) <= 1,
        (
            f"pick={pick_when}"
            if len(duplicate_lines) <= 1
            else f"pick={pick_when} reused_by_lines={','.join(str(line) for line in duplicate_lines)}"
        ),
    )
    return trace


def _strict_ess_pass_trace(
    row: dict,
    output_by_line: dict[int, dict],
    exact_plan: dict | None,
    runtime_plan: dict | None,
) -> list[dict]:
    trace = []

    def add(step: str, result, detail: str = ""):
        trace.append({
            "step": step,
            "result": result,
            "detail": detail,
        })

    line = int(row["_line"])
    notes_l = _get_col(row, "Notes").lower()
    current_triplet = _fmt_triplet(output_by_line.get(line, {}))
    current_single = current_triplet.split(" / ")[0] if " / " in current_triplet else current_triplet
    exact_pick = _fmt((exact_plan or {}).get("pick").sent_time if (exact_plan or {}).get("pick") else None)
    chosen_pool = (runtime_plan or {}).get("pool") or []
    exact_slot = int((runtime_plan or {}).get("slot_index", -1)) + 1 if runtime_plan else None
    current_slot = _pool_slot_for_minute(chosen_pool, current_single)
    group_sorted = (runtime_plan or {}).get("group_sorted") or []

    prev_line = None
    prev_output = "-"
    if group_sorted:
        for idx, group_row in enumerate(group_sorted):
            if int(group_row.get("_line", -1)) != line:
                continue
            if idx > 0:
                prev_line = int(group_sorted[idx - 1].get("_line", -1))
                prev_output = _fmt_triplet(output_by_line.get(prev_line, {}))
            break

    add(
        "strict_pass_authoritative_call_site",
        MAIN_LINE_REFS.get("ess_only_strict_authoritative_call"),
        (
            f"call_line={MAIN_LINE_REFS.get('ess_only_strict_authoritative_call')}, "
            f"manual_pick_branch={MAIN_LINE_REFS.get('ess_only_strict_manual_occ_pick_branch')}, "
            f"manual_notes_write={MAIN_LINE_REFS.get('ess_only_strict_manual_notes_write')}"
        ),
    )
    add(
        "strict_pass_row_is_all_ack_ess",
        _row_is_all_ack_to_ess(row),
        f"notes_has_all_ack={'requester span(all-ack->ess)' in notes_l}",
    )
    add(
        "strict_pass_exact_pick",
        exact_pick != "-",
        f"exact_pick={exact_pick}, current={current_single}",
    )
    add(
        "strict_pass_pool_slot_alignment",
        current_slot == exact_slot,
        f"current_slot={current_slot}, exact_slot={exact_slot}",
    )
    add(
        "strict_pass_reuses_earlier_slot",
        bool(current_slot is not None and exact_slot is not None and current_slot < exact_slot),
        f"current_slot={current_slot}, exact_slot={exact_slot}",
    )
    add(
        "strict_pass_matches_previous_group_output",
        bool(prev_line is not None and current_triplet == prev_output),
        f"prev_line={prev_line}, prev_output={prev_output}",
    )
    add(
        "strict_pass_local_branches_should_target_exact_pick",
        bool(exact_pick != "-" and current_single != exact_pick),
        (
            "both the authoritative strict branch and the manual occ-pick branch "
            f"would target exact_pick={exact_pick}; current remains {current_single}"
        ),
    )
    return trace


def _main_flow_hazard_trace(
    row: dict,
    output_by_line: dict[int, dict],
    exact_plan: dict | None,
    runtime_plan: dict | None,
) -> list[dict]:
    trace = []

    def add(step: str, result, detail: str = ""):
        trace.append({
            "step": step,
            "result": result,
            "detail": detail,
        })

    notes_l = _get_col(row, "Notes").lower()
    current_triplet = _fmt_triplet(output_by_line.get(int(row["_line"]), {}))
    current_single = current_triplet.split(" / ")[0] if " / " in current_triplet else current_triplet
    exact_pick = _fmt((exact_plan or {}).get("pick").sent_time if (exact_plan or {}).get("pick") else None)
    runtime_pick = runtime_plan.get("pick_when") if runtime_plan else "-"

    add(
        "strict_pass_shared_triplet_call_site",
        MAIN_LINE_REFS.get("ess_only_strict_shared_triplet_call"),
        (
            f"call_line={MAIN_LINE_REFS.get('ess_only_strict_shared_triplet_call')}, "
            f"helper_line={MAIN_LINE_REFS.get('apply_shared_occurrence_triplet_def')}"
        ),
    )
    add(
        "row_has_ess_guard_and_lock",
        ("esscontinuationguard[" in notes_l) and ("occurrencelocked" in notes_l),
        f"notes_has_guard={'esscontinuationguard[' in notes_l}, notes_has_lock={'occurrencelocked' in notes_l}",
    )
    add(
        "current_output_matches_exact_pick",
        current_single == exact_pick,
        f"current={current_single}, exact_pick={exact_pick}",
    )
    add(
        "current_output_matches_runtime_override_pick",
        current_single == runtime_pick,
        f"current={current_single}, runtime_pick={runtime_pick}",
    )
    add(
        "shared_decision_carryover_risk",
        bool(
            "requester span(all-ack->ess)" in notes_l
            and "esscontinuationguard[" in notes_l
            and exact_pick != "-"
            and current_single != exact_pick
        ),
        (
            "strict ESS continuation path can apply state['shared_decision'] before final pass; "
            f"current={current_single}, exact_pick={exact_pick}"
        ),
    )
    return trace


def _simulate_family(
    family_subject: str,
    family_rows: list[dict],
    output_by_line: dict[int, dict],
    all_emails: list[DebugEmail],
    ess_team: list[str],
) -> None:
    reply_pool, ack_pool, direct_pool, consultant_ess_pool, ess_pool, occurrence_acky_pool = _collect_family_pools(
        family_subject,
        family_rows,
        all_emails,
        ess_team,
    )

    sorted_rows = sorted(family_rows, key=lambda r: int(r.get("_line", 10**9)))
    special_rows = [row for row in sorted_rows if _row_is_all_ack_to_ess(row)]
    special_slot_by_line = {int(row["_line"]): idx for idx, row in enumerate(special_rows)}
    normal_slot_by_line = {int(row["_line"]): idx for idx, row in enumerate(sorted_rows)}
    usable_non_ess_lane = (len(reply_pool) >= len(sorted_rows) and len(ack_pool) >= len(sorted_rows)) or (len(direct_pool) >= len(sorted_rows))

    print("=" * 100)
    print(f"family_subject_norm={family_subject}")
    print("SIMULATED OCCURRENCE LOCK")
    print(
        f"reply_pool={len(reply_pool)} ack_pool={len(ack_pool)} direct_pool={len(direct_pool)} "
        f"consultant_ess_pool={len(consultant_ess_pool)} ess_pool={len(ess_pool)} "
        f"occurrence_acky_pool={len(occurrence_acky_pool)}"
    )
    print(f"special_rows={len(special_rows)} total_rows={len(sorted_rows)} usable_non_ess_lane={usable_non_ess_lane}")
    print("-" * 100)

    for row in sorted_rows:
        line = int(row["_line"])
        out_row = output_by_line.get(line, {})
        requester = _get_col(row, "Requester", "Consultant")
        notes_l = _get_col(row, "Notes").lower()
        current_triplet = _fmt_triplet(out_row) if out_row else "- / - / -"

        scope = "requester"
        slot_index = normal_slot_by_line.get(line, 0)
        expected_kind = "reply"
        expected_email = None
        would_lock = False
        blocker = ""

        if _row_is_all_ack_to_ess(row):
            scope = "subject_wide_ess"
            slot_index = special_slot_by_line.get(line, 0)
            if usable_non_ess_lane:
                expected_kind = "reply_or_direct"
                if slot_index < len(direct_pool):
                    expected_email = direct_pool[slot_index]
                elif slot_index < len(reply_pool):
                    expected_email = reply_pool[slot_index]
                else:
                    blocker = "non_ess_slot_missing"
            elif slot_index < len(consultant_ess_pool):
                expected_kind = "ess_over_ess"
                expected_email = consultant_ess_pool[slot_index]
            elif slot_index < len(ess_pool):
                expected_kind = "ess_over_ess_generic"
                expected_email = ess_pool[slot_index]
            elif slot_index < len(occurrence_acky_pool):
                expected_kind = "ess_acky_sequence"
                expected_email = occurrence_acky_pool[slot_index]
            else:
                blocker = "no_special_lane"
        else:
            if slot_index < len(reply_pool):
                expected_email = reply_pool[slot_index]
                expected_kind = "reply"
            elif slot_index < len(direct_pool):
                expected_email = direct_pool[slot_index]
                expected_kind = "direct_resolution"
            elif slot_index < len(ack_pool):
                expected_email = ack_pool[slot_index]
                expected_kind = "ess_acky"
            elif slot_index < len(consultant_ess_pool):
                expected_email = consultant_ess_pool[slot_index]
                expected_kind = "consultant_ess_fallback"
            elif slot_index < len(ess_pool):
                expected_email = ess_pool[slot_index]
                expected_kind = "ess_fallback"
            else:
                blocker = "no_lane"

        expected_when = _fmt(expected_email.sent_time) if expected_email else "-"
        expected_triplet = "-"
        if expected_email and (expected_kind.startswith("ess_over_ess") or expected_kind == "ess_acky_sequence"):
            expected_triplet = f"{expected_when} / {expected_when} / {expected_when}"
            would_lock = "requester span(all-ack->ess)" in notes_l
        elif expected_email:
            out_triplet = _row_triplet(out_row) if out_row else (None, None, None)
            lane_when = _to_ist(expected_email.sent_time).replace(second=0, microsecond=0)
            if out_triplet[1] and _to_ist(out_triplet[1]).replace(second=0, microsecond=0) == lane_when:
                would_lock = True
            elif out_triplet[2] and _to_ist(out_triplet[2]).replace(second=0, microsecond=0) == lane_when:
                would_lock = True

        print(
            f"line={line} requester={requester} scope={scope} slot={slot_index + 1} "
            f"expected_kind={expected_kind} would_lock={would_lock}"
        )
        print(f"  current_triplet={current_triplet}")
        print(f"  expected_pick={expected_when}")
        if expected_triplet != "-":
            print(f"  expected_triplet={expected_triplet}")
        if blocker:
            print(f"  blocker={blocker}")

    print("-" * 100)
    print("LOCK PASS TRACE")
    for row in sorted_rows:
        line = int(row["_line"])
        out_row = output_by_line.get(line, {})
        requester = _get_col(row, "Requester", "Consultant")
        notes_l = _get_col(row, "Notes").lower()
        current_triplet = _fmt_triplet(out_row) if out_row else "- / - / -"
        is_all_ack = _row_is_all_ack_to_ess(row)
        slot_index = normal_slot_by_line.get(line, 0)
        scope = "requester"
        shared_occ_exists = True
        shared_lane_kind = "reply"
        shared_pick = None
        reason = ""

        if is_all_ack:
            scope = "subject_wide_ess"
            slot_index = special_slot_by_line.get(line, 0)
            if usable_non_ess_lane:
                shared_lane_kind = "reply"
                if slot_index < len(direct_pool):
                    shared_pick = direct_pool[slot_index]
                elif slot_index < len(reply_pool):
                    shared_pick = reply_pool[slot_index]
                else:
                    shared_occ_exists = False
                    reason = "no_non_ess_slot_for_special_row"
            elif slot_index < len(consultant_ess_pool):
                shared_lane_kind = "ess_over_ess"
                shared_pick = consultant_ess_pool[slot_index]
            elif slot_index < len(ess_pool):
                shared_lane_kind = "ess_over_ess"
                shared_pick = ess_pool[slot_index]
            elif slot_index < len(occurrence_acky_pool):
                shared_lane_kind = "ess_acky_sequence"
                shared_pick = occurrence_acky_pool[slot_index]
            else:
                shared_occ_exists = False
                reason = "no_special_lane"
        else:
            if slot_index < len(reply_pool):
                shared_pick = reply_pool[slot_index]
                shared_lane_kind = "reply"
            elif slot_index < len(ack_pool):
                shared_pick = ack_pool[slot_index]
                shared_lane_kind = "ess_acky"
            elif slot_index < len(direct_pool):
                shared_pick = direct_pool[slot_index]
                shared_lane_kind = "reply"
            elif slot_index < len(consultant_ess_pool):
                shared_pick = consultant_ess_pool[slot_index]
                shared_lane_kind = "ess_over_ess"
            elif slot_index < len(ess_pool):
                shared_pick = ess_pool[slot_index]
                shared_lane_kind = "ess_over_ess"
            else:
                shared_occ_exists = False
                reason = "no_lane"

        shared_pick_when = _fmt(shared_pick.sent_time) if shared_pick else "-"
        notes_has_all_ack = "requester span(all-ack->ess)" in notes_l
        authoritative_apply = bool(
            shared_occ_exists
            and _lane_is_authoritative_occurrence(shared_lane_kind)
            and notes_has_all_ack
            and shared_pick is not None
        )

        lock_candidate_triplet = "-"
        if authoritative_apply:
            lock_candidate_triplet = (
                f"{shared_pick_when} / {shared_pick_when} / {shared_pick_when}"
            )
        else:
            out_triplet = _row_triplet(out_row) if out_row else (None, None, None)
            if shared_pick and (out_triplet[1] or out_triplet[2]):
                lane_when = _to_ist(shared_pick.sent_time).replace(second=0, microsecond=0)
                ack_match = bool(
                    out_triplet[1]
                    and _to_ist(out_triplet[1]).replace(second=0, microsecond=0) == lane_when
                )
                resolved_match = bool(
                    out_triplet[2]
                    and _to_ist(out_triplet[2]).replace(second=0, microsecond=0) == lane_when
                )
                if ack_match or resolved_match:
                    lock_candidate_triplet = current_triplet

        runtime_should_lock = lock_candidate_triplet != "-"
        runtime_mismatch = ""
        if authoritative_apply and current_triplet != lock_candidate_triplet:
            runtime_mismatch = "runtime_should_have_collapsed_to_occurrence_triplet"
        elif runtime_should_lock and "occurrencelocked" not in notes_l:
            runtime_mismatch = "runtime_should_have_marked_occurrence_locked"
        elif not shared_occ_exists:
            runtime_mismatch = reason

        print(
            f"line={line} requester={requester} scope={scope} shared_occ_exists={shared_occ_exists} "
            f"lane_kind={shared_lane_kind} notes_all_ack={notes_has_all_ack}"
        )
        print(f"  current_triplet={current_triplet}")
        print(f"  shared_pick={shared_pick_when}")
        print(f"  authoritative_apply={authoritative_apply}")
        if lock_candidate_triplet != "-":
            print(f"  lock_candidate_triplet={lock_candidate_triplet}")
        print(f"  runtime_should_lock={runtime_should_lock}")
        if runtime_mismatch:
            print(f"  runtime_mismatch={runtime_mismatch}")

    print("-" * 100)
    print("EXACT PLANNER TRACE")
    exact_plans_by_line: dict[int, dict | None] = {}
    runtime_override_by_line: dict[int, dict | None] = {}
    for row in sorted_rows:
        line = int(row["_line"])
        current_triplet = _fmt_triplet(output_by_line.get(line, {}))
        notes_l = _get_col(row, "Notes").lower()
        plan = _exact_shared_occurrence_plan(
            row,
            family_rows,
            output_by_line,
            all_emails,
            ess_team,
        )
        exact_plans_by_line[line] = plan
        runtime_override_by_line[line] = _runtime_style_override_plan(
            row,
            family_rows,
            output_by_line,
            all_emails,
            ess_team,
        )
        print(f"line={line} requester={_get_col(row, 'Requester', 'Consultant')}")
        print(f"  current_triplet={current_triplet}")
        if not plan or not plan.get('exists'):
            reason = (plan or {}).get("reason", "no_plan")
            print(f"  exact_plan_exists=False")
            print(f"  exact_plan_reason={reason}")
            continue
        pick_when = _fmt(plan.get("pick").sent_time if plan.get("pick") else None)
        lane_kind = plan.get("lane_kind")
        authoritative_apply = bool(
            _lane_is_authoritative_occurrence(lane_kind)
            and "requester span(all-ack->ess)" in notes_l
            and plan.get("pick") is not None
        )
        expected_triplet = (
            f"{pick_when} / {pick_when} / {pick_when}"
            if authoritative_apply
            else "-"
        )
        print(f"  exact_plan_exists=True")
        print(
            f"  scope={plan.get('scope')} group_size={plan.get('group_size')} "
            f"slot={plan.get('slot_index', 0) + 1} allow_acky={plan.get('allow_acky')}"
        )
        print(
            f"  lane_kind={lane_kind} pick={pick_when} "
            f"reply_pool={len(plan.get('reply_pool') or [])} "
            f"acky_pool={len(plan.get('acky_pool') or [])} "
            f"consultant_ess_pool={len(plan.get('consultant_ess_pool') or [])} "
            f"ess_pool={len(plan.get('ess_pool') or [])} "
            f"occurrence_acky_pool={len(plan.get('occurrence_acky_pool') or [])}"
        )
        print(
            f"  strong_non_ess_live={plan.get('strong_non_ess_live')} "
            f"notes_all_ack={'requester span(all-ack->ess)' in notes_l}"
        )
        print(f"  authoritative_apply={authoritative_apply}")
        if expected_triplet != "-":
            print(f"  expected_triplet={expected_triplet}")
            if current_triplet != expected_triplet:
                print("  exact_runtime_mismatch=planner_says_row_should_be_collapsed")

    print("-" * 100)
    print("PROPOSED LOCKED RESULTS")
    flow_summaries = []
    for row in sorted_rows:
        line = int(row["_line"])
        requester = _get_col(row, "Requester", "Consultant")
        desc = _get_col(row, "Description")
        current_triplet = _fmt_triplet(output_by_line.get(line, {}))
        exact_plan = exact_plans_by_line.get(line)
        proposed_triplet = _proposed_locked_triplet(
            row,
            output_by_line,
            exact_plan,
        )
        flow_summaries.append(
            _build_row_flow_summary(
                row,
                output_by_line,
                exact_plan,
            )
        )
        changed = current_triplet != proposed_triplet
        print(
            f"line={line} requester={requester} changed={changed}"
        )
        print(f"  desc={desc}")
        print(f"  current={current_triplet}")
        print(f"  proposed={proposed_triplet}")

    print("-" * 100)
    print("FLOW SUMMARY JSON")
    print(json.dumps(flow_summaries, indent=2))

    print("-" * 100)
    print("OVERRIDE FUNCTION TRACE")
    for row in sorted_rows:
        line = int(row["_line"])
        print(f"line={line} requester={_get_col(row, 'Requester', 'Consultant')}")
        for item in _override_function_trace(
            row,
            exact_plans_by_line.get(line),
        ):
            print(f"  step={item['step']} | result={item['result']}")
            if item["detail"]:
                print(f"    detail={item['detail']}")

    print("-" * 100)
    print("RUNTIME-STYLE OVERRIDE TRACE")
    duplicate_pick_lines: dict[str, list[int]] = {}
    for line, runtime_plan in runtime_override_by_line.items():
        if not runtime_plan:
            continue
        pick_when = runtime_plan.get("pick_when") or "-"
        duplicate_pick_lines.setdefault(pick_when, []).append(line)
    for row in sorted_rows:
        line = int(row["_line"])
        print(f"line={line} requester={_get_col(row, 'Requester', 'Consultant')}")
        for item in _runtime_style_override_trace(
            row,
            runtime_override_by_line.get(line),
            exact_plans_by_line.get(line),
            duplicate_pick_lines,
        ):
            print(f"  step={item['step']} | result={item['result']}")
            if item["detail"]:
                print(f"    detail={item['detail']}")

    print("-" * 100)
    print("MAIN FLOW HAZARD TRACE")
    for row in sorted_rows:
        line = int(row["_line"])
        print(f"line={line} requester={_get_col(row, 'Requester', 'Consultant')}")
        for item in _main_flow_hazard_trace(
            row,
            output_by_line,
            exact_plans_by_line.get(line),
            runtime_override_by_line.get(line),
        ):
            print(f"  step={item['step']} | result={item['result']}")
            if item["detail"]:
                print(f"    detail={item['detail']}")

    print("-" * 100)
    print("STRICT ESS PASS TRACE")
    for row in sorted_rows:
        line = int(row["_line"])
        print(f"line={line} requester={_get_col(row, 'Requester', 'Consultant')}")
        for item in _strict_ess_pass_trace(
            row,
            output_by_line,
            exact_plans_by_line.get(line),
            runtime_override_by_line.get(line),
        ):
            print(f"  step={item['step']} | result={item['result']}")
            if item["detail"]:
                print(f"    detail={item['detail']}")

    print("-" * 100)
    print("RUNTIME MIRROR TRACE")
    for row in sorted_rows:
        line = int(row["_line"])
        print(f"line={line} requester={_get_col(row, 'Requester', 'Consultant')}")
        for item in _runtime_mirror_trace(
            row,
            output_by_line,
            exact_plans_by_line.get(line),
        ):
            print(
                f"  step={item['step']} | result={item['result']}"
            )
            if item["detail"]:
                print(f"    detail={item['detail']}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Simulate occurrence lock behavior outside main runtime")
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
    all_emails = list(_iter_emails(Path(args.eml_dir)))

    if not debug_rows:
        print("No matching debug rows found.")
        return

    families = _cluster_family_rows(debug_rows)
    for family_subject, family_rows in families.items():
        _simulate_family(
            family_subject,
            family_rows,
            output_by_line,
            all_emails,
            ess_team,
        )


if __name__ == "__main__":
    main()
