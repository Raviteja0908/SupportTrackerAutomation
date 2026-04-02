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


def _print_slot_rows(rows: list[dict]) -> None:
    print("family rows by slot:")
    for slot, row in enumerate(sorted(rows, key=lambda r: int(r.get("_line", 10**9))), start=1):
        req = _get_col(row, "Requester", "Consultant")
        desc = _get_col(row, "Description")
        notes = _get_col(row, "Notes")
        print(
            f"  slot={slot} | line={row.get('_line')} | requester={req} | triplet={_fmt_triplet(row)}"
        )
        print(f"    desc={desc}")
        if notes:
            print(f"    notes={notes}")


def _trace_row_flow(
    family_rows: list[dict],
    output_by_line: dict[int, dict],
    consultant_ess_pool: list[DebugEmail],
    ess_pool: list[DebugEmail],
    reply_pool: list[DebugEmail],
    ack_pool: list[DebugEmail],
    direct_pool: list[DebugEmail],
) -> None:
    print("TRACE FLOW")
    sorted_rows = sorted(family_rows, key=lambda r: int(r.get("_line", 10**9)))
    special_rows = [row for row in sorted_rows if _row_is_all_ack_to_ess(row)]
    special_slot_by_line = {int(row["_line"]): idx for idx, row in enumerate(special_rows)}
    normal_slot_by_line = {int(row["_line"]): idx for idx, row in enumerate(sorted_rows)}

    usable_non_ess_lane = (len(reply_pool) >= len(sorted_rows) and len(ack_pool) >= len(sorted_rows)) or (len(direct_pool) >= len(sorted_rows))
    usable_consultant_ess_lane_special = len(consultant_ess_pool) >= len(special_rows) if special_rows else False

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
        slot_idx = normal_slot_by_line.get(line, 0)
        expected_kind = "reply"
        expected_pick = None
        blocker = ""

        if is_all_ack and special_rows:
            scope = "subject_wide_ess"
            slot_idx = special_slot_by_line.get(line, 0)
            if usable_non_ess_lane:
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
            if slot_idx < len(reply_pool):
                expected_kind = "reply"
                expected_pick = reply_pool[slot_idx]
            elif slot_idx < len(direct_pool):
                expected_kind = "direct_resolution"
                expected_pick = direct_pool[slot_idx]
            elif slot_idx < len(ack_pool):
                expected_kind = "ess_acky"
                expected_pick = ack_pool[slot_idx]
            elif slot_idx < len(ess_pool):
                expected_kind = "ess_over_ess_fallback"
                expected_pick = ess_pool[slot_idx]
            else:
                blocker = "no_slot_pick"

        expected_when = _fmt(expected_pick.sent_time) if expected_pick else "-"
        expected_triplet = (
            f"{expected_when} / {expected_when} / {expected_when}"
            if expected_pick and expected_kind.startswith("ess_over_ess")
            else "-"
        )
        has_final_guard = "esscontinuationguard[allthreestrictessonly]" in notes_l
        likely_drop = ""
        if expected_pick and expected_kind.startswith("ess_over_ess") and not has_final_guard:
            likely_drop = "final_ess_guard_not_applied"
        elif blocker:
            likely_drop = blocker

        print(
            f"  line={line} | requester={req} | occ={is_occ} | all_ack_to_ess={is_all_ack} | "
            f"scope={scope} | slot={slot_idx + 1} | expected_kind={expected_kind}"
        )
        print(f"    current_triplet={current_triplet}")
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
        print(f"family_subject_norm={subject_norm}")
        print(f"family_requesters={', '.join(requesters) if requesters else '<none>'}")
        print("-" * 100)
        _print_slot_rows(family_rows)
        print("-" * 100)

        matched_emails = []
        for email in all_emails:
            if not email.sent_time:
                continue
            if not _subject_match(subject_norm, email.subject):
                continue
            matched_emails.append(email)
        matched_emails.sort(key=lambda x: x.sent_time or datetime.max)
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
            output_by_line,
            consultant_ess_pool,
            ess_pool,
            reply_pool,
            ack_pool,
            direct_pool,
        )


if __name__ == "__main__":
    main()
