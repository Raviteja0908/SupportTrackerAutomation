import argparse
from collections import defaultdict
from datetime import timedelta
from pathlib import Path

from test_request_anchor_flow import (
    _best_before,
    _classify_reply_kind,
    _family_subject_norm,
    _find_family_rows,
    _fmt,
    _get_col,
    _id_like_tokens,
    _is_ess_sender,
    _iter_emails,
    _load_csv_rows,
    _match_tokens,
    _requester_match_any,
    _row_subject_match,
    _system_like_sender,
    _to_ist,
    _extract_quoted_blocks_relaxed,
    _extract_quoted_blocks_with_subject,
    _ess_name_only,
    _minute_dedupe,
    _read_json_list,
    normalize_subject,
)


QUOTED_TAGS = (
    "quotedrequestonly",
    "quotedrequestonlynopair",
    "quotedrequestonlypreservedliveack",
    "quotedrequestonlydirectreply",
    "quotedrequestonlyhybridliveack",
    "quotedrequestonlyrawemlfallback",
)


def _notes_l(row: dict) -> str:
    return (_get_col(row, "Notes") or "").lower()


def _quoted_family_rows(debug_rows: list[dict], contains_filters: list[str]) -> dict[str, list[dict]]:
    families = {}
    for family_key, rows in _find_family_rows(debug_rows).items():
        family_notes = " ; ".join(_notes_l(row) for row in rows)
        if not any(tag in family_notes for tag in QUOTED_TAGS):
            continue
        if contains_filters:
            hay = " || ".join(_get_col(row, "Description") for row in rows).lower()
            if not any(token.lower() in hay for token in contains_filters):
                continue
        families[family_key] = rows
    return families


def _analyze_family(family_rows: list[dict], matched_emails, ess_team: list[str]) -> dict:
    subject_norm = _family_subject_norm(family_rows[0]) if family_rows else ""
    row_tokens = _match_tokens(subject_norm)
    row_id_tokens = _id_like_tokens(subject_norm)
    requesters = sorted(
        {
            _get_col(row, "Requester", "Consultant")
            for row in family_rows
            if _get_col(row, "Requester", "Consultant")
        }
    )

    live_requests = []
    quoted_requests = []
    ack_candidates = []
    reply_candidates = []
    sibling_subjects = set()

    for email in matched_emails:
        email_ist = _to_ist(email.sent_time) if email.sent_time else None
        if not email_ist:
            continue
        email_norm = normalize_subject(email.subject or "")
        if email_norm and email_norm != subject_norm:
            sibling_subjects.add(email.subject)

        cls = _classify_reply_kind(email)
        is_ess = _is_ess_sender(email, ess_team)
        req_match = _requester_match_any(email, requesters)

        if (not is_ess) and (not _system_like_sender(email)):
            live_requests.append((email_ist, email))
        if is_ess and (cls.get("ack_like") or cls.get("explicit_ack") or cls.get("short_ess_ack") or (not req_match)):
            ack_candidates.append((email_ist, email))
        if is_ess and (not cls.get("ack_like")) and (not cls.get("explicit_ack")) and (not cls.get("short_ess_ack")) and (not cls.get("thanks_info")) and (not cls.get("nonfinal_followup")):
            reply_candidates.append((email_ist, email))

        quoted_blocks = _extract_quoted_blocks_with_subject(email) or _extract_quoted_blocks_relaxed(email)
        for from_line, quoted_ist, quoted_subj in quoted_blocks:
            if not quoted_ist or quoted_ist >= email_ist:
                continue
            if (email_ist - quoted_ist) > timedelta(hours=48):
                continue
            if quoted_subj and not _row_subject_match(subject_norm, row_tokens, row_id_tokens, quoted_subj):
                continue
            addr_hits = []
            if from_line:
                import re
                addr_hits = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", from_line, flags=re.I)
            emails_l = [addr.lower() for addr in addr_hits]
            domains_l = [addr.split("@", 1)[-1] for addr in emails_l if "@" in addr]
            if emails_l:
                is_quoted_ess = any(_is_ess_sender(type(email)("", "", addr, None, "", "", Path()), ess_team) for addr in emails_l)
            else:
                is_quoted_ess = _ess_name_only(from_line, ess_team)
            if is_quoted_ess or any(domain.endswith("invenio-solutions.com") for domain in domains_l):
                continue
            quoted_requests.append((quoted_ist, (email, from_line, quoted_subj)))

    live_requests = _minute_dedupe(live_requests)
    quoted_requests = _minute_dedupe(quoted_requests)
    ack_candidates = _minute_dedupe(ack_candidates)
    reply_candidates = _minute_dedupe(reply_candidates)

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
            (
                req_pick[0].replace(second=0, microsecond=0),
                ack_ist.replace(second=0, microsecond=0),
                resolved_pick[0].replace(second=0, microsecond=0),
                req_kind,
            )
        )

    unique_episodes = []
    seen = set()
    for episode in episodes:
        key = episode[:3]
        if key in seen:
            continue
        seen.add(key)
        unique_episodes.append(episode)

    notes_l = " ; ".join(_notes_l(row) for row in family_rows)
    root_causes = []
    if not live_requests:
        root_causes.append("no_live_request")
    if not quoted_requests:
        root_causes.append("no_quoted_request")
    if not ack_candidates:
        root_causes.append("no_ack_candidate")
    if not reply_candidates:
        root_causes.append("no_reply_candidate")
    if not unique_episodes:
        root_causes.append("no_coherent_episode")
    if sibling_subjects:
        root_causes.append("sibling_variant_candidates_present")
    if "quotedrequestonlynopair" in notes_l:
        root_causes.append("quoted_no_pair_lane")
    if "quotedrequestonlypreservedliveack" in notes_l:
        root_causes.append("preserved_live_ack_lane")
    if "requester span(all-ack->ess)" in notes_l:
        root_causes.append("all_ack_requester_span_lane")
    if "riskyfallbackvalidator[direct-reply]" in notes_l:
        root_causes.append("direct_reply_fallback_lane")
    if len(unique_episodes) == 1 and len(family_rows) > 1:
        root_causes.append("single_episode_multi_row_family")

    return {
        "subject_norm": subject_norm,
        "requesters": requesters,
        "family_rows": family_rows,
        "live_requests": live_requests,
        "quoted_requests": quoted_requests,
        "ack_candidates": ack_candidates,
        "reply_candidates": reply_candidates,
        "unique_episodes": unique_episodes,
        "sibling_subjects": sorted(sibling_subjects),
        "root_causes": root_causes,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit all quoted-family rows and show likely root causes for drift.")
    parser.add_argument("--debug-csv", required=True, help="debug_subjects CSV path")
    parser.add_argument("--output-csv", required=False, help="automation_output CSV path (reserved)")
    parser.add_argument("--eml-dir", required=True, help="Directory containing exported EML files")
    parser.add_argument("--ess-team", default=str(Path("config") / "ess_team.json"), help="ESS team JSON path")
    parser.add_argument("--contains", action="append", default=[], help="Optional description substring filter")
    parser.add_argument("--limit", type=int, default=20, help="Max family count to print")
    args = parser.parse_args()

    debug_rows = _load_csv_rows(Path(args.debug_csv))
    ess_team = _read_json_list(Path(args.ess_team))
    families = _quoted_family_rows(debug_rows, args.contains)
    all_emails = list(_iter_emails(Path(args.eml_dir)))

    printed = 0
    print(f"quoted_family_count={len(families)}")
    for family_key, family_rows in families.items():
        if printed >= args.limit:
            break
        subject_norm = _family_subject_norm(family_rows[0]) if family_rows else ""
        row_tokens = _match_tokens(subject_norm)
        row_id_tokens = _id_like_tokens(subject_norm)
        matched = [
            email
            for email in all_emails
            if _row_subject_match(subject_norm, row_tokens, row_id_tokens, email.subject)
        ]
        analysis = _analyze_family(family_rows, matched, ess_team)

        print("=" * 100)
        print(f"family_subject_norm={analysis['subject_norm']}")
        print(f"family_requesters={', '.join(analysis['requesters']) or '-'}")
        print(f"family_row_count={len(analysis['family_rows'])}")
        print(f"matched_email_count={len(matched)}")
        print(f"root_causes={', '.join(analysis['root_causes']) or 'none'}")
        if analysis["sibling_subjects"]:
            print("sibling_subjects:")
            for subj in analysis["sibling_subjects"][:5]:
                print(f"  - {subj}")
        print("rows:")
        for row in analysis["family_rows"]:
            print(
                f"  line={row.get('_line')} | requester={_get_col(row, 'Requester', 'Consultant')} | "
                f"current={_get_col(row, 'Created Date & Time') or '-'} / "
                f"{_get_col(row, 'Actual Response Date & Time') or '-'} / "
                f"{_get_col(row, 'Actual Resolved Date & Time') or '-'}"
            )
            print(f"    desc={_get_col(row, 'Description')}")
            print(f"    notes={_get_col(row, 'Notes')}")
        print(
            f"candidate_counts=live:{len(analysis['live_requests'])} "
            f"quoted:{len(analysis['quoted_requests'])} "
            f"ack:{len(analysis['ack_candidates'])} "
            f"reply:{len(analysis['reply_candidates'])} "
            f"episodes:{len(analysis['unique_episodes'])}"
        )
        if analysis["unique_episodes"]:
            print("episodes:")
            for idx, episode in enumerate(analysis["unique_episodes"][:5], start=1):
                req, ack, resolved, req_kind = episode
                print(f"  {idx}. req={_fmt(req)} ack={_fmt(ack)} resolved={_fmt(resolved)} req_kind={req_kind}")
        printed += 1


if __name__ == "__main__":
    main()
