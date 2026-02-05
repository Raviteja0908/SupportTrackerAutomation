import os
import sys
import re
from pathlib import Path

from src.output.run_logger import RunLogger
from src.pst_reader import read_pst_emails
from src.rules.subject_normalizer import normalize_subject, extract_subject_from_description, normalize_subject_for_match
from src.rules.environment import resolve_environment
from src.rules.interface import resolve_interface_code
from src.rules.service_request import resolve_service_request
from src.rules.incident_type import resolve_incident_type
from src.rules.time_resolver import resolve_times_with_debug
from src.excel.template_filler import fill_template
from src.output.csv_writer import write_csv
from src.output.run_logger import MarkingReason
from src.utils import (
    load_json_list,
    load_aspose_license,
    load_subject_exclusions,
)


def main() -> int:
    input_dir = Path(os.getenv("INPUT_DIR", "/app/input"))
    output_dir = Path(os.getenv("OUTPUT_DIR", "/app/output"))
    config_dir = Path(os.getenv("CONFIG_DIR", "/app/config"))

    output_dir.mkdir(parents=True, exist_ok=True)
    logger = RunLogger(output_dir / "processing.log")

    load_aspose_license(logger)

    pst_files = list(input_dir.glob("*.pst"))
    if not pst_files:
        logger.log("[ERROR] No PST files found.")
        return 1

    ess_team = load_json_list(config_dir / "ess_team.json")
    subject_exclusions = load_subject_exclusions(config_dir / "subject_exclusions.json")

    logger.log(f"[INFO] Found {len(pst_files)} PST file(s).")

    eml_root = output_dir / "eml"
    eml_root.mkdir(parents=True, exist_ok=True)

    emails = []
    for pst_path in pst_files:
        logger.log(f"[INFO] Reading PST: {pst_path}")
        emails.extend(read_pst_emails(pst_path, logger, eml_root))

    logger.log(f"[INFO] Total emails extracted: {len(emails)}")

    threads = {}
    alt_index = {}
    for email in emails:
        subject_norm = normalize_subject(email.subject)
        if not subject_norm:
            continue
        if subject_exclusions and any(x in subject_norm.lower() for x in subject_exclusions):
            continue
        threads.setdefault(subject_norm, []).append(email)
    for key in threads.keys():
        alt_key = normalize_subject_for_match(key)
        if alt_key and alt_key != key:
            alt_index.setdefault(alt_key, []).append(key)

    logger.log(f"[INFO] Total threads: {len(threads)}")
    write_csv(
        output_dir / "thread_keys.csv",
        [{"SubjectKey": k} for k in sorted(threads.keys())],
        ["SubjectKey"],
    )

    # Load Excel template from output folder
    template_path = output_dir / "Support_Tracker_DEC_25_Incident_Business_done.xlsx"
    if not template_path.exists():
        logger.log(f"[ERROR] Template not found: {template_path}")
        return 1

    # Prepare CSV outputs for auditing
    automation_rows = []
    debug_rows = []
    same_time_rows = []

    def _match_tokens(text: str):
        if not text:
            return set()
        t = re.sub(r"[^a-z0-9]+", " ", text.lower())
        return {p for p in t.split() if p}

    def _looks_like_date_only(text: str) -> bool:
        if not text:
            return True
        s = text.strip()
        date_patterns = [
            r"^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$",
            r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$",
            r"^\d{1,2}[-/.]\d{1,2}$",
        ]
        return any(re.match(p, s) for p in date_patterns)

    def _interface_prefix(text: str) -> str:
        if not text:
            return ""
        # Match leading interface-like tokens (e.g., CS001, ID082, VMI001)
        m = re.match(r"^([a-z]{2,}\d{2,})", text.strip(), flags=re.IGNORECASE)
        return (m.group(1).lower() if m else "")

    def _token_overlap_score(a: set, b: set) -> float:
        if not a or not b:
            return 0.0
        inter = len(a & b)
        return (2 * inter) / (len(a) + len(b))

    def find_thread(subject_norm, requester):
        if not subject_norm:
            return [], "Empty subject"
        if subject_norm in threads:
            return threads[subject_norm], "Exact"
        alt_subject = normalize_subject_for_match(subject_norm)
        if alt_subject in threads:
            return threads[alt_subject], "AltExact"

        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return [], "No tokens"

        date_only = _looks_like_date_only(subject_norm)
        subj_prefix = _interface_prefix(subject_norm)

        candidates = []
        for key, thread in threads.items():
            key_tokens = _match_tokens(key)
            if not key_tokens:
                continue

            key_prefix = _interface_prefix(key)
            if subj_prefix and key_prefix and subj_prefix != key_prefix:
                continue

            score = _token_overlap_score(subj_tokens, key_tokens)
            contains = (subject_norm in key or key in subject_norm)

            if date_only:
                # Date-only subjects are high-risk; require stronger signal
                if score >= 0.8 or (subj_prefix and key_prefix == subj_prefix and score >= 0.5):
                    candidates.append((score, key, thread, "Score"))
                continue

            if score >= 0.6:
                candidates.append((score, key, thread, "Score"))
            elif contains and not date_only and len(subject_norm) >= 12 and score >= 0.4:
                candidates.append((score, key, thread, "Contains"))

        if alt_subject and alt_subject in alt_index:
            for key in alt_index[alt_subject]:
                thread = threads.get(key, [])
                candidates.append((0.95, key, thread, "AltKey"))

        if not candidates:
            return [], "No match"

        # Resolve by requester if ambiguous
        req = (requester or "").strip().lower()
        if req:
            for _, key, thread, _ in sorted(candidates, key=lambda x: (-x[0], -len(x[1]))):
                for e in thread:
                    if req in (e.sender_name or "").lower() or req in (e.sender_email or "").lower():
                        return thread, f"AmbiguousResolvedByRequester:{key}"

        # Pick best scored candidate
        candidates.sort(key=lambda x: (-x[0], -len(x[1])))
        top = candidates[0]
        if len(candidates) == 1:
            return top[2], f"{top[3]}:{top[1]}"

        # If multiple strong matches, avoid false positives
        if len(candidates) > 1 and top[0] < 0.75:
            return [], f"Ambiguous:{len(candidates)}"

        return top[2], f"{top[3]}:{top[1]}"

    def resolve_row(row_context):
        description = row_context.get("Description", "")
        requester = row_context.get("Consultant", "") or row_context.get("Requester", "")
        category_type = row_context.get("Category Type", "")

        subject_text = extract_subject_from_description(description)
        subject_norm = normalize_subject(subject_text)
        thread, match_note = find_thread(subject_norm, requester)

        env = resolve_environment(description)
        interface_code = resolve_interface_code(description)
        service_request = resolve_service_request(category_type)
        incident_type = resolve_incident_type(category_type, description)

        times, debug = resolve_times_with_debug(
            thread=thread,
            requester_name=requester,
            ess_team=ess_team,
        )

        if os.getenv("DEBUG_TIMES", "0") == "1":
            logger.log(
                "[DEBUG] "
                f"Requester={requester} | "
                f"SubjectKey={subject_norm} | "
                f"Match={match_note} | "
                f"Created={times.created} ({debug.created_src}) | "
                f"Response={times.response} ({debug.ack_src}) | "
                f"Resolved={times.resolved} ({debug.resolved_src}) | "
                f"Notes={debug.notes}"
            )

        automation_rows.append(
            {
                "Description": description,
                "Requester": requester,
                "Environment": env,
                "Created Date & Time": times.created,
                "Actual Response Date & Time": times.response,
                "Actual Resolved Date & Time": times.resolved,
                "ServiceRequest/Incident?": service_request,
                "ServiceRequest/Incident type?": incident_type,
                "Interface Code": interface_code,
            }
        )

        debug_rows.append(
            {
                "Description": description,
                "Requester": requester,
                "SubjectKey": subject_norm,
                "MatchFound": "Y" if thread else "N",
                "ThreadSize": len(thread),
                "CreatedSource": debug.created_src,
                "AckSource": debug.ack_src,
                "ResolvedSource": debug.resolved_src,
                "Notes": f"{debug.notes}; Match={match_note}",
            }
        )

        if times.created and times.response and times.resolved:
            if (
                times.created == times.response
                and times.response == times.resolved
            ):
                same_time_rows.append(
                    {
                        "Description": description,
                        "Requester": requester,
                        "SubjectKey": subject_norm,
                        "MatchFound": "Y" if thread else "N",
                        "ThreadSize": len(thread),
                        "Created": times.created,
                        "Response": times.response,
                        "Resolved": times.resolved,
                        "CreatedSource": debug.created_src,
                        "AckSource": debug.ack_src,
                        "ResolvedSource": debug.resolved_src,
                        "Notes": f"{debug.notes}; Match={match_note}",
                    }
                )

        resolved_values = {
            "Environment": env,
            "Created Date & Time": times.created,
            "Actual Response Date & Time": times.response,
            "Actual Resolved Date & Time": times.resolved,
            "ServiceRequest/Incident?": service_request,
            "ServiceRequest/Incident type?": incident_type,
            "Interface Code": interface_code,
        }
        if "Ack delayed" in debug.notes:
            resolved_values["_MarkBlue"] = True
        return resolved_values

    fill_result = fill_template(
        template_path=template_path,
        output_path=output_dir / "Support_Tracker_DEC_25_Incident_Business_done_filled.xlsx",
        row_resolver=resolve_row,
        logger=logger,
    )

    write_csv(
        output_dir / "automation_output.csv",
        automation_rows,
        [
            "Description",
            "Requester",
            "Environment",
            "Created Date & Time",
            "Actual Response Date & Time",
            "Actual Resolved Date & Time",
            "ServiceRequest/Incident?",
            "ServiceRequest/Incident type?",
            "Interface Code",
        ],
    )

    write_csv(
        output_dir / "debug_subjects.csv",
        debug_rows,
        [
            "Description",
            "Requester",
            "SubjectKey",
            "MatchFound",
            "ThreadSize",
            "CreatedSource",
            "AckSource",
            "ResolvedSource",
            "Notes",
        ],
    )

    write_csv(
        output_dir / "debug_same_times.csv",
        same_time_rows,
        [
            "Description",
            "Requester",
            "SubjectKey",
            "MatchFound",
            "ThreadSize",
            "Created",
            "Response",
            "Resolved",
            "CreatedSource",
            "AckSource",
            "ResolvedSource",
            "Notes",
        ],
    )

    subject_keys = [r.get("SubjectKey", "") for r in debug_rows if r.get("SubjectKey")]
    unique_subjects = set(subject_keys)
    matched_subjects = set(
        r.get("SubjectKey", "") for r in debug_rows if r.get("MatchFound") == "Y"
    )
    logger.log(
        f"[INFO] Subject match summary: {len(matched_subjects)}/{len(unique_subjects)} unique subjects matched."
    )

    logger.log(
        f"[INFO] Completed. Filled rows: {fill_result.filled_count}, "
        f"Skipped (maintenance): {fill_result.maintenance_count}, "
        f"Marked unknown: {fill_result.unknown_count}"
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
