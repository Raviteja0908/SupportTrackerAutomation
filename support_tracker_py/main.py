import os
import sys
import re
import html
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl.styles import PatternFill

from src.output.run_logger import RunLogger
from src.pst_reader import read_pst_emails
from src.rules.subject_normalizer import normalize_subject, extract_subject_from_description, normalize_subject_for_match
from src.rules.environment import resolve_environment
from src.rules.interface import resolve_interface_code
from src.rules.service_request import resolve_service_request
from src.rules.incident_type import resolve_incident_type
from src.rules.time_resolver import (
    resolve_times_with_debug,
    _match_requester,
    _is_ess_sender,
    _is_ack_like_reply,
    _extract_request_time_from_email,
    _format_time,
    _to_ist,
    TimeResult,
    TimeDebug,
)
from src.excel.template_filler import fill_template, EXPECTED_HEADERS
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
    inc_index = {}
    for email in emails:
        subject_norm = normalize_subject(email.subject)
        if not subject_norm:
            continue
        if subject_exclusions and any(x in subject_norm.lower() for x in subject_exclusions):
            # Keep maintenance subjects in threads so they can be filled.
            if "maintenance" not in subject_norm.lower():
                continue
        threads.setdefault(subject_norm, []).append(email)
    for key in threads.keys():
        alt_key = normalize_subject_for_match(key)
        if alt_key and alt_key != key:
            alt_index.setdefault(alt_key, []).append(key)

    # Build INC index from subject + body (including HTML) for safe fallback
    for key, thread in threads.items():
        incs = set()
        for e in thread:
            text = f"{e.subject or ''}\n{e.body or ''}\n{getattr(e, 'body_html', '') or ''}"
            for m in re.findall(r"\bINC\d{6,}\b", text, flags=re.IGNORECASE):
                incs.add(m.upper())
        for inc in incs:
            inc_index.setdefault(inc, []).append(key)

    logger.log(f"[INFO] Total threads: {len(threads)}")
    write_csv(
        output_dir / "thread_keys.csv",
        [{"SubjectKey": k} for k in sorted(threads.keys())],
        ["SubjectKey"],
    )

    # Load Excel template from output folder (default name or a single .xlsx)
    template_path = output_dir / "Support_Tracker_DEC_25_Incident_Business_done.xlsx"
    if not template_path.exists():
        # Fallback: pick a single .xlsx that is not an output "_filled" file
        candidates = [
            p for p in output_dir.glob("*.xlsx")
            if not p.name.lower().endswith("_filled.xlsx")
            and "_filled_" not in p.name.lower()
        ]
        if len(candidates) == 1:
            template_path = candidates[0]
            logger.log(f"[INFO] Using template: {template_path.name}")
        else:
            logger.log(f"[ERROR] Template not found: {template_path}")
            if candidates:
                logger.log("[ERROR] Multiple .xlsx candidates found. Please keep only one template:")
                for c in candidates:
                    logger.log(f" - {c.name}")
            return 1

    def _normalize_header(text: str) -> str:
        if text is None:
            return ""
        return " ".join(str(text).replace("\n", " ").split()).strip().lower()

    def _find_header_row(ws):
        expected = {_normalize_header(h) for h in EXPECTED_HEADERS}
        best_row = 1
        best_hits = 0
        for row in range(1, min(15, ws.max_row) + 1):
            row_values = [
                _normalize_header(ws.cell(row, c).value)
                for c in range(1, ws.max_column + 1)
            ]
            hits = sum(1 for v in row_values if v in expected and v)
            if hits > best_hits:
                best_hits = hits
                best_row = row
        return best_row

    def _build_col_map(ws, header_row):
        col_map = {}
        for col in range(1, ws.max_column + 1):
            raw = ws.cell(header_row, col).value
            if raw is None:
                continue
            key = _normalize_header(raw)
            if key and key not in col_map:
                col_map[key] = col
        return col_map

    # Prepare CSV outputs for auditing
    automation_rows = []
    debug_rows = []
    same_time_rows = []
    row_states = []

    def _match_tokens(text: str):
        if not text:
            return set()
        t = re.sub(r"[^a-z0-9]+", " ", text.lower())
        return {p for p in t.split() if p}

    def _part_tokens(text: str):
        if not text:
            return set()
        t = text.lower()
        tokens = set()
        for m in re.findall(r"\bpt\s*\d+\b", t):
            tokens.add(m.replace(" ", ""))
        for m in re.findall(r"\bpart\s*\d+\b", t):
            tokens.add(m.replace(" ", ""))
        for m in re.findall(r"\bpt\d+\b", t):
            tokens.add(m)
        return tokens

    def _extract_dr_ids(text: str):
        if not text:
            return set()
        ids = set()
        for m in re.findall(r"\bDR\s*ID\s*[:#]?\s*(DR-\d{6,}-\d+)\b", text, flags=re.IGNORECASE):
            ids.add(m.upper())
        for m in re.findall(r"\b(DR-\d{6,}-\d+)\b", text, flags=re.IGNORECASE):
            ids.add(m.upper())
        return ids

    def _thread_dr_ids(thread):
        ids = set()
        for e in thread:
            ids |= _extract_dr_ids(e.subject or "")
            ids |= _extract_dr_ids(e.body or "")
            ids |= _extract_dr_ids(getattr(e, "body_html", "") or "")
        return ids

    def _is_deployment_request_subject(text: str) -> bool:
        if not text:
            return False
        s = text.lower()
        return "deployment request" in s

    def _is_deployment_success_subject(text: str) -> bool:
        if not text:
            return False
        s = text.lower()
        return "success:" in s and "deployment" in s

    def _iface_tokens(text: str) -> set:
        if not text:
            return set()
        tokens = re.findall(r"\b[a-z]{1,5}\d{2,}\b", text, flags=re.IGNORECASE)
        return {t.lower() for t in tokens}

    ARROW_SPLIT_RE = r"\s*(?:--?>|Ã¢â€ â€™|Ã¢Å¾â€|Ã¢Å¾Â¡|=>)\s*"

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

    def _extract_date_tokens(text: str) -> list[str]:
        if not text:
            return []
        tokens = []
        tokens += re.findall(r"\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b", text)
        tokens += re.findall(r"\b\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\b", text)
        # de-dup, keep order
        seen = set()
        out = []
        for t in tokens:
            if t not in seen:
                seen.add(t)
                out.append(t)
        return out

    def _extract_date_tokens_from_description(description: str) -> list[str]:
        if not description:
            return []
        if not re.search(ARROW_SPLIT_RE, description):
            return []
        parts = re.split(ARROW_SPLIT_RE, description)
        if not parts:
            return []
        tail = parts[-1].strip()
        if not tail:
            return []
        return _extract_date_tokens(tail)

    def _has_explicit_date_marker(text: str) -> bool:
        if not text:
            return False
        if not re.search(ARROW_SPLIT_RE, text):
            return False
        parts = re.split(ARROW_SPLIT_RE, text)
        if not parts:
            return False
        tail = parts[-1].strip()
        if not tail:
            return False
        return bool(_extract_date_tokens(tail))

    def _thread_has_date_token(thread, date_tokens: list[str]) -> bool:
        if not date_tokens or not thread:
            return False
        parsed_dates = []
        for t in date_tokens:
            dt = _parse_date_token(t)
            if dt:
                parsed_dates.append(dt)
        for e in thread:
            if parsed_dates:
                sent_d = None
                try:
                    sent_d = _to_ist(e.sent_time).date()
                except Exception:
                    try:
                        sent_d = e.sent_time.date()
                    except Exception:
                        sent_d = None
                if sent_d and sent_d in parsed_dates:
                    return True
            text = f"{e.subject or ''}\n{e.body or ''}\n{getattr(e, 'body_html', '') or ''}"
            for t in date_tokens:
                if t in text:
                    return True
        return False

    def _interface_prefix(text: str) -> str:
        if not text:
            return ""
        # Match leading interface-like tokens (e.g., CS001, ID082, VMI001)
        m = re.match(r"^([a-z]{2,}\d{2,})", text.strip(), flags=re.IGNORECASE)
        return (m.group(1).lower() if m else "")

    def _interface_tokens(text: str) -> set:
        if not text:
            return set()
        # Find all interface-like tokens anywhere in the subject
        tokens = re.findall(r"\b[a-z]{1,5}\d{2,}\b", text, flags=re.IGNORECASE)
        return {t.lower() for t in tokens}

    def _inc_tokens(text: str) -> set:
        if not text:
            return set()
        # Incident tokens like INC2385330
        tokens = re.findall(r"\binc\d{6,}\b", text, flags=re.IGNORECASE)
        return {t.lower() for t in tokens}

    def _sig_num_tokens(text: str) -> set:
        """Return significant numeric tokens for disambiguation (exclude years)."""
        if not text:
            return set()
        out = set()
        for n in re.findall(r"\b\d{3,5}\b", text):
            try:
                iv = int(n)
            except Exception:
                continue
            if 1900 <= iv <= 2099:
                continue
            out.add(n)
        return out

    def _subject_for_description(description: str) -> str:
        subject_text = extract_subject_from_description(description or "")
        if description and re.search(r"(?:--?>|â†’|âž”|âž¡|=>)", description):
            parts = re.split(r"\s*(?:--?>|â†’|âž”|âž¡|=>)\s*", description, maxsplit=1)
            if len(parts) >= 2:
                _right = parts[1].strip()
                _right_l = _right.lower()
                if "deployment request" in _right_l or ("success:" in _right_l and "deployment" in _right_l):
                    subject_text = _right
        return subject_text

    def _requester_key(value: str) -> str:
        if not value:
            return ""
        s = str(value).strip().lower()
        return re.sub(r"[^a-z0-9]", "", s)

    def _parse_time_str(value: str):
        if not value:
            return None
        try:
            return datetime.strptime(value, "%d-%m-%Y %H:%M")
        except Exception:
            return None

    def _ack_missing(times: TimeResult | None, debug: TimeDebug | None) -> bool:
        if not times or not debug:
            return True
        notes = (debug.notes or "")
        if "ACK NOT FOUND" in notes:
            return True
        if times.created and times.response and times.created == times.response:
            return True
        if times.created and times.response and times.resolved:
            return times.created == times.response == times.resolved
        return False

    def _thread_has_consultant_on_date(thread, date_tokens: list[str], requester: str) -> bool:
        if not date_tokens or not thread or not requester:
            return False
        for e in thread:
            if _match_requester(e.sender_name, e.sender_email, requester):
                if _email_date_matches(e.sent_time, date_tokens):
                    return True
        return False

    def _thread_has_consultant_on_or_near_date(
        thread,
        date_tokens: list[str],
        requester: str,
        grace_hours: int = 12,
    ) -> bool:
        if not date_tokens or not thread or not requester:
            return False
        anchor_date = _anchor_date(date_tokens)
        if not anchor_date:
            return False
        anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
        window_start = anchor_start - timedelta(hours=grace_hours)
        for e in thread:
            if not _match_requester(e.sender_name, e.sender_email, requester):
                continue
            try:
                sent = _to_ist(e.sent_time)
            except Exception:
                continue
            if sent.date() == anchor_date or (window_start <= sent < anchor_start):
                return True
        return False

    def _thread_has_requester(thread, requester: str) -> bool:
        if not thread or not requester:
            return False
        for e in thread:
            if _match_requester(e.sender_name, e.sender_email, requester):
                return True
        return False

    def _pick_reply_after_ack(consultant_after, ack_ist, grace_minutes: int = 16):
        if not consultant_after:
            return None
        # Prefer the earliest non-ack/non-reminder reply after ack.
        non_ack = [e for e in consultant_after if not _is_ack_like_reply(e)]
        if non_ack:
            non_ack.sort(key=lambda e: e.sent_time)
            for e in non_ack:
                try:
                    if _to_ist(e.sent_time) > ack_ist:
                        return e
                except Exception:
                    continue
        if len(consultant_after) >= 2:
            first_dt = _to_ist(consultant_after[0].sent_time)
            if first_dt <= ack_ist + timedelta(minutes=grace_minutes) and _is_ack_like_reply(consultant_after[0]):
                # Skip to the next non-ack reply if available.
                for e in consultant_after[1:]:
                    if not _is_ack_like_reply(e):
                        return e
                return consultant_after[1]
        return consultant_after[0]

    def _find_unique_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens=None):
        if not subject_norm or not requester or not date_tokens:
            return None
        iface_tokens = iface_tokens or set()
        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return None
        subj_inc_set = _inc_tokens(subject_norm)
        subj_num_set = _sig_num_tokens(subject_norm)
        hits = []
        for key, thread in threads.items():
            if not _thread_has_requester(thread, requester):
                continue
            key_tokens = _match_tokens(key)
            if not key_tokens:
                continue
            if subj_inc_set:
                key_inc_set = _inc_tokens(key)
                if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                    continue
            key_num_set = _sig_num_tokens(key)
            if subj_num_set and key_num_set and subj_num_set.isdisjoint(key_num_set):
                continue
            if iface_tokens:
                key_iface_set = _interface_tokens(key)
                if not key_iface_set or key_iface_set.isdisjoint(iface_tokens):
                    continue
            score = _token_overlap_score(subj_tokens, key_tokens)
            contains = (subject_norm in key or key in subject_norm)
            if score < 0.45 and not contains:
                continue
            if not (
                _thread_has_date_token(thread, date_tokens)
                or _thread_has_consultant_on_or_near_date(thread, date_tokens, requester)
            ):
                continue
            hits.append((score, key, thread))
        if len(hits) == 1:
            return hits[0][2], f"RequesterDateUnique:{hits[0][1]}"
        return None

    def _find_best_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens=None):
        if not subject_norm or not requester or not date_tokens:
            return None
        anchor_date = _anchor_date(date_tokens)
        if not anchor_date:
            return None
        iface_tokens = iface_tokens or set()
        anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return None
        subj_inc_set = _inc_tokens(subject_norm)
        subj_num_set = _sig_num_tokens(subject_norm)
        best = None
        for key, thread in threads.items():
            if not _thread_has_requester(thread, requester):
                continue
            key_tokens = _match_tokens(key)
            if not key_tokens:
                continue
            if subj_inc_set:
                key_inc_set = _inc_tokens(key)
                if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                    continue
            key_num_set = _sig_num_tokens(key)
            if subj_num_set and key_num_set and subj_num_set.isdisjoint(key_num_set):
                continue
            if iface_tokens:
                key_iface_set = _interface_tokens(key)
                if not key_iface_set or key_iface_set.isdisjoint(iface_tokens):
                    continue
            score = _token_overlap_score(subj_tokens, key_tokens)
            contains = (subject_norm in key or key in subject_norm)
            if score < 0.35 and not contains:
                continue

            # Require a date signal (token in thread or consultant on/near anchor date).
            date_signal = (
                _thread_has_date_token(thread, date_tokens)
                or _thread_has_consultant_on_or_near_date(thread, date_tokens, requester)
            )
            if not date_signal:
                if not contains and score < 0.55:
                    continue

            # Find closest consultant reply to anchor start.
            consultant_replies = [
                e for e in thread
                if e.sent_time and _match_requester(e.sender_name, e.sender_email, requester)
            ]
            if not consultant_replies:
                continue
            consultant_non_ack = [e for e in consultant_replies if not _is_ack_like_reply(e)]
            if not consultant_non_ack:
                continue
            reply_pool = consultant_non_ack
            min_delta = None
            for e in reply_pool:
                try:
                    delta = abs(_to_ist(e.sent_time) - anchor_start)
                except Exception:
                    continue
                if min_delta is None or delta < min_delta:
                    min_delta = delta
            if min_delta is None:
                continue
            # Too far from anchor date → skip.
            if min_delta > timedelta(days=5):
                continue

            # Prefer smaller delta, then higher score, then date token presence.
            token_bonus = 0.2 if _thread_has_date_token(thread, date_tokens) else 0.0
            score = min(1.0, score + (0.1 if contains else 0.0) + token_bonus)
            candidate = (min_delta, -score, key, thread)
            if best is None or candidate < best:
                best = candidate

        if best:
            _, _, key, thread = best
            return thread, f"RequesterDateBest:{key}"
        return None

    def _union_alt_threads(alt_subject, requester=None):
        if not alt_subject:
            return None
        keys = alt_index.get(alt_subject, [])
        if not keys:
            return None
        merged = []
        seen = set()
        has_requester = False
        for k in keys:
            for e in threads.get(k, []):
                dedup_key = (e.subject, e.sender_email, e.sent_time)
                if dedup_key in seen:
                    continue
                seen.add(dedup_key)
                merged.append(e)
                if requester and _match_requester(e.sender_name, e.sender_email, requester):
                    has_requester = True
        if requester and not has_requester:
            return None
        return merged

    def _parse_date_token(token: str):
        if not token:
            return None
        t = token.strip()
        try:
            # yyyy-mm-dd / yyyy/mm/dd / yyyy.mm.dd
            if re.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$", t):
                parts = re.split(r"[-/.]", t)
                y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                return datetime(y, m, d).date()
            # dd-mm-yyyy or dd-mm-yy (day-first)
            if re.match(r"^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$", t):
                parts = re.split(r"[-/.]", t)
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y = 2000 + y
                return datetime(y, m, d).date()
        except Exception:
            return None
        return None

    def _email_date_matches(sent_time, date_tokens: list[str]) -> bool:
        if not sent_time or not date_tokens:
            return False
        try:
            sent_d = _to_ist(sent_time).date()
        except Exception:
            try:
                sent_d = sent_time.date()
            except Exception:
                return False
        for t in date_tokens:
            dt = _parse_date_token(t)
            if dt and dt == sent_d:
                return True
        return False

    def _anchor_date(date_tokens: list[str]):
        if not date_tokens:
            return None
        for t in date_tokens:
            dt = _parse_date_token(t)
            if dt:
                return dt
        return None

    def _coerce_row_created_date(value):
        if value is None:
            return None
        if isinstance(value, datetime):
            try:
                return _to_ist(value).date()
            except Exception:
                return value.date()
        # openpyxl may return date objects for date-only cells.
        if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day") and not isinstance(value, str):
            try:
                return datetime(int(value.year), int(value.month), int(value.day)).date()
            except Exception:
                pass
        s = str(value).strip()
        if not s:
            return None
        # Common datetime string formats seen in exports/templates.
        for fmt in (
            "%d-%m-%Y %H:%M",
            "%d/%m/%Y %H:%M",
            "%d.%m.%Y %H:%M",
            "%d-%m-%Y %H:%M:%S",
            "%d/%m/%Y %H:%M:%S",
            "%d.%m.%Y %H:%M:%S",
            "%d-%m-%Y %I:%M %p",
            "%d/%m/%Y %I:%M %p",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y/%m/%d %H:%M:%S",
            "%Y/%m/%d %H:%M",
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%d.%m.%Y",
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%Y.%m.%d",
        ):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                pass
        # Try to pull a date token from mixed strings, e.g. "2026-01-06 10:22:11 UTC".
        m = re.search(r"\b\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\b", s)
        if m:
            d = _parse_date_token(m.group(0))
            if d:
                return d
        m = re.search(r"\b\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}\b", s)
        if m:
            d = _parse_date_token(m.group(0))
            if d:
                return d
        dt = _parse_time_str(s)
        if dt:
            try:
                return _to_ist(dt).date()
            except Exception:
                return dt.date()
        try:
            d = _parse_date_token(s)
            return d
        except Exception:
            return None

    def _is_stale_anchor_date(date_tokens: list[str], row_created_value, max_days: int = 14) -> bool:
        anchor = _anchor_date(date_tokens)
        if not anchor:
            return False
        created_d = _coerce_row_created_date(row_created_value)
        if not created_d:
            return False
        return abs((created_d - anchor).days) > max_days

    def _anchor_relevant_to_thread(anchor_date, thread, max_days: int = 21) -> bool:
        if not anchor_date or not thread:
            return False
        thread_dates = []
        for e in thread:
            if not e.sent_time:
                continue
            try:
                thread_dates.append(_to_ist(e.sent_time).date())
            except Exception:
                continue
        if not thread_dates:
            return False
        min_d = min(thread_dates)
        max_d = max(thread_dates)
        return (min_d - timedelta(days=max_days)) <= anchor_date <= (max_d + timedelta(days=max_days))

    def _baseline_row_created_value(row_context: dict):
        # Prefer immutable ServiceNow baseline if present.
        # "SLA Start Date & Time" is not overwritten by our resolver.
        sla_start = row_context.get("Sla Start Date & Time") or row_context.get("SLA Start Date & Time")
        if _coerce_row_created_date(sla_start):
            return sla_start
        return row_context.get("Created Date & Time")

    def _requester_min_delta_days(thread, requester: str, baseline_date):
        if not thread or not requester or not baseline_date:
            return None
        best = None
        for e in thread:
            if not e.sent_time:
                continue
            if not _match_requester(e.sender_name, e.sender_email, requester):
                continue
            try:
                sent_d = _to_ist(e.sent_time).date()
            except Exception:
                try:
                    sent_d = e.sent_time.date()
                except Exception:
                    continue
            delta = abs((sent_d - baseline_date).days)
            if best is None or delta < best:
                best = delta
        return best

    def _find_best_thread_near_baseline(subject_norm: str, requester: str, baseline_date, iface_tokens=None):
        if not subject_norm or not requester or not baseline_date:
            return None
        iface_tokens = iface_tokens or set()
        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return None
        subj_inc_set = _inc_tokens(subject_norm)
        subj_iface_set = _interface_tokens(subject_norm)
        if not subj_iface_set and iface_tokens:
            subj_iface_set = iface_tokens

        best = None
        for key, thread in threads.items():
            if not _thread_has_requester(thread, requester):
                continue
            key_tokens = _match_tokens(key)
            if not key_tokens:
                continue
            if subj_inc_set:
                key_inc_set = _inc_tokens(key)
                if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                    continue
            if subj_iface_set:
                key_iface_set = _interface_tokens(key)
                if key_iface_set and subj_iface_set.isdisjoint(key_iface_set):
                    continue

            score = _token_overlap_score(subj_tokens, key_tokens)
            contains = (subject_norm in key or key in subject_norm)
            if score < 0.35 and not contains:
                continue

            delta_days = _requester_min_delta_days(thread, requester, baseline_date)
            if delta_days is None:
                continue
            # Keep this bounded so we don't jump to unrelated long-history threads.
            if delta_days > 14:
                continue

            candidate = (delta_days, -score, -len(key_tokens), key, thread)
            if best is None or candidate < best:
                best = candidate

        if best:
            delta_days, _neg_score, _neg_len, key, thread = best
            return thread, f"BaselineDateRefined:{key}", delta_days
        return None

    def _email_on_or_after_date(sent_time, anchor_date) -> bool:
        if not sent_time or not anchor_date:
            return False
        try:
            sent_d = _to_ist(sent_time).date()
        except Exception:
            try:
                sent_d = sent_time.date()
            except Exception:
                return False
        return sent_d >= anchor_date

    def _thread_has_non_ess_near_time(thread, target_dt, hours=24) -> bool:
        if not thread or not target_dt:
            return False
        try:
            target_ist = _to_ist(target_dt)
        except Exception:
            return False
        window = timedelta(hours=hours)
        for e in thread:
            if _is_ess_sender(e, ess_team):
                continue
            if not e.sent_time:
                continue
            try:
                sent_ist = _to_ist(e.sent_time)
            except Exception:
                continue
            if abs(sent_ist - target_ist) <= window:
                return True
        return False

    def _latest_non_ess_before(thread, cutoff_dt):
        if not thread or not cutoff_dt:
            return None
        out = None
        cutoff_ist = _to_ist(cutoff_dt)
        for e in thread:
            if _is_ess_sender(e, ess_team):
                continue
            if not e.sent_time:
                continue
            try:
                sent_ist = _to_ist(e.sent_time)
            except Exception:
                continue
            if sent_ist <= cutoff_ist and (out is None or sent_ist > _to_ist(out.sent_time)):
                out = e
        return out

    def _apply_final_time_guards(
        thread,
        requester,
        times,
        debug,
        base_times,
        base_debug,
        baseline_created_date=None,
        is_dep_req=False,
        is_dep_succ=False,
    ):
        if not times or not debug or not thread:
            return times, debug
        if is_dep_req or is_dep_succ:
            return times, debug

        notes = debug.notes or ""
        if "Maintenance override" in notes:
            return times, debug

        created_dt = _parse_time_str(times.created)
        resp_dt = _parse_time_str(times.response)
        res_dt = _parse_time_str(times.resolved)
        created_far_from_baseline = False
        if created_dt and baseline_created_date:
            try:
                created_far_from_baseline = abs((_to_ist(created_dt).date() - baseline_created_date).days) > 14
            except Exception:
                created_far_from_baseline = False

        # Guard 1: when Created comes from parsed quote and there is no nearby
        # non-ESS evidence, anchor Created back to latest non-ESS request.
        if (
            created_dt
            and isinstance(debug.created_src, str)
            and debug.created_src.startswith("PARSED_")
            and any(not _is_ess_sender(e, ess_team) for e in thread)
            and (
                not _thread_has_non_ess_near_time(thread, created_dt, hours=24)
                or created_far_from_baseline
            )
        ):
            cutoff = resp_dt or res_dt
            if not cutoff:
                consultant_times = [
                    e.sent_time for e in thread
                    if e.sent_time and _match_requester(e.sender_name, e.sender_email, requester)
                ]
                if consultant_times:
                    try:
                        cutoff = max(consultant_times, key=_to_ist)
                    except Exception:
                        cutoff = None
            req = _latest_non_ess_before(thread, cutoff) if cutoff else None
            if req:
                times = TimeResult(_format_time(req.sent_time), times.response, times.resolved)
                debug = TimeDebug(
                    req.sender_email or req.sender_name,
                    debug.ack_src,
                    debug.resolved_src,
                    f"{notes}; GuardCreatedFromNonESS{'Baseline' if created_far_from_baseline else ''}",
                )
                notes = debug.notes
                created_dt = _parse_time_str(times.created)

        # Guard 2: enforce Created <= Response using non-ESS request fallback.
        if created_dt and resp_dt and _to_ist(created_dt) > _to_ist(resp_dt):
            req = _latest_non_ess_before(thread, resp_dt)
            if req:
                times = TimeResult(_format_time(req.sent_time), times.response, times.resolved)
                debug = TimeDebug(
                    req.sender_email or req.sender_name,
                    debug.ack_src,
                    debug.resolved_src,
                    f"{notes}; GuardCreatedBeforeResponse",
                )
                notes = debug.notes
                created_dt = _parse_time_str(times.created)
            elif base_times:
                base_created_dt = _parse_time_str(base_times.created)
                if base_created_dt and _to_ist(base_created_dt) <= _to_ist(resp_dt):
                    times = TimeResult(base_times.created, times.response, times.resolved)
                    src = base_debug.created_src if base_debug else debug.created_src
                    debug = TimeDebug(
                        src,
                        debug.ack_src,
                        debug.resolved_src,
                        f"{notes}; GuardCreatedRevertBase",
                    )
                    notes = debug.notes
                    created_dt = _parse_time_str(times.created)

        # Guard 3: enforce Resolved >= Response when both exist.
        if resp_dt and res_dt and _to_ist(res_dt) < _to_ist(resp_dt):
            ack_ist = _to_ist(resp_dt)
            window_end = ack_ist + timedelta(hours=48)
            consultant_after = [
                e for e in thread
                if _match_requester(e.sender_name, e.sender_email, requester)
                and e.sent_time
                and _to_ist(e.sent_time) > ack_ist
                and _to_ist(e.sent_time) <= window_end
                and not _is_ack_like_reply(e)
            ]
            consultant_after.sort(key=lambda e: e.sent_time)
            if consultant_after:
                pick = consultant_after[0]
                times = TimeResult(times.created, times.response, _format_time(pick.sent_time))
                debug = TimeDebug(
                    debug.created_src,
                    debug.ack_src,
                    pick.sender_email or pick.sender_name,
                    f"{notes}; GuardResolvedAfterResponse",
                )
            elif base_times:
                base_res_dt = _parse_time_str(base_times.resolved)
                if base_res_dt and _to_ist(base_res_dt) >= _to_ist(resp_dt):
                    src = base_debug.resolved_src if base_debug else debug.resolved_src
                    times = TimeResult(times.created, times.response, base_times.resolved)
                    debug = TimeDebug(
                        debug.created_src,
                        debug.ack_src,
                        src,
                        f"{notes}; GuardResolvedRevertBase",
                    )

        # Guard 4 (ESS-only baseline alignment): when row baseline day is far from
        # current created day, prefer requester reply nearest to baseline day.
        if (
            baseline_created_date
            and thread
            and requester
            and "ESS-only; no non-ESS request" in notes
        ):
            created_dt = _parse_time_str(times.created)
            created_ist = _to_ist(created_dt) if created_dt else None
            current_delta = abs((created_ist.date() - baseline_created_date).days) if created_ist else 9999
            if current_delta > 2:
                requester_replies = [
                    e for e in thread
                    if e.sent_time and _match_requester(e.sender_name, e.sender_email, requester)
                ]
                if requester_replies:
                    requester_replies.sort(key=lambda e: e.sent_time)
                    requester_non_ack = [e for e in requester_replies if not _is_ack_like_reply(e)]
                    if not requester_non_ack:
                        return times, debug
                    reply_pool = requester_non_ack
                    # Do not baseline-force across multi-episode threads.
                    # If requester has a much later episode, keep current timing.
                    latest_reply = reply_pool[-1]

                    baseline_mid = _to_ist(datetime(
                        baseline_created_date.year,
                        baseline_created_date.month,
                        baseline_created_date.day,
                    ))
                    def _rank(e):
                        sent = _to_ist(e.sent_time)
                        day_delta = abs((sent.date() - baseline_created_date).days)
                        return (day_delta, abs(sent - baseline_mid))
                    best = min(reply_pool, key=_rank)
                    best_day_delta = abs((_to_ist(best.sent_time).date() - baseline_created_date).days)
                    allow_baseline_force = (_to_ist(latest_reply.sent_time) - _to_ist(best.sent_time)) <= timedelta(days=2)
                    if allow_baseline_force and best_day_delta <= 2 and best_day_delta < current_delta:
                        t = _format_time(best.sent_time)
                        if t:
                            times = TimeResult(t, t, t)
                            debug = TimeDebug(
                                best.sender_email or best.sender_name,
                                best.sender_email or best.sender_name,
                                best.sender_email or best.sender_name,
                                f"{notes}; GuardBaselineRequesterSpan",
                            )
        return times, debug

    def _token_overlap_score(a: set, b: set) -> float:
        if not a or not b:
            return 0.0
        inter = len(a & b)
        return (2 * inter) / (len(a) + len(b))

    # Pre-scan template to count repeated subjects per consultant (safe, read-only)
    group_counts = {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(template_path)
        ws = wb.active
        header_row = _find_header_row(ws)
        col_map = _build_col_map(ws, header_row)
        desc_col = col_map.get("description")
        consultant_col = col_map.get("consultant") or col_map.get("requester")
        if desc_col and consultant_col:
            for row in range(header_row + 1, ws.max_row + 1):
                desc_val = ws.cell(row, desc_col).value
                cons_val = ws.cell(row, consultant_col).value
                if not desc_val or not cons_val:
                    continue
                subject_text = _subject_for_description(str(desc_val))
                subject_norm = normalize_subject(subject_text)
                if not subject_norm:
                    continue
                if subject_exclusions and any(x in subject_norm.lower() for x in subject_exclusions):
                    if "maintenance" not in subject_norm.lower():
                        continue
                key = (subject_norm, _requester_key(cons_val))
                group_counts[key] = group_counts.get(key, 0) + 1
        wb.close()
    except Exception as e:
        logger.log(f"[WARNING] Pre-scan for repeated subjects failed: {e}")
        group_counts = {}

    episode_counters = {}
    duplicate_group_state = {}
    created_history = []
    env_cache = {}

    # Build deployment request/success index by DR ID (safe override case only)
    deployment_index = {}
    for key, thread in threads.items():
        subj_l = (key or "").lower()
        is_req = _is_deployment_request_subject(subj_l)
        is_succ = _is_deployment_success_subject(subj_l)
        if not (is_req or is_succ):
            continue
        dr_ids = _thread_dr_ids(thread)
        if not dr_ids:
            continue
        info = {
            "thread": thread,
            "subject_key": key,
            "iface": _iface_tokens(key),
        }
        for dr in dr_ids:
            bucket = deployment_index.setdefault(dr, {"request": [], "success": []})
            if is_req:
                bucket["request"].append(info)
            else:
                bucket["success"].append(info)

    def _find_iface_specific_thread(subject_norm, requester, date_tokens=None, iface_tokens=None, baseline_date=None):
        iface_tokens = iface_tokens or set()
        if not subject_norm or not iface_tokens:
            return None
        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return None
        short_stop = {
            "re", "fw", "fwd", "aw", "wg", "sv",
            "in", "on", "at", "to", "of", "for", "not",
            "and", "or", "the", "a", "an", "is", "it",
            "by", "as", "if", "no", "we", "us",
            "today", "below", "details", "file", "files",
            "input", "output", "received", "failed",
        }
        subj_short = {t for t in subj_tokens if len(t) <= 3 and t.isalpha() and t not in short_stop}

        ranked = []
        for key, thread in threads.items():
            key_iface_set = _interface_tokens(key)
            if not key_iface_set or key_iface_set.isdisjoint(iface_tokens):
                continue
            key_tokens = _match_tokens(key)
            if not key_tokens:
                continue
            score = _token_overlap_score(subj_tokens, key_tokens)
            contains = (subject_norm in key or key in subject_norm)
            if score < 0.35 and not contains:
                continue
            short_overlap = len(subj_short & key_tokens) if subj_short else 0

            has_consultant_date = False
            if date_tokens and requester:
                has_consultant_date = _thread_has_consultant_on_or_near_date(thread, date_tokens, requester)

            delta_days = 9999
            if baseline_date and requester:
                d = _requester_min_delta_days(thread, requester, baseline_date)
                if d is not None:
                    delta_days = d

            ranked.append((0 if has_consultant_date else 1, -short_overlap, delta_days, -score, key, thread))

        if not ranked:
            return None
        if subj_short:
            max_short = max((-r[1]) for r in ranked)
            if max_short > 0:
                ranked = [r for r in ranked if (-r[1]) == max_short]
        ranked.sort()
        best = ranked[0]
        return best[5], f"IfaceSpecific:{best[4]}"

    def find_thread(subject_norm, requester, date_tokens=None, prefer_consultant_date=False, iface_tokens=None, baseline_date=None):
        if not subject_norm:
            return [], "Empty subject"
        iface_tokens = iface_tokens or set()
        alt_subject = normalize_subject_for_match(subject_norm)
        enable_safe_iface_hint = os.getenv("SAFE_IFACE_HINT", "1") == "1"

        def _strong_exact_signal(t):
            if not t:
                return False
            if requester and _thread_has_requester(t, requester):
                return True
            if date_tokens and requester and _thread_has_consultant_on_or_near_date(t, date_tokens, requester):
                return True
            if baseline_date and requester:
                d = _requester_min_delta_days(t, requester, baseline_date)
                if d is not None and d <= 2:
                    return True
            return False

        def _prefer_iface_pick(exact_t, iface_t):
            if not exact_t or not iface_t:
                return False
            if requester:
                exact_has_req = _thread_has_requester(exact_t, requester)
                iface_has_req = _thread_has_requester(iface_t, requester)
                if iface_has_req and not exact_has_req:
                    return True
            if date_tokens and requester:
                exact_has_date = _thread_has_consultant_on_or_near_date(exact_t, date_tokens, requester)
                iface_has_date = _thread_has_consultant_on_or_near_date(iface_t, date_tokens, requester)
                if iface_has_date and not exact_has_date:
                    return True
            if baseline_date and requester:
                exact_delta = _requester_min_delta_days(exact_t, requester, baseline_date)
                iface_delta = _requester_min_delta_days(iface_t, requester, baseline_date)
                if iface_delta is not None and (exact_delta is None or (iface_delta + 1) < exact_delta):
                    return True
            return False

        exact_thread = None
        exact_note = None
        if subject_norm in threads:
            t = threads[subject_norm]
            key_iface_set = _interface_tokens(subject_norm)
            if iface_tokens and not key_iface_set and enable_safe_iface_hint:
                iface_pick = _find_iface_specific_thread(
                    subject_norm,
                    requester,
                    date_tokens=date_tokens,
                    iface_tokens=iface_tokens,
                    baseline_date=baseline_date,
                )
                if iface_pick:
                    iface_thread, _iface_note = iface_pick
                    if (not _strong_exact_signal(t)) or _prefer_iface_pick(t, iface_thread):
                        return iface_pick
            # Interface token is a hint, not a blocker for exact subject match.
            # Keep exact matches even when the subject has no interface token.
            if (not iface_tokens) or (not key_iface_set) or (not iface_tokens.isdisjoint(key_iface_set)):
                if prefer_consultant_date and date_tokens and requester and not _thread_has_requester(t, requester):
                    alt = _find_unique_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens)
                    if alt:
                        return alt
                    alt = _find_best_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens)
                    if alt:
                        return alt
                    # Last-resort union of alt-subject threads when requester is missing
                    if alt_subject:
                        union = _union_alt_threads(alt_subject, requester)
                        if union:
                            return union, f"AltUnion:{alt_subject}"
                return t, "Exact"
            exact_thread = t
            exact_note = "Exact"
        if alt_subject in threads:
            t = threads[alt_subject]
            key_iface_set = _interface_tokens(alt_subject)
            if iface_tokens and not key_iface_set and enable_safe_iface_hint:
                iface_pick = _find_iface_specific_thread(
                    subject_norm,
                    requester,
                    date_tokens=date_tokens,
                    iface_tokens=iface_tokens,
                    baseline_date=baseline_date,
                )
                if iface_pick:
                    iface_thread, _iface_note = iface_pick
                    if (not _strong_exact_signal(t)) or _prefer_iface_pick(t, iface_thread):
                        return iface_pick
            # Interface token is a hint, not a blocker for exact alt-subject match.
            if (not iface_tokens) or (not key_iface_set) or (not iface_tokens.isdisjoint(key_iface_set)):
                if prefer_consultant_date and date_tokens and requester and not _thread_has_requester(t, requester):
                    alt = _find_unique_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens)
                    if alt:
                        return alt
                    alt = _find_best_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens)
                    if alt:
                        return alt
                    # Last-resort union of alt-subject threads when requester is missing
                    if alt_subject:
                        union = _union_alt_threads(alt_subject, requester)
                        if union:
                            return union, f"AltUnion:{alt_subject}"
                return t, "AltExact"
            if exact_thread is None:
                exact_thread = t
                exact_note = "AltExact"

        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return [], "No tokens"

        date_only = _looks_like_date_only(subject_norm)
        subj_prefix = _interface_prefix(subject_norm)
        subj_iface_set = _interface_tokens(subject_norm)
        if not subj_iface_set and iface_tokens:
            subj_iface_set = iface_tokens
        subj_inc_set = _inc_tokens(subject_norm)
        subj_num_set = _sig_num_tokens(subject_norm)

        subj_part_set = _part_tokens(subject_norm)
        short_stop = {
            "re", "fw", "fwd", "aw", "wg", "sv",
            "in", "on", "at", "to", "of", "for", "not",
            "and", "or", "the", "a", "an", "is", "it",
            "by", "as", "if", "no", "we", "us",
        }
        short_tokens = {t for t in subj_tokens if len(t) <= 3 and t.isalpha() and t not in short_stop}

        candidates = []
        for key, thread in threads.items():
            key_tokens = _match_tokens(key)
            if not key_tokens:
                continue

            key_prefix = _interface_prefix(key)
            if subj_prefix and key_prefix and subj_prefix != key_prefix:
                continue

            key_iface_set = _interface_tokens(key)
            if subj_iface_set and key_iface_set and subj_iface_set.isdisjoint(key_iface_set):
                continue

            key_inc_set = _inc_tokens(key)
            if subj_inc_set:
                # Require INC match when subject has INC
                if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                    continue
            key_num_set = _sig_num_tokens(key)
            if subj_num_set and key_num_set and subj_num_set.isdisjoint(key_num_set):
                continue

            score = _token_overlap_score(subj_tokens, key_tokens)
            contains = (subject_norm in key or key in subject_norm)
            boost = 0.0
            if subj_inc_set and key_inc_set and not subj_inc_set.isdisjoint(key_inc_set):
                boost += 0.2
            if subj_iface_set and key_iface_set and not subj_iface_set.isdisjoint(key_iface_set):
                boost += 0.2
            if contains:
                boost += 0.1
            if short_tokens:
                short_hit = len(short_tokens & key_tokens)
                if short_hit:
                    boost += min(0.15, 0.05 * short_hit)
            score = min(1.0, score + boost)

            if date_only:
                # Date-only subjects are high-risk; require stronger signal
                if score >= 0.8 or (subj_prefix and key_prefix == subj_prefix and score >= 0.5):
                    candidates.append((score, key, thread, "Score"))
                continue

            if score >= 0.55:
                candidates.append((score, key, thread, "Score"))
            elif short_tokens and (short_tokens & key_tokens) and score >= 0.35:
                # Keep near-miss siblings that match distinguishing short tokens
                # (e.g., INT/GB) so VariantTok can disambiguate correctly.
                candidates.append((score, key, thread, "ScoreShort"))
            elif contains and not date_only and len(subject_norm) >= 12 and score >= 0.35:
                candidates.append((score, key, thread, "Contains"))

        if alt_subject and alt_subject in alt_index:
            for key in alt_index[alt_subject]:
                thread = threads.get(key, [])
                candidates.append((0.95, key, thread, "AltKey"))

        # If the subject has INC IDs, prefer exact INC set matches.
        # This prevents matching a superset INC subject when an exact one exists.
        if subj_inc_set and candidates:
            exact_inc = []
            for score, key, thread, note in candidates:
                key_inc_set = _inc_tokens(key)
                if key_inc_set == subj_inc_set:
                    exact_inc.append((score, key, thread, note + "+IncExact"))
            if exact_inc:
                candidates = exact_inc

        # If we have an interface prefix hint (from description), prefer
        # threads that include those interface tokens.
        if iface_tokens and candidates:
            iface_candidates = []
            for score, key, thread, note in candidates:
                key_iface_set = _interface_tokens(key)
                if key_iface_set and not iface_tokens.isdisjoint(key_iface_set):
                    iface_candidates.append((score, key, thread, note + "+Iface"))
            if iface_candidates:
                candidates = iface_candidates

        if subj_part_set and candidates:
            part_candidates = []
            for score, key, thread, note in candidates:
                if _part_tokens(key) & subj_part_set:
                    part_candidates.append((score, key, thread, note + '+Part'))
            if part_candidates:
                candidates = part_candidates

        # Variant-token disambiguation for sibling subjects (e.g., INT vs GB).
        # Keep this non-destructive: boost scores and only hard-pick when one
        # candidate is uniquely best on variant-token overlap.
        if len(candidates) > 1:
            variant_stop = {
                "re", "fw", "fwd", "aw", "wg", "sv",
                "in", "on", "at", "to", "of", "for", "not",
                "and", "or", "the", "a", "an", "is", "it",
                "by", "as", "if", "no", "we", "us",
                "today", "below", "details", "file", "files",
                "input", "output", "received", "failed",
            }
            cand_with_tokens = []
            token_counts = {}
            for score, key, thread, note in candidates:
                key_tokens = _match_tokens(key)
                cand_with_tokens.append((score, key, thread, note, key_tokens))
                for t in key_tokens:
                    token_counts[t] = token_counts.get(t, 0) + 1

            varying_tokens = {
                t for t in subj_tokens
                if t not in variant_stop and 0 < token_counts.get(t, 0) < len(candidates)
            }
            if varying_tokens:
                best_overlap = 0
                winners = []
                boosted = []
                for score, key, thread, note, key_tokens in cand_with_tokens:
                    overlap = len(varying_tokens & key_tokens)
                    if overlap > best_overlap:
                        best_overlap = overlap
                        winners = [(score, key, thread, note)]
                    elif overlap == best_overlap:
                        winners.append((score, key, thread, note))

                    bonus = min(0.18, 0.06 * overlap) if overlap > 0 else 0.0
                    boosted_score = min(1.0, score + bonus)
                    boosted_note = note + "+VariantTok" if overlap > 0 else note
                    boosted.append((boosted_score, key, thread, boosted_note))

                # Only force a pick when variant-token evidence has a unique winner.
                if best_overlap > 0 and len(winners) == 1:
                    _, key, thread, note = winners[0]
                    return thread, f"{note}+VariantTokUnique:{key}"

                # Otherwise keep all candidates and only adjust scores/notes.
                candidates = boosted

        # (Short-token disambiguation is handled via score boosting above)

        # If subject carries distinguishing short tokens but none of the current
        # candidates contain them, run a global variant scan to avoid sibling drift
        # (e.g., selecting GB thread for an INT subject).
        if short_tokens and candidates:
            cand_has_variant = False
            for _, key, _, _ in candidates:
                if short_tokens & _match_tokens(key):
                    cand_has_variant = True
                    break
            if not cand_has_variant:
                variant_global = []
                for key, thread in threads.items():
                    key_tokens = _match_tokens(key)
                    if not (short_tokens & key_tokens):
                        continue

                    key_prefix = _interface_prefix(key)
                    if subj_prefix and key_prefix and subj_prefix != key_prefix:
                        continue

                    key_iface_set = _interface_tokens(key)
                    if subj_iface_set and key_iface_set and subj_iface_set.isdisjoint(key_iface_set):
                        continue

                    key_inc_set = _inc_tokens(key)
                    if subj_inc_set and (not key_inc_set or subj_inc_set.isdisjoint(key_inc_set)):
                        continue

                    score = _token_overlap_score(subj_tokens, key_tokens)
                    contains = (subject_norm in key or key in subject_norm)
                    if score < 0.35 and not contains:
                        continue
                    variant_overlap = len(short_tokens & key_tokens)
                    variant_global.append((variant_overlap, score, key, thread))

                if variant_global:
                    variant_global.sort(key=lambda x: (-x[0], -x[1], -len(x[2])))
                    best = variant_global[0]
                    second = variant_global[1] if len(variant_global) > 1 else None
                    if second is None or best[0] > second[0] or (best[0] == second[0] and (best[1] - second[1]) >= 0.08):
                        return best[3], f"VariantGlobal:{best[2]}"

        if not candidates and subj_inc_set:
            inc_upper = {i.upper() for i in subj_inc_set}
            inc_threads = []
            for inc in inc_upper:
                for key in inc_index.get(inc, []):
                    thread = threads.get(key, [])
                    if thread:
                        inc_threads.append((key, thread))

            # Only accept if there is a single unique thread (safe fallback)
            unique = {(key, id(thread)): thread for key, thread in inc_threads}
            if len(unique) == 1:
                only_thread = next(iter(unique.values()))
                return only_thread, "INCBodyOnly"

            if len(unique) > 1 and requester:
                req = (requester or "").strip().lower()
                for key, thread in inc_threads:
                    for e in thread:
                        if req in (e.sender_name or "").lower() or req in (e.sender_email or "").lower():
                            return thread, f"INCBodyOnly+Requester:{key}"

        if not candidates:
            if exact_thread is not None:
                return exact_thread, exact_note
            return [], "No match"

        # Prefer candidates that actually contain a non-ESS request email.
        # This avoids matching ESS-only threads when a real request exists elsewhere.
        def _thread_has_non_ess(t):
            for e in t:
                if not _is_ess_sender(e, ess_team):
                    return True
            return False

        non_ess_candidates = [c for c in candidates if _thread_has_non_ess(c[2])]
        if non_ess_candidates:
            # When the row has an explicit date marker and consultant, keep
            # ESS-only candidates so we don't drop a valid consultant reply
            # thread for date-anchored rows.
            if not (prefer_consultant_date and date_tokens and requester):
                candidates = non_ess_candidates

        # If we still don't have a consultant-on-date candidate, do a global
        # scan for explicit date markers and pick the best thread that has a
        # requester reply on/near that date.
        if prefer_consultant_date and date_tokens and requester:
            has_consultant_date = False
            for _, _, thread, _ in candidates:
                if _thread_has_consultant_on_or_near_date(thread, date_tokens, requester):
                    has_consultant_date = True
                    break
            if not has_consultant_date:
                subj_tokens = _match_tokens(subject_norm)
                subj_inc_set = _inc_tokens(subject_norm)
                subj_num_set = _sig_num_tokens(subject_norm)
                best = None
                for key, thread in threads.items():
                    if not _thread_has_consultant_on_or_near_date(thread, date_tokens, requester):
                        continue
                    key_tokens = _match_tokens(key)
                    if not key_tokens:
                        continue
                    if subj_inc_set:
                        key_inc_set = _inc_tokens(key)
                        if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                            continue
                    key_num_set = _sig_num_tokens(key)
                    if subj_num_set and key_num_set and subj_num_set.isdisjoint(key_num_set):
                        continue
                    score = _token_overlap_score(subj_tokens, key_tokens)
                    contains = (subject_norm in key or key in subject_norm)
                    if score < 0.4 and not contains:
                        continue
                    candidate = (-score, key, thread)
                    if best is None or candidate < best:
                        best = candidate
                if best:
                    _, key, thread = best
                    return thread, f"ConsultantDateGlobal:{key}"

        # If we have an explicit date marker and a requester, prefer threads that
        # contain a consultant reply on that date (whole mail chain search).
        if prefer_consultant_date and date_tokens and requester and len(candidates) > 1:
            consultant_date_candidates = []
            for score, key, thread, note in candidates:
                if _thread_has_consultant_on_or_near_date(thread, date_tokens, requester):
                    consultant_date_candidates.append((score, key, thread, f"{note}+ConsultantDate"))
            if consultant_date_candidates:
                candidates = consultant_date_candidates

        # If the subject contains an explicit date and we still have multiple candidates,
        # prefer threads that contain the same date token in subject/body.
        if date_tokens and len(candidates) > 1:
            date_matches = []
            for score, key, thread, note in candidates:
                if _thread_has_date_token(thread, date_tokens):
                    date_matches.append((score, key, thread, f"{note}+Date"))
            if date_matches:
                candidates = date_matches

        # Compute distinguishing subject tokens over the current candidate set.
        # These tokens are used as compatibility signals for requester-based
        # resolution and low-score fallback disambiguation.
        variant_stop2 = {
            "re", "fw", "fwd", "aw", "wg", "sv",
            "in", "on", "at", "to", "of", "for", "not",
            "and", "or", "the", "a", "an", "is", "it",
            "by", "as", "if", "no", "we", "us",
            "today", "below", "details", "file", "files",
            "input", "output", "received", "failed",
        }
        cand_token_map = {}
        cand_counts = {}
        for _, key, _, _ in candidates:
            toks = _match_tokens(key)
            cand_token_map[key] = toks
            for t in toks:
                cand_counts[t] = cand_counts.get(t, 0) + 1
        discr_tokens = {
            t for t in subj_tokens
            if t not in variant_stop2 and 0 < cand_counts.get(t, 0) < len(candidates)
        }
        max_discr_overlap = 0
        if discr_tokens:
            for _, key, _, _ in candidates:
                ov = len(discr_tokens & cand_token_map.get(key, set()))
                if ov > max_discr_overlap:
                    max_discr_overlap = ov
        max_short_overlap = 0
        if short_tokens:
            for _, key, _, _ in candidates:
                ov = len(short_tokens & cand_token_map.get(key, set()))
                if ov > max_short_overlap:
                    max_short_overlap = ov

        # Resolve by requester only if it appears in exactly one compatible candidate.
        req = (requester or "").strip().lower()
        if req:
            req_hits = []
            for _, key, thread, _ in sorted(candidates, key=lambda x: (-x[0], -len(x[1]))):
                if short_tokens and max_short_overlap > 0:
                    ov_short = len(short_tokens & cand_token_map.get(key, set()))
                    if ov_short < max_short_overlap:
                        continue
                if any(req in (e.sender_name or "").lower() or req in (e.sender_email or "").lower() for e in thread):
                    if discr_tokens and max_discr_overlap > 0:
                        ov = len(discr_tokens & cand_token_map.get(key, set()))
                        if ov < max_discr_overlap:
                            continue
                    req_hits.append((key, thread))
            if len(req_hits) == 1:
                key, thread = req_hits[0]
                return thread, f"AmbiguousResolvedByRequester:{key}"

        # Pick best scored candidate
        candidates.sort(key=lambda x: (-x[0], -len(x[1])))
        top = candidates[0]
        if len(candidates) == 1:
            return top[2], f"{top[3]}:{top[1]}"

        # Smart low-score fallback: prefer candidate with best combined evidence
        # instead of returning unknown too aggressively.
        if len(candidates) > 1 and top[0] < 0.75:
            def _evidence(c):
                score, key, thread, _note = c
                ev = score
                if req and any(req in (e.sender_name or "").lower() or req in (e.sender_email or "").lower() for e in thread):
                    ev += 0.08
                if date_tokens and _thread_has_date_token(thread, date_tokens):
                    ev += 0.06
                if discr_tokens and max_discr_overlap > 0:
                    ov = len(discr_tokens & cand_token_map.get(key, set()))
                    ev += (0.10 * ov / max_discr_overlap)
                if _thread_has_non_ess(thread):
                    ev += 0.02
                return ev

            ranked = sorted(candidates, key=lambda c: (-_evidence(c), -len(c[1])))
            best = ranked[0]
            second = ranked[1] if len(ranked) > 1 else None
            best_ev = _evidence(best)
            second_ev = _evidence(second) if second else -1

            if (best_ev - second_ev) >= 0.05:
                return best[2], f"{best[3]}+SmartFallback:{best[1]}"

            if discr_tokens and max_discr_overlap > 0:
                best_ov = len(discr_tokens & cand_token_map.get(best[1], set()))
                second_ov = len(discr_tokens & cand_token_map.get(second[1], set())) if second else -1
                if best_ov > second_ov:
                    return best[2], f"{best[3]}+SmartFallback:{best[1]}"

            if exact_thread is not None:
                return exact_thread, exact_note
            return [], f"Ambiguous:{len(candidates)}"

        return top[2], f"{top[3]}:{top[1]}"

    def resolve_row(row_context):
        nonlocal created_history
        skip_history_update = False
        description = row_context.get("Description", "")
        requester = row_context.get("Consultant", "") or row_context.get("Requester", "")
        category_type = row_context.get("Category Type", "")
        row_index = row_context.get("RowIndex")
        # Defaults used in row_states even when deployment override fires
        date_anchor_missing = False
        date_anchor_after = False
        base_times = None
        base_debug = None
        group_total = 0

        subject_text = _subject_for_description(description)
        # Interface tokens from description prefix help disambiguate
        # when the subject text itself loses the interface prefix.
        desc_prefix = ""
        if description and re.search(r"(?:--?>|â†’|âž”|âž¡|Ã¢â€ â€™|Ã¢Å¾â€|Ã¢Å¾Â¡|=>)", description):
            parts = re.split(r"\s*(?:--?>|â†’|âž”|âž¡|Ã¢â€ â€™|Ã¢Å¾â€|Ã¢Å¾Â¡|=>)\s*", description, maxsplit=1)
            if len(parts) >= 2:
                desc_prefix = parts[0].strip()
        iface_hint_tokens = _interface_tokens(desc_prefix) if desc_prefix else set()
        subject_norm = normalize_subject(subject_text)
        raw_subject_norm = normalize_subject(description or "")
        date_tokens = _extract_date_tokens(subject_text)
        date_tokens += _extract_date_tokens_from_description(description)
        # de-dup while preserving order
        if date_tokens:
            seen = set()
            deduped = []
            for t in date_tokens:
                if t in seen:
                    continue
                seen.add(t)
                deduped.append(t)
            date_tokens = deduped
        explicit_marker = _has_explicit_date_marker(subject_text) or _has_explicit_date_marker(description)
        stale_anchor = _is_stale_anchor_date(date_tokens, _baseline_row_created_value(row_context))
        if stale_anchor:
            # Subject contains a historical date token (common in long mail chains).
            # Ignore it for matching/anchoring to avoid jumping to an old sibling thread.
            date_tokens = []
            explicit_marker = False
        baseline_created_date = _coerce_row_created_date(_baseline_row_created_value(row_context))
        thread = None
        match_note = ""
        if raw_subject_norm and raw_subject_norm != subject_norm and raw_subject_norm in threads:
            raw_thread = threads.get(raw_subject_norm) or []
            if (not requester) or _thread_has_requester(raw_thread, requester):
                thread = raw_thread
                match_note = "RowExact"
        deployment_like_row = _is_deployment_request_subject(subject_text or "") or _is_deployment_success_subject(subject_text or "")
        if thread is None:
            thread, match_note = find_thread(
                subject_norm,
                requester,
                date_tokens=date_tokens,
                prefer_consultant_date=explicit_marker,
                iface_tokens=iface_hint_tokens,
                baseline_date=baseline_created_date,
            )
        # Thread safety refinement:
        # if selected thread does not contain requester replies, try a requester-backed
        # candidate before resolving times. This prevents "No ESS/requester replies"
        # on wrong sibling threads.
        if (not deployment_like_row) and thread and requester and not _thread_has_requester(thread, requester):
            refined_thread = None
            refined_note = ""
            if date_tokens:
                alt = _find_unique_requester_date_thread(
                    subject_norm,
                    requester,
                    date_tokens,
                    iface_tokens=iface_hint_tokens,
                )
                if alt:
                    refined_thread, refined_note = alt
            if not refined_thread and date_tokens:
                alt = _find_best_requester_date_thread(
                    subject_norm,
                    requester,
                    date_tokens,
                    iface_tokens=iface_hint_tokens,
                )
                if alt:
                    refined_thread, refined_note = alt
            if not refined_thread and baseline_created_date:
                refined = _find_best_thread_near_baseline(
                    subject_norm,
                    requester,
                    baseline_created_date,
                    iface_tokens=iface_hint_tokens,
                )
                if refined:
                    refined_thread, refined_note, _ = refined
            if not refined_thread and requester:
                subj_tokens = _match_tokens(subject_norm)
                subj_inc_set = _inc_tokens(subject_norm)
                best = None
                for key, cand_thread in threads.items():
                    if not _thread_has_requester(cand_thread, requester):
                        continue
                    key_tokens = _match_tokens(key)
                    if not key_tokens:
                        continue
                    if subj_inc_set:
                        key_inc_set = _inc_tokens(key)
                        if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                            continue
                    if iface_hint_tokens:
                        key_iface_set = _interface_tokens(key)
                        if key_iface_set and iface_hint_tokens.isdisjoint(key_iface_set):
                            continue
                    score = _token_overlap_score(subj_tokens, key_tokens) if subj_tokens else 0.0
                    contains = bool(subject_norm and (subject_norm in key or key in subject_norm))
                    # Keep this less strict than initial thread match because this
                    # block runs only after we detected "no requester in selected
                    # thread" and must recover from sibling drift.
                    if score < 0.30 and not contains:
                        continue
                    delta = _requester_min_delta_days(cand_thread, requester, baseline_created_date) if baseline_created_date else 9999
                    if delta is None:
                        delta = 9999
                    candidate = (delta, -score, -len(key_tokens), key, cand_thread)
                    if best is None or candidate < best:
                        best = candidate
                if best:
                    _delta, _neg_score, _neg_len, best_key, best_thread = best
                    refined_thread = best_thread
                    refined_note = f"RequesterRefined:{best_key}"
            if refined_thread and refined_thread is not thread:
                thread = refined_thread
                match_note = f"{match_note}; {refined_note}" if match_note else refined_note
            elif not refined_thread:
                # Keep explicit trace when requester-backed recovery was not found.
                match_note = f"{match_note}; NoRequesterThreadRecovery" if match_note else "NoRequesterThreadRecovery"
        # Baseline thread refinement (safe):
        # If selected thread's requester episode is far from ServiceNow baseline,
        # prefer a closer requester-backed sibling thread.
        # Keep this narrow to avoid disturbing correctly matched rows.
        if (not deployment_like_row) and thread and requester and baseline_created_date:
            current_delta = _requester_min_delta_days(thread, requester, baseline_created_date)
            match_note_l = (match_note or "").lower()
            ambiguous_match = (
                ("score" in match_note_l)
                or ("ambiguous" in match_note_l)
                or ("norequesterthreadrecovery" in match_note_l)
            )
            needs_tight_refine = bool(explicit_marker or date_tokens or ambiguous_match)
            # Default threshold remains conservative, but tighten to >1 day for
            # ambiguous/date-driven rows where 1-2 day drift is still a real miss.
            refine_needed = (
                current_delta is None
                or current_delta > 7
                or (needs_tight_refine and current_delta > 1)
            )
            if refine_needed:
                refined = _find_best_thread_near_baseline(
                    subject_norm,
                    requester,
                    baseline_created_date,
                    iface_tokens=iface_hint_tokens,
                )
                if refined:
                    new_thread, refine_note, new_delta = refined
                    if current_delta is None or new_delta + 1 < current_delta:
                        thread = new_thread
                        match_note = f"{match_note}; BaselineThreadRefined:{refine_note}" if match_note else f"BaselineThreadRefined:{refine_note}"

        # When a stale date marker existed, do one stricter baseline refinement.
        if (not deployment_like_row) and thread and requester and baseline_created_date and stale_anchor:
            current_delta = _requester_min_delta_days(thread, requester, baseline_created_date)
            if current_delta is None or current_delta > 3:
                refined = _find_best_thread_near_baseline(
                    subject_norm,
                    requester,
                    baseline_created_date,
                    iface_tokens=iface_hint_tokens,
                )
                if refined:
                    new_thread, refine_note, new_delta = refined
                    if current_delta is None or new_delta + 1 < current_delta:
                        thread = new_thread
                        match_note = f"{match_note}; {refine_note}"

        thread_key = (
            id(thread) if thread else 0,
            len(thread) if thread else 0,
        )
        env_cache_key = (
            (subject_text or "").strip().lower(),
            _requester_key(requester),
            thread_key,
            (description or "").strip().lower() if not thread else "",
        )
        env = env_cache.get(env_cache_key)
        if env is None:
            consultant_body_text = ""
            if thread and requester:
                consultant_bodies = []
                for e in thread:
                    if not _match_requester(e.sender_name, e.sender_email, requester):
                        continue
                    # Read full consultant content for env detection:
                    # include selected plain body and raw html payload when present.
                    if e.body:
                        consultant_bodies.append(e.body)
                    if getattr(e, "body_html", None):
                        consultant_bodies.append(e.body_html)
                consultant_body_text = "\n".join(consultant_bodies)

            # Environment: subject -> consultant replies -> description.
            # Do not use whole thread as fallback to avoid cross-topic leakage
            # inside long email chains.
            env = resolve_environment(subject_text, consultant_body_text)
            if not env:
                env = resolve_environment(subject_text, description or "")
            env_cache[env_cache_key] = env
        interface_code = resolve_interface_code(description)
        incident_type = resolve_incident_type(category_type, description)
        service_request = resolve_service_request(category_type)

        # Honor explicit sheet hint first when present. This protects rows where
        # category type is wrong in source data, without hardcoding any subject.
        sr_hint = str(
            row_context.get("ServiceRequest/Incident?", "")
            or row_context.get("Servicerequest/incident?", "")
            or ""
        ).strip().lower()
        cat_l = (category_type or "").lower()

        if "file process/data check" in cat_l:
            service_request = "ServiceRequest"
            incident_type = "SR-File process"
        elif sr_hint.startswith("service"):
            service_request = "ServiceRequest"
        elif sr_hint.startswith("incident"):
            service_request = "Incident"
            if not (isinstance(incident_type, str) and incident_type.startswith("INC-")):
                incident_type = "INC-Exceptions"
        else:
            if isinstance(incident_type, str) and incident_type.startswith("INC-"):
                service_request = "Incident"

        deployment_override = None
        is_dep_req = _is_deployment_request_subject(subject_text)
        is_dep_succ = _is_deployment_success_subject(subject_text)
        if is_dep_req or is_dep_succ:
            row_iface = _iface_tokens(subject_text)
            def _dep_env(text: str) -> str:
                if not text:
                    return ""
                s = text.lower()
                if "qa" in s:
                    return "qa"
                if "prod" in s or "production" in s:
                    return "prod"
                if "uat" in s:
                    return "uat"
                return ""
            def _thread_has_sender(thread, email: str) -> bool:
                if not thread or not email:
                    return False
                target = email.lower()
                for e in thread:
                    if (e.sender_email or "").lower() == target:
                        return True
                return False
            row_env = _dep_env(subject_text)
            row_dr_ids = set()
            if thread:
                row_dr_ids |= _thread_dr_ids(thread)
            # Always include DR IDs directly from the row text; thread match can drift.
            row_dr_ids |= _extract_dr_ids(description or "")
            row_dr_ids |= _extract_dr_ids(subject_text or "")

            for dr in row_dr_ids:
                bucket = deployment_index.get(dr)
                if not bucket:
                    continue
                req_candidates = bucket["request"]
                succ_candidates = bucket["success"]
                if not req_candidates:
                    continue

                # Require environment match (PROD/UAT) when present on the row.
                if row_env:
                    req_candidates = [c for c in req_candidates if _dep_env(c["subject_key"]) == row_env]
                    succ_candidates = [c for c in succ_candidates if _dep_env(c["subject_key"]) == row_env]
                    if not req_candidates:
                        continue

                if row_iface:
                    req_candidates = [c for c in req_candidates if row_iface & c["iface"]] or req_candidates
                    succ_candidates = [c for c in succ_candidates if row_iface & c["iface"]] or succ_candidates

                # For ai-assist driven deployments, prefer threads actually sent by ai-assist.
                ai_sender = "ai-assist@umusic.com"
                if any(_thread_has_sender(c["thread"], ai_sender) for c in req_candidates + succ_candidates):
                    req_candidates = [c for c in req_candidates if _thread_has_sender(c["thread"], ai_sender)] or req_candidates
                    succ_candidates = [c for c in succ_candidates if _thread_has_sender(c["thread"], ai_sender)] or succ_candidates

                if is_dep_req:
                    # Always anchor request rows from DR-matched request candidates.
                    req_thread = min(req_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]
                    succ_thread = (
                        min(succ_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]
                        if succ_candidates else None
                    )
                else:
                    if not succ_candidates:
                        continue
                    succ_thread = thread if thread else min(succ_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]
                    req_thread = min(req_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]

                req_email = min(req_thread, key=lambda e: e.sent_time)
                succ_email = min(succ_thread, key=lambda e: e.sent_time) if succ_thread else None
                times = TimeResult(
                    _format_time(req_email.sent_time),
                    _format_time(req_email.sent_time),
                    _format_time(succ_email.sent_time) if succ_email else _format_time(req_email.sent_time),
                )
                debug = TimeDebug(
                    req_email.sender_email or req_email.sender_name,
                    req_email.sender_email or req_email.sender_name,
                    (succ_email.sender_email or succ_email.sender_name) if succ_email else (req_email.sender_email or req_email.sender_name),
                    f"{'DeploymentPair' if succ_email else 'DeploymentRequestOnly'} DR={dr}; Match={match_note}",
                )
                deployment_override = (times, debug)
                break

        if deployment_override:
            times, debug = deployment_override
            if times and debug:
                base_times = TimeResult(times.created, times.response, times.resolved)
                base_debug = TimeDebug(debug.created_src, debug.ack_src, debug.resolved_src, debug.notes)
        else:
            times = None
            debug = None
            requester_key = _requester_key(requester)
            group_key = (subject_norm, requester_key)
            group_total = group_counts.get(group_key, 0)
            has_row_ids = bool(_inc_tokens(subject_text) or _extract_dr_ids(subject_text))
            date_anchor_missing = False
            date_anchor_after = False
            # Date-anchor override (narrow): if row has a date token and the
            # consultant replied on that date, resolve using thread up to that reply.
            if thread and requester and date_tokens:
                pick = None
                date_anchor_exact = False
                date_anchor_near = False
                date_anchor_crossover = False
                anchor_date = _anchor_date(date_tokens)
                # Guard: ignore stale explicit date markers that are far away from
                # the actual matched thread date range.
                if anchor_date and not _anchor_relevant_to_thread(anchor_date, thread):
                    anchor_date = None

                # Prefer same-day consultant replies first.
                consultant_on_date = [
                    e for e in thread
                    if _match_requester(e.sender_name, e.sender_email, requester)
                    and _email_date_matches(e.sent_time, date_tokens)
                ]
                consultant_on_date.sort(key=lambda e: e.sent_time)
                if consultant_on_date:
                    pick = consultant_on_date[-1]
                    date_anchor_exact = True
                # Midnight crossover guard (narrow): if the row has an explicit date marker,
                # prefer the consultant reply closest to midnight (±2h) around the anchor date.
                elif anchor_date and explicit_marker:
                    anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
                    crossover_window = timedelta(hours=2)
                    crossover_candidates = [
                        e for e in thread
                        if _match_requester(e.sender_name, e.sender_email, requester)
                        and abs(_to_ist(e.sent_time) - anchor_start) <= crossover_window
                    ]
                    if crossover_candidates:
                        crossover_candidates.sort(
                            key=lambda e: (
                                abs(_to_ist(e.sent_time) - anchor_start),
                                _to_ist(e.sent_time),
                            )
                        )
                        pick = crossover_candidates[0]
                        date_anchor_near = True
                        date_anchor_crossover = True

                if not pick and anchor_date:
                    # Grace window before the anchor date (e.g., replies just before midnight)
                    anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
                    grace_hours = 12
                    window_start = anchor_start - timedelta(hours=grace_hours)
                    consultant_near_before = [
                        e for e in thread
                        if _match_requester(e.sender_name, e.sender_email, requester)
                        and window_start <= _to_ist(e.sent_time) < anchor_start
                    ]
                    consultant_near_before.sort(key=lambda e: e.sent_time)
                    if consultant_near_before:
                        pick = consultant_near_before[-1]
                        date_anchor_near = True
                    else:
                        consultant_after = [
                            e for e in thread
                            if _match_requester(e.sender_name, e.sender_email, requester)
                            and _email_on_or_after_date(e.sent_time, anchor_date)
                        ]
                        consultant_after.sort(key=lambda e: e.sent_time)
                        if consultant_after:
                            pick = consultant_after[0]
                            date_anchor_after = True
                if pick:
                    sliced_thread = [e for e in thread if e.sent_time <= pick.sent_time]
                    if sliced_thread:
                        times, debug = resolve_times_with_debug(
                            thread=sliced_thread,
                            requester_name=requester,
                            ess_team=ess_team,
                            subject_norm=subject_norm,
                        )
                        debug = TimeDebug(
                            debug.created_src,
                            debug.ack_src,
                            debug.resolved_src,
                            f"{debug.notes}; "
                            f"{'DateAnchor' if date_anchor_exact else ('DateAnchorCrossover' if date_anchor_crossover else ('DateAnchorNear' if date_anchor_near else 'DateAnchorAfter'))}",
                        )
                        # If the row has an explicit date marker (e.g. "-->07-01-2026"),
                        # keep date-specific rows ordered, but do NOT overwrite a
                        # valid request-created time.
                        if explicit_marker and pick:
                            anchor_dt = _to_ist(pick.sent_time)
                            created_dt = _parse_time_str(times.created)
                            created_dt_ist = _to_ist(created_dt) if created_dt else None
                            if not created_dt_ist or created_dt_ist.date() != anchor_dt.date():
                                t = _format_time(pick.sent_time)
                                if t:
                                    # Preserve created if present; only enforce
                                    # response/resolved not before anchored reply.
                                    resp_dt = _parse_time_str(times.response)
                                    res_dt = _parse_time_str(times.resolved)
                                    resp_dt_ist = _to_ist(resp_dt) if resp_dt else None
                                    res_dt_ist = _to_ist(res_dt) if res_dt else None
                                    created_from_parse = isinstance(debug.created_src, str) and debug.created_src.startswith("PARSED_FROM_")
                                    deployment_like = "deployment request" in (subject_text or "").lower()
                                    force_created_to_anchor = bool(created_from_parse)
                                    if (
                                        not force_created_to_anchor
                                        and deployment_like
                                        and created_dt_ist
                                        and abs((anchor_dt.date() - created_dt_ist.date()).days) >= 1
                                    ):
                                        # For explicit-date deployment-like rows, avoid
                                        # carrying stale created dates from quoted history.
                                        force_created_to_anchor = True
                                    new_created = t if force_created_to_anchor else (times.created or t)
                                    new_response = times.response
                                    new_resolved = times.resolved
                                    if not resp_dt_ist or resp_dt_ist < anchor_dt:
                                        new_response = t
                                    if not res_dt_ist or res_dt_ist < anchor_dt:
                                        new_resolved = t
                                    times = TimeResult(new_created, new_response, new_resolved)
                                    note_extra = "; DateAnchorCreatedAdjusted" if force_created_to_anchor else ""
                                    debug = TimeDebug(
                                        (pick.sender_email or pick.sender_name) if force_created_to_anchor else (debug.created_src or (pick.sender_email or pick.sender_name)),
                                        (pick.sender_email or pick.sender_name) if (new_response == t) else debug.ack_src,
                                        (pick.sender_email or pick.sender_name) if (new_resolved == t) else debug.resolved_src,
                                        f"{debug.notes}; DateAnchorOccurrence{note_extra}",
                                    )
                    # DateAnchor ESS-only override (strict):
                    # Only if this subject+consultant occurs once AND the
                    # anchored reply is the latest consultant reply overall.
                    # Additionally, only force all-three-same when the
                    # anchored slice has a single ESS mail. If there are
                    # multiple ESS mails, keep normal span logic to avoid
                    # collapsing all three times.
                    if group_total == 1:
                        consultant_all = [
                            e for e in thread
                            if _match_requester(e.sender_name, e.sender_email, requester)
                        ]
                        consultant_all.sort(key=lambda e: e.sent_time)
                        if consultant_all:
                            consultant_non_ack = [e for e in consultant_all if not _is_ack_like_reply(e)]
                            if consultant_non_ack:
                                latest_all = consultant_non_ack[-1]
                            else:
                                latest_all = None
                            if latest_all and _to_ist(latest_all.sent_time) == _to_ist(pick.sent_time):
                                non_ess_present = any(
                                    not _is_ess_sender(e, ess_team) for e in thread
                                )
                                if not non_ess_present:
                                    ess_in_slice = [
                                        e for e in sliced_thread
                                        if _is_ess_sender(e, ess_team)
                                    ]
                                    if len(ess_in_slice) <= 1:
                                        t = _format_time(pick.sent_time)
                                        if t:
                                            times = TimeResult(t, t, t)
                                            debug = TimeDebug(
                                                pick.sender_email or pick.sender_name,
                                                pick.sender_email or pick.sender_name,
                                                pick.sender_email or pick.sender_name,
                                                f"{debug.notes}; DateAnchorESSOnly",
                                            )
                else:
                    date_anchor_missing = True
            if times is None:
                times, debug = resolve_times_with_debug(
                    thread=thread,
                    requester_name=requester,
                    ess_team=ess_team,
                    subject_norm=subject_norm,
                )
            if date_anchor_missing and times and debug:
                debug = TimeDebug(
                    debug.created_src,
                    debug.ack_src,
                    debug.resolved_src,
                    f"{debug.notes}; DateAnchorMissing",
                )
            if stale_anchor and times and debug:
                debug = TimeDebug(
                    debug.created_src,
                    debug.ack_src,
                    debug.resolved_src,
                    f"{debug.notes}; DateAnchorIgnoredStale",
                )
            if times and debug and base_times is None:
                base_times = TimeResult(times.created, times.response, times.resolved)
                base_debug = TimeDebug(debug.created_src, debug.ack_src, debug.resolved_src, debug.notes)

            # Episode override (narrow): repeated same-subject for same consultant
            # Only apply when base resolution lacks a reliable ack and we have
            # enough distinct consultant replies to map to each occurrence.
            if (
                group_total >= 2
                and not date_tokens
                and not has_row_ids
                and thread
                and requester
                and _ack_missing(times, debug)
            ):
                consultant_replies_all = [
                    e for e in thread
                    if _match_requester(e.sender_name, e.sender_email, requester)
                ]
                consultant_replies = [e for e in consultant_replies_all if not _is_ack_like_reply(e)]
                consultant_replies.sort(key=lambda e: e.sent_time)
                if len(consultant_replies) >= group_total and len(consultant_replies) >= 2:
                    idx = episode_counters.get(group_key, 0)
                    episode_counters[group_key] = idx + 1
                    pick = consultant_replies[min(idx, len(consultant_replies) - 1)]
                    cutoff = pick.sent_time
                    sliced_thread = [e for e in thread if e.sent_time <= cutoff]
                    if sliced_thread:
                        new_times, new_debug = resolve_times_with_debug(
                            thread=sliced_thread,
                            requester_name=requester,
                            ess_team=ess_team,
                            subject_norm=subject_norm,
                        )
                        orig_all_same = (
                            times.created
                            and times.response
                            and times.resolved
                            and times.created == times.response == times.resolved
                        )
                        new_all_same = (
                            new_times.created
                            and new_times.response
                            and new_times.resolved
                            and new_times.created == new_times.response == new_times.resolved
                        )
                        if not orig_all_same or not new_all_same:
                            times = new_times
                            debug = TimeDebug(
                                new_debug.created_src,
                                new_debug.ack_src,
                                new_debug.resolved_src,
                                f"{new_debug.notes}; EpisodeOverride#{idx + 1}",
                            )

            # Single-occurrence latest-reply override (narrow):
            # If there's no date/id hint and no ack was found, use latest consultant
            # reply for response/resolved (created stays as-is).
            if (
                times
                and debug
                and thread
                and requester
                and group_total == 1
                and not date_tokens
                and not has_row_ids
                and _ack_missing(times, debug)
            ):
                consultant_replies_all = [
                    e for e in thread
                    if _match_requester(e.sender_name, e.sender_email, requester)
                ]
                consultant_replies = [e for e in consultant_replies_all if not _is_ack_like_reply(e)]
                consultant_replies.sort(key=lambda e: e.sent_time)
                if consultant_replies:
                    latest = consultant_replies[-1]
                    created_dt = _parse_time_str(times.created)
                    latest_dt = _to_ist(latest.sent_time)
                    created_dt_ist = _to_ist(created_dt) if created_dt else None
                    max_span = timedelta(days=7)
                    within_span = True
                    if created_dt_ist:
                        within_span = (latest_dt - created_dt_ist) <= max_span
                    if within_span and (not created_dt_ist or latest_dt >= created_dt_ist):
                        latest_str = _format_time(latest.sent_time)
                        if latest_str:
                            times = TimeResult(times.created, latest_str, latest_str)
                            debug = TimeDebug(
                                debug.created_src,
                                latest.sender_email or latest.sender_name,
                                latest.sender_email or latest.sender_name,
                                f"{debug.notes}; LatestConsultantReply",
                            )

            # Duplicate de-collision (narrow):
            # If repeated rows for same (subject, consultant) still resolve to an
            # identical time triplet, advance by occurrence using consultant reply
            # sequence inside the same matched thread.
            # This specifically protects repeated long-chain rows where date-marker
            # branches can collapse multiple occurrences to the same pick.
            if (
                times
                and debug
                and thread
                and requester
                and group_total >= 2
                and (date_tokens or explicit_marker or stale_anchor)
                and not is_dep_req
                and not is_dep_succ
                and "maintenance" not in (description or "").lower()
            ):
                current_triplet = (times.created, times.response, times.resolved)
                state = duplicate_group_state.get(group_key, {"seen": 0, "last_triplet": None})
                seen = state.get("seen", 0)
                last_triplet = state.get("last_triplet")

                if seen >= 1 and last_triplet == current_triplet:
                    consultant_replies_all = [
                        e for e in thread
                        if e.sent_time and _match_requester(e.sender_name, e.sender_email, requester)
                    ]
                    consultant_replies = [e for e in consultant_replies_all if not _is_ack_like_reply(e)]
                    consultant_replies.sort(key=lambda e: e.sent_time)

                    if len(consultant_replies) > seen:
                        pick = consultant_replies[seen]
                        sliced_thread = [e for e in thread if e.sent_time <= pick.sent_time]
                        if sliced_thread:
                            new_times, new_debug = resolve_times_with_debug(
                                thread=sliced_thread,
                                requester_name=requester,
                                ess_team=ess_team,
                                subject_norm=subject_norm,
                            )
                            new_triplet = (new_times.created, new_times.response, new_times.resolved)
                            if new_triplet != current_triplet:
                                times = new_times
                                debug = TimeDebug(
                                    new_debug.created_src,
                                    new_debug.ack_src,
                                    new_debug.resolved_src,
                                    f"{new_debug.notes}; DuplicateOccurrenceAdjust#{seen + 1}",
                                )
                                current_triplet = new_triplet

                duplicate_group_state[group_key] = {
                    "seen": seen + 1,
                    "last_triplet": current_triplet,
                }

            # Outlier-only ordering guard (narrow): adjust only when clearly off from
            # the local flow (median of last 3 created times).
            row_has_ids = has_row_ids or bool(_inc_tokens(description) or _extract_dr_ids(description))
            date_tokens_match_thread = bool(date_tokens) and _thread_has_date_token(thread, date_tokens) if thread else False
            skip_history_update = False
            if (
                times
                and debug
                and thread
                and requester
                and created_history
                and (not date_tokens or date_anchor_missing or date_anchor_after or not date_tokens_match_thread)
                and (not row_has_ids or date_anchor_missing or date_anchor_after)
                and not is_dep_req
                and not is_dep_succ
                and "maintenance" not in (description or "").lower()
            ):
                ess_only_no_request = "ESS-only; no non-ESS request" in (debug.notes or "")
                created_dt = _parse_time_str(times.created)
                created_dt_ist = _to_ist(created_dt) if created_dt else None
                if created_dt_ist:
                    orig_all_same = (
                        times.created
                        and times.response
                        and times.resolved
                        and times.created == times.response == times.resolved
                    )
                    recent = created_history[-3:]
                    # If recent history is already monotonic, anchor to the latest
                    # timestamp to keep ordering tight. Otherwise use the median
                    # to avoid reacting to a single outlier.
                    expected_created = None
                    is_monotonic = False
                    if len(recent) >= 2:
                        is_monotonic = all(recent[i] <= recent[i + 1] for i in range(len(recent) - 1))
                        if is_monotonic:
                            expected_created = recent[-1]
                    if expected_created is None:
                        recent_sorted = sorted(recent)
                        expected_created = recent_sorted[len(recent_sorted) // 2]
                    expected_day_start = expected_created.replace(hour=0, minute=0, second=0, microsecond=0)
                    expected_day_end = expected_day_start + timedelta(days=1)
                    # Tighten the guard when recent history is already monotonic.
                    if is_monotonic:
                        grace_window = timedelta(hours=6)
                        max_forward_window = timedelta(days=1)
                    else:
                        grace_window = timedelta(hours=12)
                        max_forward_window = timedelta(days=2)
                    anchor_min = expected_created - grace_window
                    anchor_max = expected_created + max_forward_window

                    adjusted = False
                    # 1) Day-based correction: if the created date is outside the expected day,
                    # try to align to a consultant reply within the expected day.
                    if created_dt_ist.date() != expected_created.date():
                        day_thread = []
                        for e in thread:
                            if not e.sent_time:
                                continue
                            try:
                                sent_ist = _to_ist(e.sent_time)
                            except Exception:
                                continue
                            if expected_day_start <= sent_ist < expected_day_end:
                                day_thread.append(e)
                        consultant_in_day = [
                            e for e in day_thread
                            if _match_requester(e.sender_name, e.sender_email, requester)
                        ]
                        candidate_in_day = consultant_in_day
                        # For INC/DR rows, require a consultant reply in the window.
                        allow_ess_fallback = not row_has_ids
                        if not candidate_in_day and ess_only_no_request and allow_ess_fallback:
                            candidate_in_day = [
                                e for e in day_thread
                                if _is_ess_sender(e, ess_team)
                            ]
                        if candidate_in_day:
                            pick = min(
                                candidate_in_day,
                                key=lambda e: abs(_to_ist(e.sent_time) - expected_created),
                            )
                            sliced = [e for e in day_thread if e.sent_time <= pick.sent_time]
                            if sliced:
                                new_times, new_debug = resolve_times_with_debug(
                                    thread=sliced,
                                    requester_name=requester,
                                    ess_team=ess_team,
                                    subject_norm=subject_norm,
                                )
                                new_created_dt = _parse_time_str(new_times.created)
                                new_created_dt_ist = _to_ist(new_created_dt) if new_created_dt else None
                                if new_created_dt_ist and new_created_dt_ist.date() == expected_created.date():
                                    orig_delta = abs(created_dt_ist - expected_created)
                                    new_delta = abs(new_created_dt_ist - expected_created)
                                    new_all_same = (
                                        new_times.created
                                        and new_times.response
                                        and new_times.resolved
                                        and new_times.created == new_times.response == new_times.resolved
                                    )
                                    if (not orig_all_same and new_all_same):
                                        # Skip adjustment if it collapses all three times
                                        # when the original did not.
                                        pass
                                    elif new_delta < orig_delta:
                                        times = new_times
                                        debug = TimeDebug(
                                            new_debug.created_src,
                                            new_debug.ack_src,
                                            new_debug.resolved_src,
                                            f"{new_debug.notes}; OrderOutlierAdjustedDay",
                                        )
                                        adjusted = True

                    # 2) Window-based correction (fallback)
                    if not adjusted and (created_dt_ist < anchor_min or created_dt_ist > anchor_max):
                        window_thread = []
                        for e in thread:
                            if not e.sent_time:
                                continue
                            try:
                                sent_ist = _to_ist(e.sent_time)
                            except Exception:
                                continue
                            if anchor_min <= sent_ist <= anchor_max:
                                window_thread.append(e)
                        consultant_in_window = [
                            e for e in window_thread
                            if _match_requester(e.sender_name, e.sender_email, requester)
                        ]
                        candidate_in_window = consultant_in_window
                        # For INC/DR rows, require a consultant reply in the window.
                        allow_ess_fallback = not row_has_ids
                        if not candidate_in_window and ess_only_no_request and allow_ess_fallback:
                            candidate_in_window = [
                                e for e in window_thread
                                if _is_ess_sender(e, ess_team)
                            ]
                        if candidate_in_window:
                            pick = min(
                                candidate_in_window,
                                key=lambda e: abs(_to_ist(e.sent_time) - expected_created),
                            )
                            sliced = [e for e in window_thread if e.sent_time <= pick.sent_time]
                            if sliced:
                                new_times, new_debug = resolve_times_with_debug(
                                    thread=sliced,
                                    requester_name=requester,
                                    ess_team=ess_team,
                                    subject_norm=subject_norm,
                                )
                                new_created_dt = _parse_time_str(new_times.created)
                                new_created_dt_ist = _to_ist(new_created_dt) if new_created_dt else None
                                if new_created_dt_ist and anchor_min <= new_created_dt_ist <= anchor_max:
                                    orig_delta = abs(created_dt_ist - expected_created)
                                    new_delta = abs(new_created_dt_ist - expected_created)
                                    new_all_same = (
                                        new_times.created
                                        and new_times.response
                                        and new_times.resolved
                                        and new_times.created == new_times.response == new_times.resolved
                                    )
                                    if (not orig_all_same and new_all_same):
                                        # Skip adjustment if it collapses all three times
                                        # when the original did not.
                                        pass
                                    elif new_delta < orig_delta:
                                        times = new_times
                                        debug = TimeDebug(
                                            new_debug.created_src,
                                            new_debug.ack_src,
                                            new_debug.resolved_src,
                                            f"{new_debug.notes}; OrderOutlierAdjusted",
                                        )
                                        adjusted = True

                    if (
                        is_monotonic
                        and not adjusted
                        and (created_dt_ist < anchor_min or created_dt_ist > anchor_max)
                    ):
                        # Don't let a clear outlier poison the monotonic anchor.
                        skip_history_update = True
                        debug = TimeDebug(
                            debug.created_src,
                            debug.ack_src,
                            debug.resolved_src,
                            f"{debug.notes}; OrderOutlierSkippedHistory",
                        )

        # Post-ack resolution update (narrow): if a later consultant reply exists,
        # move only "Resolved" to the first reply after ack.
        if (
            times
            and debug
            and thread
            and requester
            and not is_dep_req
            and not is_dep_succ
            and "Maintenance override" not in (debug.notes or "")
        ):
            # Include related threads with the same normalized subject-for-match
            # to catch consultant replies that landed under a prefixed subject.
            related_threads = [thread]
            used_related = False
            try:
                alt_subject = normalize_subject_for_match(subject_norm)
                if alt_subject and alt_subject in alt_index:
                    for k in alt_index[alt_subject]:
                        t = threads.get(k)
                        if t and t is not thread:
                            related_threads.append(t)
                            used_related = True
            except Exception:
                pass

            merged = []
            seen = set()
            for t in related_threads:
                for e in t:
                    dedup_key = (e.subject, e.sender_email, e.sent_time)
                    if dedup_key in seen:
                        continue
                    seen.add(dedup_key)
                    merged.append(e)

            consultant_replies = [
                e for e in merged
                if _match_requester(e.sender_name, e.sender_email, requester)
            ]
            notes_l_now = ((debug.notes or "").lower() if debug else "")
            skip_resolved_after_ack = (
                "failed subject; no ack phrase" in notes_l_now
                or "no ess or requester replies" in notes_l_now
            )
            if consultant_replies and not skip_resolved_after_ack:
                ack_dt = _parse_time_str(times.response)
                res_dt = _parse_time_str(times.resolved)
                if ack_dt:
                    enable_postack_fallback_30m = os.getenv("POSTACK_FALLBACK_30M", "0") == "1"
                    ack_ist = _to_ist(ack_dt)
                    res_ist = _to_ist(res_dt) if res_dt else None
                    # If resolved is missing or effectively equals ack (same/very close),
                    # but there is a later consultant reply, move resolved to that reply.
                    if (not res_ist) or (res_ist <= (ack_ist + timedelta(minutes=20))):
                        window_hours = 48
                        window_end = ack_ist + timedelta(hours=window_hours)
                        anchor_date = _anchor_date(date_tokens or [])
                        if anchor_date:
                            anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
                            window_end = max(window_end, anchor_start + timedelta(hours=36))
                        consultant_after = [
                            e for e in consultant_replies
                            if _to_ist(e.sent_time) > ack_ist
                            and _to_ist(e.sent_time) <= window_end
                            and not _is_ack_like_reply(e)
                        ]
                        if anchor_date:
                            on_anchor = [
                                e for e in consultant_after
                                if _to_ist(e.sent_time).date() == anchor_date
                            ]
                            if on_anchor:
                                consultant_after = on_anchor
                        consultant_after.sort(key=lambda e: e.sent_time)
                        fallback_30m_used = False
                        # If no consultant reply found, broaden to related threads
                        # with strong subject overlap that also contain the requester.
                        if not consultant_after:
                            try:
                                min_score = 0.72
                                if explicit_marker and anchor_date:
                                    min_score = 0.60
                                subj_tokens = _match_tokens(subject_norm)
                                subj_inc_set = _inc_tokens(subject_norm)
                                for key, t in threads.items():
                                    if t in related_threads:
                                        continue
                                    if not _thread_has_requester(t, requester):
                                        continue
                                    key_tokens = _match_tokens(key)
                                    if not key_tokens:
                                        continue
                                    if subj_inc_set:
                                        key_inc_set = _inc_tokens(key)
                                        if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                                            continue
                                    score = _token_overlap_score(subj_tokens, key_tokens)
                                    contains = (subject_norm in key or key in subject_norm)
                                    if score < min_score and not contains:
                                        continue
                                    related_threads.append(t)
                                    used_related = True
                                if used_related:
                                    merged = []
                                    seen = set()
                                    for t in related_threads:
                                        for e in t:
                                            dedup_key = (e.subject, e.sender_email, e.sent_time)
                                            if dedup_key in seen:
                                                continue
                                            seen.add(dedup_key)
                                            merged.append(e)
                                    consultant_replies = [
                                        e for e in merged
                                        if _match_requester(e.sender_name, e.sender_email, requester)
                                    ]
                                    consultant_after = [
                                        e for e in consultant_replies
                                        if _to_ist(e.sent_time) > ack_ist
                                        and _to_ist(e.sent_time) <= window_end
                                        and not _is_ack_like_reply(e)
                                    ]
                                    if anchor_date:
                                        on_anchor = [
                                            e for e in consultant_after
                                            if _to_ist(e.sent_time).date() == anchor_date
                                        ]
                                        if on_anchor:
                                            consultant_after = on_anchor
                                    consultant_after.sort(key=lambda e: e.sent_time)
                            except Exception:
                                pass

                        # Conservative fallback: only when strict non-ack filter yields
                        # no candidate; choose first requester reply >=30m after ack.
                        if not consultant_after and enable_postack_fallback_30m:
                            fallback_after = [
                                e for e in consultant_replies
                                if _to_ist(e.sent_time) >= (ack_ist + timedelta(minutes=30))
                                and _to_ist(e.sent_time) <= window_end
                                and not _is_ack_like_reply(e)
                            ]
                            if anchor_date:
                                on_anchor = [
                                    e for e in fallback_after
                                    if _to_ist(e.sent_time).date() == anchor_date
                                ]
                                if on_anchor:
                                    fallback_after = on_anchor
                            fallback_after.sort(key=lambda e: e.sent_time)
                            if fallback_after:
                                consultant_after = fallback_after
                                fallback_30m_used = True

                        if consultant_after:
                            next_reply = _pick_reply_after_ack(consultant_after, ack_ist)
                            next_str = _format_time(next_reply.sent_time)
                            if next_str:
                                times = TimeResult(times.created, times.response, next_str)
                                debug = TimeDebug(
                                    debug.created_src,
                                    debug.ack_src,
                                    next_reply.sender_email or next_reply.sender_name,
                                    f"{debug.notes}; "
                                    f"{'ResolvedAfterAckFallback30m' if fallback_30m_used else ('ResolvedAfterAckRelated' if used_related else 'ResolvedAfterAck')}",
                                )

        # Maintenance override: if consultant replied, set all three times to the
        # consultant's last reply time. Do not mark red.
        if "maintenance" in (description or "").lower() and thread and requester:
            consultant_replies = [
                e for e in thread
                if _match_requester(e.sender_name, e.sender_email, requester)
            ]
            if consultant_replies:
                last_reply = max(consultant_replies, key=lambda e: e.sent_time)
                times = TimeResult(
                    _format_time(last_reply.sent_time),
                    _format_time(last_reply.sent_time),
                    _format_time(last_reply.sent_time),
                )
                debug = TimeDebug(
                    last_reply.sender_email or last_reply.sender_name,
                    last_reply.sender_email or last_reply.sender_name,
                    last_reply.sender_email or last_reply.sender_name,
                    "Maintenance override; consultant reply",
                )

        if times and debug:
            times, debug = _apply_final_time_guards(
                thread=thread,
                requester=requester,
                times=times,
                debug=debug,
                base_times=base_times,
                base_debug=base_debug,
                baseline_created_date=baseline_created_date,
                is_dep_req=is_dep_req,
                is_dep_succ=is_dep_succ,
            )

        final_created_dt = _parse_time_str(times.created)
        final_created_dt_ist = _to_ist(final_created_dt) if final_created_dt else None
        if final_created_dt_ist and not skip_history_update:
            created_history.append(final_created_dt_ist)

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

        state_row_has_ids = bool(
            _inc_tokens(subject_text)
            or _extract_dr_ids(subject_text)
            or _inc_tokens(description)
            or _extract_dr_ids(description)
        )
        row_states.append(
            {
                "row_index": row_index,
                "list_index": len(automation_rows) - 1,
                "description": description,
                "category_type": category_type,
                "requester": requester,
                "subject_norm": subject_norm,
                "date_tokens": date_tokens,
                "explicit_marker": explicit_marker,
                "date_anchor_missing": date_anchor_missing,
                "date_anchor_after": date_anchor_after,
                "stale_anchor": stale_anchor,
                "baseline_created_date": baseline_created_date,
                "group_total": group_total,
                "date_tokens_match_thread": bool(date_tokens) and _thread_has_date_token(thread, date_tokens),
                "thread": thread,
                "times": times,
                "debug": debug,
                "row_has_ids": state_row_has_ids,
                "is_dep_req": is_dep_req,
                "is_dep_succ": is_dep_succ,
                "match_note": match_note,
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

    def _sequence_ordering_pass(ws, col_map, header_row):
        created_col = col_map.get("created date & time")
        response_col = col_map.get("actual response date & time")
        resolved_col = col_map.get("actual resolved date & time")
        env_col = col_map.get("environment")
        if not created_col or not response_col or not resolved_col:
            return

        # Performance caches (no logic change).
        _ist_cache = {}
        _req_match_cache = {}
        _is_ess_cache = {}
        _ack_like_cache = {}
        _emails_for_requester_cache = {}

        def _email_ist(e):
            key = id(e)
            if key in _ist_cache:
                return _ist_cache[key]
            try:
                v = _to_ist(e.sent_time) if getattr(e, "sent_time", None) else None
            except Exception:
                v = None
            _ist_cache[key] = v
            return v

        def _req_match(e, requester_name):
            key = (id(e), requester_name or "")
            if key in _req_match_cache:
                return _req_match_cache[key]
            v = _match_requester(e.sender_name, e.sender_email, requester_name)
            _req_match_cache[key] = v
            return v

        def _ess_sender(e):
            key = id(e)
            if key in _is_ess_cache:
                return _is_ess_cache[key]
            v = _is_ess_sender(e, ess_team)
            _is_ess_cache[key] = v
            return v

        def _ack_like(e):
            key = id(e)
            if key in _ack_like_cache:
                return _ack_like_cache[key]
            v = _is_ack_like_reply(e)
            _ack_like_cache[key] = v
            return v

        def _ack_like_text_fallback(e):
            # Defensive fallback for rows where plain-text body parsing misses
            # reminder/update phrases but HTML still contains them.
            txt = f"{getattr(e, 'body', '') or ''}\n{getattr(e, 'body_html', '') or ''}".lower()
            if not txt:
                return False
            markers = (
                "could you please provide us an update regarding the below",
                "could you please provide us an update on the below",
                "please provide us an update regarding the below",
                "please provide us an update on the below",
                "thank you for the information",
                "thanks for the information",
                "thank you for the update",
                "thanks for the update",
                "thanks for the info",
                "noted with thanks",
                "duly noted",
            )
            return any(m in txt for m in markers)

        def _parse_quoted_sent_time(sent_line: str):
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

        def _extract_quoted_requester_reply_ist(email_obj, requester_name: str, subject_norm_value: str, lower_ist, upper_ist):
            raw = f"{getattr(email_obj, 'body', '') or ''}\n{getattr(email_obj, 'body_html', '') or ''}"
            if not raw:
                return None
            txt = raw
            txt = re.sub(r"(?is)<style.*?>.*?</style>", " ", txt)
            txt = re.sub(r"(?is)<script.*?>.*?</script>", " ", txt)
            txt = re.sub(r"(?i)<\s*br\s*/?>", "\n", txt)
            txt = re.sub(r"(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>", "\n", txt)
            txt = re.sub(r"(?is)<[^>]+>", " ", txt)
            txt = html.unescape(txt)
            lines = [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]
            if not lines:
                return None

            row_tokens = _match_tokens(subject_norm_value or "")
            cands = []
            for i, ln in enumerate(lines):
                if not ln.lower().startswith("from:"):
                    continue
                from_line = ln[5:].strip()
                sent_line = ""
                subj_line = ""
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
                if requester_name and not _match_requester(from_line, from_line, requester_name):
                    continue
                if subj_line and row_tokens:
                    subj_text = re.sub(r"(?i)^(subject|objet)\s*:\s*", "", subj_line).strip()
                    subj_norm = normalize_subject(subj_text)
                    subj_tokens = _match_tokens(subj_norm)
                    score = _token_overlap_score(row_tokens, subj_tokens) if subj_tokens else 0.0
                    contains = bool(subject_norm_value and subj_norm and (subject_norm_value in subj_norm or subj_norm in subject_norm_value))
                    if score < 0.45 and not contains:
                        continue
                sent_dt = _parse_quoted_sent_time(sent_line)
                if not sent_dt:
                    continue
                sent_ist = _to_ist(sent_dt)
                # Keep this permissive: created/ack can already be stale in broken rows.
                if lower_ist and sent_ist < (lower_ist - timedelta(days=5)):
                    continue
                if upper_ist and sent_ist >= upper_ist:
                    continue
                cands.append(sent_ist)
            if not cands:
                return None
            cands.sort()
            return cands[-1]

        def _extract_quoted_request_before_ist(email_obj, subject_norm_value: str, upper_ist):
            raw = f"{getattr(email_obj, 'body', '') or ''}\n{getattr(email_obj, 'body_html', '') or ''}"
            if not raw:
                return None
            txt = raw
            txt = re.sub(r"(?is)<style.*?>.*?</style>", " ", txt)
            txt = re.sub(r"(?is)<script.*?>.*?</script>", " ", txt)
            txt = re.sub(r"(?i)<\s*br\s*/?>", "\n", txt)
            txt = re.sub(r"(?i)</\s*(p|div|tr|td|th|li|h[1-6])\s*>", "\n", txt)
            txt = re.sub(r"(?is)<[^>]+>", " ", txt)
            txt = html.unescape(txt)
            lines = [ln.strip() for ln in txt.splitlines() if ln and ln.strip()]
            if not lines:
                return None

            row_tokens = _match_tokens(subject_norm_value or "")
            cands = []
            for i, ln in enumerate(lines):
                if not ln.lower().startswith("from:"):
                    continue
                sent_line = ""
                subj_line = ""
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
                if subj_line and row_tokens:
                    subj_text = re.sub(r"(?i)^(subject|objet)\s*:\s*", "", subj_line).strip()
                    subj_norm = normalize_subject(subj_text)
                    subj_tokens = _match_tokens(subj_norm)
                    score = _token_overlap_score(row_tokens, subj_tokens) if subj_tokens else 0.0
                    contains = bool(subject_norm_value and subj_norm and (subject_norm_value in subj_norm or subj_norm in subject_norm_value))
                    if score < 0.45 and not contains:
                        continue
                sent_dt = _parse_quoted_sent_time(sent_line)
                if not sent_dt:
                    continue
                sent_ist = _to_ist(sent_dt)
                if upper_ist and sent_ist >= upper_ist:
                    continue
                cands.append(sent_ist)
            if not cands:
                return None
            cands.sort()
            return cands[-1]

        def _emails_for_requester(requester_name):
            k = requester_name or ""
            if k in _emails_for_requester_cache:
                return _emails_for_requester_cache[k]
            out = [e for e in emails if getattr(e, "sent_time", None) and _req_match(e, requester_name)]
            _emails_for_requester_cache[k] = out
            return out

        _expanded_thread_cache = {}
        _requester_pool_cache = {}

        def _expanded_thread(subject_norm_value, base_thread, requester_name, include_non_ess=False, reference_ist=None):
            if not base_thread:
                return []
            ref_day = ""
            if reference_ist:
                try:
                    ref_day = reference_ist.date().isoformat()
                except Exception:
                    ref_day = ""
            cache_key = (subject_norm_value or "", requester_name or "", id(base_thread), bool(include_non_ess), ref_day)
            if cache_key in _expanded_thread_cache:
                return _expanded_thread_cache[cache_key]

            merged = list(base_thread)
            try:
                alt_subject = normalize_subject_for_match(subject_norm_value or "")
                if alt_subject and alt_subject in alt_index:
                    for key in alt_index[alt_subject]:
                        t = threads.get(key) or []
                        if not t or t is base_thread:
                            continue
                        if requester_name and not any(_req_match(e, requester_name) for e in t):
                            if include_non_ess:
                                has_non_ess = any((not _ess_sender(e)) for e in t)
                                if not has_non_ess:
                                    continue
                            else:
                                continue
                        if include_non_ess and reference_ist:
                            near = False
                            for e in t:
                                e_ist = _email_ist(e)
                                if not e_ist:
                                    continue
                                if abs((e_ist - reference_ist).total_seconds()) <= (14 * 24 * 3600):
                                    near = True
                                    break
                            if not near:
                                continue
                        merged.extend(t)
            except Exception:
                pass

            dedup = {}
            for e in merged:
                k = (e.subject, e.sender_email, e.sender_name, e.sent_time)
                dedup[k] = e
            out = list(dedup.values())
            out.sort(key=lambda e: (0 if getattr(e, "sent_time", None) else 1, getattr(e, "sent_time", datetime.max)))
            _expanded_thread_cache[cache_key] = out
            return out

        def _system_like_sender(e):
            sender = f"{getattr(e, 'sender_email', '') or ''} {getattr(e, 'sender_name', '') or ''}".lower()
            if not sender.strip():
                return False
            markers = (
                "system-notification",
                "system notification",
                "no-reply",
                "noreply",
                "do-not-reply",
                "donotreply",
                "mailer-daemon",
                "postmaster",
            )
            return any(m in sender for m in markers)

        def _requester_pool(subject_norm_value, requester_name, center_ist=None, day_window=21):
            center_key = ""
            if center_ist:
                try:
                    center_key = center_ist.date().isoformat()
                except Exception:
                    center_key = ""
            key = (subject_norm_value or "", requester_name or "", center_key, int(day_window))
            if key in _requester_pool_cache:
                return _requester_pool_cache[key]

            row_tokens = _match_tokens(subject_norm_value or "")
            pool = []
            for e in emails:
                if not getattr(e, "sent_time", None):
                    continue
                if requester_name and not _req_match(e, requester_name):
                    continue
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if center_ist and abs((e_ist - center_ist).total_seconds()) > (day_window * 24 * 3600):
                    continue
                if row_tokens:
                    s_norm = normalize_subject(e.subject or "")
                    s_tokens = _match_tokens(s_norm)
                    score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                    contains = bool(subject_norm_value and s_norm and (subject_norm_value in s_norm or s_norm in subject_norm_value))
                    if score < 0.45 and not contains:
                        continue
                pool.append(e)
            pool.sort(key=lambda e: e.sent_time)
            _requester_pool_cache[key] = pool
            return pool

        # Build created-time list in row order.
        created_list = []
        for state in row_states:
            idx = state["list_index"]
            created_str = automation_rows[idx].get("Created Date & Time")
            created_dt = _parse_time_str(created_str)
            created_list.append(_to_ist(created_dt) if created_dt else None)

        def _find_prev(i):
            for j in range(i - 1, -1, -1):
                if created_list[j]:
                    return j, created_list[j]
            return None, None

        def _find_next(i):
            for j in range(i + 1, len(created_list)):
                if created_list[j]:
                    return j, created_list[j]
            return None, None

        grace = timedelta(hours=6)
        max_forward = timedelta(days=1)

        for i, state in enumerate(row_states):
            thread = state.get("thread")
            requester = state.get("requester")
            if not thread or not requester:
                continue

            description = state.get("description") or ""
            if "maintenance" in description.lower():
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            date_tokens = state.get("date_tokens") or []
            explicit_marker = state.get("explicit_marker")
            group_total = state.get("group_total") or 0
            row_has_ids = state.get("row_has_ids")
            debug = state.get("debug")
            debug_notes = (debug.notes if debug else "") or ""
            date_anchor_missing = state.get("date_anchor_missing") or ("DateAnchorMissing" in debug_notes)
            date_anchor_after = state.get("date_anchor_after") or ("DateAnchorAfter" in debug_notes)
            date_tokens_match_thread = state.get("date_tokens_match_thread")

            # Keep per-occurrence duplicate handling stable for repeated explicit-date rows.
            if group_total >= 2 and explicit_marker:
                continue

            # Skip INC/DR rows unless anchor was missing/after OR the row has an
            # explicit marker (safe path for ordered repeated ID rows).
            if row_has_ids and not (date_anchor_missing or date_anchor_after or explicit_marker):
                continue

            prev_idx, prev_time = _find_prev(i)
            next_idx, next_time = _find_next(i)
            if not prev_time and not next_time:
                continue

            created_dt = created_list[i]
            if prev_time and next_time and prev_time <= next_time:
                window_start = prev_time - grace
                window_end = next_time + grace
                expected_center = prev_time + (next_time - prev_time) / 2
            elif prev_time:
                window_start = prev_time - grace
                window_end = prev_time + max_forward
                expected_center = prev_time
            else:
                window_start = next_time - max_forward
                window_end = next_time + grace
                expected_center = next_time

            if created_dt and window_start <= created_dt <= window_end:
                continue

            # If date tokens exist and match the thread, treat them as strong only
            # when the current created time already falls within the expected window.
            # Otherwise allow a sequence-based correction.
            if date_tokens and date_tokens_match_thread and not (date_anchor_missing or date_anchor_after):
                if created_dt and window_start <= created_dt <= window_end:
                    continue

            # Collect candidate replies in the expected window.
            window_thread = []
            for e in thread:
                sent_ist = _email_ist(e)
                if not sent_ist:
                    continue
                if window_start <= sent_ist <= window_end:
                    window_thread.append(e)

            consultant_in_window = [
                e for e in window_thread
                if _req_match(e, requester)
            ]
            candidate_in_window = consultant_in_window

            ess_only_no_request = "ESS-only; no non-ESS request" in debug_notes
            # Safety guard: do not let ESS-window fallback pick another consultant's
            # thread slice when this requester already has replies elsewhere in thread.
            if not candidate_in_window and ess_only_no_request and not row_has_ids:
                requester_exists_in_thread = any(_req_match(e, requester) for e in thread)
                if not requester_exists_in_thread:
                    candidate_in_window = [
                        e for e in window_thread
                        if _ess_sender(e)
                    ]

            if not candidate_in_window:
                continue

            pick = min(
                candidate_in_window,
                key=lambda e: abs(_email_ist(e) - expected_center),
            )
            sliced = [e for e in window_thread if e.sent_time <= pick.sent_time]
            if not sliced:
                continue

            new_times, new_debug = resolve_times_with_debug(
                thread=sliced,
                requester_name=requester,
                ess_team=ess_team,
                subject_norm=state.get("subject_norm"),
            )

            new_created_dt = _parse_time_str(new_times.created)
            new_created_dt_ist = _to_ist(new_created_dt) if new_created_dt else None
            if not new_created_dt_ist:
                continue
            if not (window_start <= new_created_dt_ist <= window_end):
                continue

            orig_times = state.get("times")
            orig_all_same = (
                orig_times.created
                and orig_times.response
                and orig_times.resolved
                and orig_times.created == orig_times.response == orig_times.resolved
            )
            new_all_same = (
                new_times.created
                and new_times.response
                and new_times.resolved
                and new_times.created == new_times.response == new_times.resolved
            )
            if (not orig_all_same and new_all_same):
                # Avoid collapsing timelines that were already distinct.
                continue

            # Apply updates to in-memory rows
            list_index = state["list_index"]
            automation_rows[list_index]["Created Date & Time"] = new_times.created
            automation_rows[list_index]["Actual Response Date & Time"] = new_times.response
            automation_rows[list_index]["Actual Resolved Date & Time"] = new_times.resolved

            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = new_debug.created_src
                debug_rows[list_index]["AckSource"] = new_debug.ack_src
                debug_rows[list_index]["ResolvedSource"] = new_debug.resolved_src
                debug_rows[list_index]["Notes"] = f"{new_debug.notes}; OrderSequenceAdjusted; Match={state.get('match_note')}"

            # Apply updates to sheet
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_times.created
                ws.cell(row_idx, response_col).value = new_times.response
                ws.cell(row_idx, resolved_col).value = new_times.resolved

            # Update local created list to keep subsequent ordering aligned
            created_list[i] = new_created_dt_ist

        # After sequence adjustments, re-apply "resolved-after-ack" using the full thread.
        # This protects cases where the ordering slice stopped before the true resolution reply
        # (e.g., ack just before midnight and resolution shortly after).
        for state in row_states:
            thread = state.get("thread")
            requester = state.get("requester")
            if not thread or not requester:
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            description = state.get("description") or ""
            if "maintenance" in description.lower():
                continue

            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if list_index < len(debug_rows):
                notes_now = (debug_rows[list_index].get("Notes", "") or "").lower()
                if "requester follow-up (no in-between request)" in notes_now:
                    continue
                if "failed subject; no ack phrase" in notes_now:
                    continue
                if "no ess or requester replies" in notes_now:
                    continue

            row_vals = automation_rows[list_index]
            ack_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            res_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            if not ack_dt:
                continue

            try:
                ack_ist = _to_ist(ack_dt)
            except Exception:
                continue
            res_ist = _to_ist(res_dt) if res_dt else None
            if res_ist and res_ist > (ack_ist + timedelta(minutes=20)):
                continue

            consultant_replies = [
                e for e in thread
                if _match_requester(e.sender_name, e.sender_email, requester)
            ]

            window_end = ack_ist + timedelta(hours=48)
            date_tokens = state.get("date_tokens") or []
            explicit_marker = state.get("explicit_marker")
            enable_postack_fallback_30m = os.getenv("POSTACK_FALLBACK_30M", "0") == "1"
            baseline_date = state.get("baseline_created_date")
            created_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            created_ist = _to_ist(created_dt) if created_dt else None
            anchor_date = _anchor_date(date_tokens)
            if anchor_date:
                anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
                window_end = max(window_end, anchor_start + timedelta(hours=36))
            consultant_after = [
                e for e in consultant_replies
                if _email_ist(e) and _email_ist(e) > ack_ist and _email_ist(e) <= window_end
                and not _ack_like(e)
            ]
            # Augment with global scan across all emails to catch the next
            # consultant reply even if the thread grouping missed it.
            subj_norm = state.get("subject_norm") or ""
            subj_tokens = _match_tokens(subj_norm)
            subj_inc_set = _inc_tokens(subj_norm)
            min_score = 0.72
            if explicit_marker and anchor_date:
                min_score = 0.60
            allow_global_postack_scan = bool(explicit_marker or anchor_date)
            if allow_global_postack_scan and len(subj_tokens) >= 3 and len(subj_norm) >= 10:
                global_candidates = []
                for e in _emails_for_requester(requester):
                    if _ack_like(e):
                        continue
                    key_norm = normalize_subject(e.subject or "")
                    key_tokens = _match_tokens(key_norm)
                    if not key_tokens:
                        continue
                    if subj_inc_set:
                        key_inc_set = _inc_tokens(key_norm)
                        if not key_inc_set or subj_inc_set.isdisjoint(key_inc_set):
                            continue
                    score = _token_overlap_score(subj_tokens, key_tokens)
                    contains = (subj_norm in key_norm or key_norm in subj_norm)
                    if score < min_score and not contains:
                        continue
                    sent_ist = _email_ist(e)
                    if not sent_ist:
                        continue
                    if baseline_date and abs((sent_ist.date() - baseline_date).days) > 2:
                        continue
                    if created_ist and sent_ist > (created_ist + timedelta(hours=72)):
                        continue
                    if sent_ist > ack_ist and sent_ist <= window_end:
                        global_candidates.append(e)

                if global_candidates:
                    # Merge global candidates with thread candidates (dedup)
                    seen = set()
                    merged = []
                    for e in consultant_after + global_candidates:
                        dedup_key = (e.subject, e.sender_email, e.sent_time)
                        if dedup_key in seen:
                            continue
                        seen.add(dedup_key)
                        merged.append(e)
                    consultant_after = merged

            if anchor_date and consultant_after:
                on_anchor = [
                    e for e in consultant_after
                    if _email_ist(e) and _email_ist(e).date() == anchor_date
                ]
                if on_anchor:
                    consultant_after = on_anchor

            consultant_after.sort(key=lambda e: e.sent_time)
            fallback_30m_used = False

            if not consultant_after and enable_postack_fallback_30m:
                fallback_after = [
                    e for e in consultant_replies
                    if _email_ist(e) and _email_ist(e) >= (ack_ist + timedelta(minutes=30))
                    and _email_ist(e) <= window_end
                    and not _ack_like(e)
                ]
                if anchor_date:
                    on_anchor = [
                        e for e in fallback_after
                        if _email_ist(e) and _email_ist(e).date() == anchor_date
                    ]
                    if on_anchor:
                        fallback_after = on_anchor
                fallback_after.sort(key=lambda e: e.sent_time)
                if fallback_after:
                    consultant_after = fallback_after
                    fallback_30m_used = True

            if not consultant_after:
                continue

            next_reply = _pick_reply_after_ack(consultant_after, ack_ist)
            if not next_reply:
                continue
            next_str = _format_time(next_reply.sent_time)
            if not next_str:
                continue

            row_vals["Actual Resolved Date & Time"] = next_str
            if list_index < len(debug_rows):
                debug_rows[list_index]["ResolvedSource"] = next_reply.sender_email or next_reply.sender_name
                note_suffix = "ResolvedAfterAckPost"
                if not consultant_replies:
                    note_suffix = "ResolvedAfterAckPostGlobal"
                if fallback_30m_used:
                    note_suffix = "ResolvedAfterAckPostFallback30m"
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; {note_suffix}"

            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, resolved_col).value = next_str

        # Cross-subject duplicate triplet guard (narrow):
        # when same consultant has identical created/ack/resolved across
        # different subjects, try to move only resolved to the next valid
        # consultant non-ack reply in that row's own thread.
        seen_by_requester_triplet = {}
        state_by_list_index = {
            s.get("list_index"): s
            for s in row_states
            if s.get("list_index") is not None
        }
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue
            if list_index < len(debug_rows):
                notes_now = (debug_rows[list_index].get("Notes", "") or "").lower()
                if "requester follow-up (no in-between request)" in notes_now:
                    continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            if not (c and a and r):
                continue

            requester = state.get("requester") or ""
            req_key = _requester_key(requester)
            triplet_key = (req_key, c, a, r)
            prev_idx = seen_by_requester_triplet.get(triplet_key)
            if prev_idx is None:
                seen_by_requester_triplet[triplet_key] = list_index
                continue

            prev_state = state_by_list_index.get(prev_idx)
            prev_subject = (prev_state or {}).get("subject_norm") if prev_state else None
            curr_subject = state.get("subject_norm")
            if not prev_subject or not curr_subject or prev_subject == curr_subject:
                continue

            thread = state.get("thread") or []
            if not thread:
                continue

            res_dt = _parse_time_str(r) or _parse_time_str(a)
            if not res_dt:
                continue
            try:
                res_ist = _to_ist(res_dt)
            except Exception:
                continue

            date_tokens = state.get("date_tokens") or []
            anchor_date = _anchor_date(date_tokens)
            candidates = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and not _ack_like(e)
            ]
            if anchor_date and candidates:
                on_anchor = [e for e in candidates if _email_ist(e) and _email_ist(e).date() == anchor_date]
                if on_anchor:
                    candidates = on_anchor
            if not candidates:
                continue

            candidates.sort(key=lambda e: e.sent_time)
            next_candidates = [
                e for e in candidates
                if _email_ist(e) and _email_ist(e) > res_ist
                and _email_ist(e) <= (res_ist + timedelta(hours=48))
            ]
            if not next_candidates:
                continue

            pick = next_candidates[0]
            new_res = _format_time(pick.sent_time)
            if not new_res or new_res == r:
                continue

            row_vals["Actual Resolved Date & Time"] = new_res
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, resolved_col).value = new_res
            if list_index < len(debug_rows):
                debug_rows[list_index]["ResolvedSource"] = pick.sender_email or pick.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; CrossSubjectDuplicateGuard"

        # Episode consistency guard (conservative):
        # If resolved is far after ack, re-anchor Created/Ack to a newer strong
        # request->ack episode in the same matched thread.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue

            try:
                c_ist = _to_ist(c_dt)
                a_ist = _to_ist(a_dt)
                r_ist = _to_ist(r_dt)
            except Exception:
                continue
            if not (c_ist <= a_ist <= r_ist):
                continue

            old_gap = r_ist - a_ist
            # Only touch rows with clearly stale ack vs resolved.
            if old_gap < timedelta(hours=18):
                continue

            subj_norm = state.get("subject_norm") or ""
            subj_tokens = _match_tokens(subj_norm)
            subj_inc_set = _inc_tokens(subj_norm)
            subj_num_set = _sig_num_tokens(subj_norm)

            episode_candidates = []
            for ack_mail in thread:
                if not ack_mail.sent_time:
                    continue
                if not _req_match(ack_mail, requester):
                    continue
                ack_mail_ist = _email_ist(ack_mail)
                if not ack_mail_ist:
                    continue
                if ack_mail_ist <= a_ist + timedelta(hours=4):
                    continue
                if ack_mail_ist > r_ist:
                    continue
                if not _is_ack_like_reply(ack_mail):
                    continue

                key_norm = normalize_subject(ack_mail.subject or "")
                key_tokens = _match_tokens(key_norm)
                if subj_tokens and key_tokens:
                    score = _token_overlap_score(subj_tokens, key_tokens)
                    contains = (subj_norm in key_norm or key_norm in subj_norm)
                    if score < 0.45 and not contains:
                        continue
                key_inc_set = _inc_tokens(key_norm)
                if subj_inc_set and key_inc_set and subj_inc_set.isdisjoint(key_inc_set):
                    continue
                key_num_set = _sig_num_tokens(key_norm)
                if subj_num_set and key_num_set and subj_num_set.isdisjoint(key_num_set):
                    continue

                req_candidates = [
                    e for e in thread
                    if e.sent_time
                    and not _ess_sender(e)
                    and _email_ist(e) and _email_ist(e) <= ack_mail_ist
                ]
                if not req_candidates:
                    continue
                req_mail = max(req_candidates, key=lambda e: e.sent_time)
                req_ist = _email_ist(req_mail)
                if not req_ist:
                    continue
                if req_ist <= c_ist + timedelta(hours=2):
                    continue
                if req_ist > ack_mail_ist:
                    continue
                req_ack_gap = ack_mail_ist - req_ist
                if req_ack_gap > timedelta(minutes=45):
                    continue
                if req_ack_gap < timedelta(minutes=1):
                    continue

                new_gap = r_ist - ack_mail_ist
                episode_candidates.append((new_gap, -ack_mail_ist.timestamp(), req_mail, ack_mail))

            if not episode_candidates:
                continue

            episode_candidates.sort(key=lambda x: (x[0], x[1]))
            _new_gap, _neg_ack_ts, req_mail, ack_mail = episode_candidates[0]
            req_ist = _to_ist(req_mail.sent_time)
            ack_ist = _to_ist(ack_mail.sent_time)
            if not (req_ist <= ack_ist <= r_ist):
                continue
            # Apply only on clear improvement.
            if (r_ist - ack_ist) >= old_gap - timedelta(hours=2):
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            if new_created == c and new_ack == a:
                continue

            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; EpisodeAckRefreshGuard"

        # Episode ack-refresh guard (strict):
        # Re-anchor Created/Ack only when there is a clearly newer request->ack
        # pair in the same thread and it materially improves Ack->Resolved gap.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue

            old_gap = r_ist - a_ist
            if old_gap < timedelta(hours=12):
                continue

            subj_norm = state.get("subject_norm") or ""
            subj_tokens = _match_tokens(subj_norm)

            requester_candidates = []
            for e in thread:
                if not e.sent_time:
                    continue
                if not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist <= a_ist + timedelta(hours=2) or e_ist > r_ist:
                    continue
                if not _ack_like(e):
                    continue
                key_norm = normalize_subject(e.subject or "")
                key_tokens = _match_tokens(key_norm)
                if subj_tokens and key_tokens:
                    score = _token_overlap_score(subj_tokens, key_tokens)
                    contains = (subj_norm in key_norm or key_norm in subj_norm)
                    if score < 0.45 and not contains:
                        continue
                requester_candidates.append(e)
            if not requester_candidates:
                continue
            requester_candidates.sort(key=lambda e: e.sent_time, reverse=True)

            applied = False
            for ack_mail in requester_candidates:
                ack_ist = _to_ist(ack_mail.sent_time)
                req_before = [
                    e for e in thread
                    if e.sent_time
                    and not _ess_sender(e)
                    and _email_ist(e) and _email_ist(e) <= ack_ist
                ]
                guard_tag = "EpisodeAckRefreshGuard"
                req_mail = None
                req_ist = None
                new_gap = None

                if req_before:
                    req_mail = max(req_before, key=lambda e: e.sent_time)
                    req_ist = _email_ist(req_mail)
                    if not req_ist:
                        continue
                    if req_ist <= c_ist + timedelta(hours=2):
                        continue
                    if not (timedelta(minutes=1) <= (ack_ist - req_ist) <= timedelta(minutes=45)):
                        continue
                    new_gap = r_ist - ack_ist
                    if new_gap >= old_gap - timedelta(hours=4):
                        continue
                else:
                    # Safe fallback for ESS-only rows:
                    # when no non-ESS request exists, allow requester self-episode
                    # re-anchor only on a tight request->ack pair and strong gap improvement.
                    has_non_ess = any(
                        e.sent_time and not _ess_sender(e)
                        and _email_ist(e)
                        for e in thread
                    )
                    if has_non_ess:
                        continue
                    req_before_self = [
                        e for e in thread
                        if e.sent_time
                        and _req_match(e, requester)
                        and _email_ist(e) and _email_ist(e) <= ack_ist
                        and not _ack_like(e)
                    ]
                    if not req_before_self:
                        continue
                    req_mail = max(req_before_self, key=lambda e: e.sent_time)
                    req_ist = _email_ist(req_mail)
                    if not req_ist:
                        continue
                    if req_ist <= c_ist + timedelta(hours=8):
                        continue
                    if not (timedelta(minutes=1) <= (ack_ist - req_ist) <= timedelta(minutes=30)):
                        continue
                    new_gap = r_ist - ack_ist
                    if new_gap >= old_gap - timedelta(hours=8):
                        continue
                    guard_tag = "EpisodeAckRefreshSelfGuard"

                new_created = _format_time(req_mail.sent_time)
                new_ack = _format_time(ack_mail.sent_time)
                if not new_created or not new_ack:
                    continue
                row_vals["Created Date & Time"] = new_created
                row_vals["Actual Response Date & Time"] = new_ack
                row_idx = state.get("row_index")
                if row_idx:
                    ws.cell(row_idx, created_col).value = new_created
                    ws.cell(row_idx, response_col).value = new_ack
                if list_index < len(debug_rows):
                    debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                    debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; {guard_tag}"
                applied = True
                break
            if applied:
                continue

        # ESS-only span rebase guard (strict):
        # For ESS-only span rows with very old Ack, allow rebasing Created/Ack to
        # a clearly later requester episode when no non-ESS request exists.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue
            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            if not any(
                tag in notes_l
                for tag in (
                    "ess-only; no non-ess request; span",
                    "ess-only; no non-ess request; requester span",
                    "ess-only; no non-ess request; requester span(ack-like)",
                    "ess-only; no non-ess request; requester span(all-ack->ess)",
                )
            ):
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue

            old_gap = r_ist - a_ist
            if old_gap < timedelta(hours=12):
                continue

            requester_after_all = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and _email_ist(e) and _email_ist(e) >= a_ist + timedelta(hours=2)
                and _email_ist(e) <= r_ist
            ]
            requester_after_all.sort(key=lambda e: e.sent_time)
            requester_after_non_ack = [e for e in requester_after_all if not _ack_like(e)]
            if not requester_after_non_ack:
                continue

            req_mail = requester_after_non_ack[0]
            req_ist = _to_ist(req_mail.sent_time)
            ack_mail = None
            # Prefer explicit ack-like requester follow-up within 60m.
            for e in requester_after_all:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist < req_ist:
                    continue
                if e_ist > req_ist + timedelta(minutes=60):
                    break
                if _ack_like(e):
                    ack_mail = e
                    break
            if ack_mail is None:
                # Conservative fallback: treat requester episode start as ack only
                # when it still materially improves stale Ack->Resolved gap.
                ack_mail = req_mail

            ack_ist = _to_ist(ack_mail.sent_time)
            if not (req_ist <= ack_ist <= r_ist):
                continue
            new_gap = r_ist - ack_ist
            if new_gap >= old_gap - timedelta(hours=6):
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ESSSpanRebaseGuard"

        # Requester episode rebase guard (very narrow):
        # For stale rows where resolved is already from requester but created/ack
        # stayed on an older non-requester episode, rebase Created/Ack to the
        # first requester non-ack episode in the same thread.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue
            old_gap = r_ist - a_ist
            if old_gap < timedelta(hours=10):
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""

            # Keep this guard narrow to avoid disturbing normal request/ack rows.
            created_is_parsed = isinstance(created_src_now, str) and created_src_now.startswith("PARSED_FROM_")
            span_row = isinstance(notes_now, str) and ("span" in notes_now.lower())
            if not (created_is_parsed or span_row):
                continue
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue
            if _match_requester(ack_src_now, ack_src_now, requester):
                continue

            requester_after = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and not _ack_like(e)
                and _email_ist(e) and _email_ist(e) >= a_ist + timedelta(hours=2)
                and _email_ist(e) <= r_ist
            ]
            requester_after.sort(key=lambda e: e.sent_time)
            if not requester_after:
                continue

            req_mail = requester_after[0]
            req_ist = _to_ist(req_mail.sent_time)
            ack_mail = req_mail
            for e in thread:
                if not e.sent_time:
                    continue
                if not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist < req_ist:
                    continue
                if e_ist > req_ist + timedelta(minutes=45):
                    break
                if _ack_like(e):
                    ack_mail = e
                    break

            ack_ist = _to_ist(ack_mail.sent_time)
            if not (req_ist <= ack_ist <= r_ist):
                continue
            new_gap = r_ist - ack_ist
            if new_gap >= old_gap - timedelta(hours=2):
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterEpisodeRebaseGuard"

        # Final ack-delay guard (narrow):
        # In ESS-only rows, prevent very-late ack timestamps from stale carry-over.
        # Do not run this on mixed/requester rows.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue
            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            if "ess-only; no non-ess request" not in notes_l:
                continue

            has_non_ess = any(
                e.sent_time
                and not _ess_sender(e)
                and not _system_like_sender(e)
                for e in thread
            )
            if has_non_ess:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            if not (c_dt and a_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            if a_ist <= c_ist:
                continue
            if (a_ist - c_ist) <= timedelta(hours=12):
                continue

            window_end = c_ist + timedelta(hours=12)
            requester_window = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and _email_ist(e) and _email_ist(e) > c_ist
                and _email_ist(e) <= window_end
            ]
            requester_window.sort(key=lambda e: e.sent_time)
            ess_window = [
                e for e in thread
                if e.sent_time
                and _ess_sender(e)
                and _email_ist(e) and _email_ist(e) > c_ist
                and _email_ist(e) <= window_end
            ]
            ess_window.sort(key=lambda e: e.sent_time)

            ack_pick = None
            for e in requester_window:
                if _ack_like(e):
                    ack_pick = e
                    break
            if ack_pick is None:
                for e in ess_window:
                    if _ack_like(e):
                        ack_pick = e
                        break
            if ack_pick is None and requester_window:
                ack_pick = requester_window[0]
            if ack_pick is None and ess_window:
                ack_pick = ess_window[0]

            new_ack = _format_time(ack_pick.sent_time) if ack_pick else c
            if not new_ack or new_ack == a:
                continue
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["AckSource"] = (ack_pick.sender_email or ack_pick.sender_name) if ack_pick else "ACK NOT FOUND"
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; AckDelayWindowGuard"

        # Quoted request rebase guard (safe):
        # When Created is stale vs Ack/Resolved, mine requester mails in the same
        # episode for quoted non-ESS request times and re-anchor Created.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            if not requester or not subject_norm:
                continue
            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            if not (c_dt and (a_dt or r_dt)):
                continue
            c_ist = _to_ist(c_dt)
            upper_ist = _to_ist(a_dt) if a_dt else _to_ist(r_dt)
            if not upper_ist:
                continue
            if (upper_ist - c_ist) < timedelta(hours=18):
                continue

            base_thread = state.get("thread") or []
            thread = _expanded_thread(
                subject_norm,
                base_thread,
                requester,
                include_non_ess=True,
                reference_ist=upper_ist,
            )
            if not thread:
                continue

            parsed_candidates = []
            for e in thread:
                if not getattr(e, "sent_time", None):
                    continue
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist > (upper_ist + timedelta(minutes=5)):
                    continue
                if e_ist < (upper_ist - timedelta(days=10)):
                    continue
                if not _req_match(e, requester):
                    continue
                if _ack_like(e) or _ack_like_text_fallback(e):
                    continue
                try:
                    parsed_dt = _extract_request_time_from_email(
                        e,
                        ess_team,
                        max_dt=e.sent_time,
                        subject_norm=subject_norm,
                    )
                except Exception:
                    parsed_dt = None
                if not parsed_dt:
                    continue
                parsed_ist = _to_ist(parsed_dt)
                if parsed_ist > (upper_ist + timedelta(minutes=5)):
                    continue
                if parsed_ist < (upper_ist - timedelta(days=10)):
                    continue
                if parsed_ist <= (c_ist + timedelta(minutes=30)):
                    continue
                parsed_candidates.append(parsed_ist)
            if not parsed_candidates:
                continue

            parsed_candidates.sort()
            pick_ist = parsed_candidates[-1]
            t = _format_time(pick_ist)
            if not t:
                continue

            row_vals["Created Date & Time"] = t
            if a_dt and _to_ist(a_dt) < pick_ist:
                row_vals["Actual Response Date & Time"] = t
            if r_dt and _to_ist(r_dt) < pick_ist:
                row_vals["Actual Resolved Date & Time"] = t

            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = row_vals.get("Created Date & Time")
                ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                ws.cell(row_idx, resolved_col).value = row_vals.get("Actual Resolved Date & Time")
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = "PARSED_FROM_QUOTED_REQUEST"
                if a_dt and _to_ist(a_dt) < pick_ist:
                    debug_rows[list_index]["AckSource"] = "PARSED_FROM_QUOTED_REQUEST"
                if r_dt and _to_ist(r_dt) < pick_ist:
                    debug_rows[list_index]["ResolvedSource"] = "PARSED_FROM_QUOTED_REQUEST"
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; QuotedRequestRebaseGuard"

        # Final non-ack resolved guard (global, safe):
        # If resolved lands on an ack-like requester reminder, rebase resolved
        # to latest requester non-ack reply in the same thread.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not r_dt:
                continue
            floor_dt = None
            if c_dt and a_dt:
                floor_dt = max(_to_ist(c_dt), _to_ist(a_dt))
            elif a_dt:
                floor_dt = _to_ist(a_dt)
            elif c_dt:
                floor_dt = _to_ist(c_dt)
            r_ist = _to_ist(r_dt)

            resolved_emails = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and _email_ist(e) and abs((_email_ist(e) - r_ist).total_seconds()) <= 61
            ]
            resolved_mail = None
            if resolved_emails:
                resolved_mail = min(
                    resolved_emails,
                    key=lambda e: abs((_email_ist(e) - r_ist).total_seconds()),
                )
            else:
                # Sheet timestamps are minute-precision; tolerate small drift.
                near_hits = [
                    e for e in thread
                    if e.sent_time
                    and _req_match(e, requester)
                    and _email_ist(e) and abs((_email_ist(e) - r_ist).total_seconds()) <= 300
                ]
                if near_hits:
                    resolved_mail = min(
                        near_hits,
                        key=lambda e: abs((_email_ist(e) - r_ist).total_seconds()),
                    )
            if not resolved_mail or not _ack_like(resolved_mail):
                continue

            non_ack_pool = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and not _ack_like(e)
                and _email_ist(e) and _email_ist(e) <= r_ist
                and (floor_dt is None or _email_ist(e) >= floor_dt)
            ]
            if not non_ack_pool:
                continue

            pick = max(non_ack_pool, key=lambda e: e.sent_time)

            new_res = _format_time(pick.sent_time)
            if not new_res or new_res == r:
                continue
            row_vals["Actual Resolved Date & Time"] = new_res
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, resolved_col).value = new_res
            if list_index < len(debug_rows):
                debug_rows[list_index]["ResolvedSource"] = pick.sender_email or pick.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ResolvedNonAckGuard"

        # Late-episode rebase guard (narrow):
        # If resolved is much later than ack and resolved belongs to requester,
        # re-anchor Created/Ack to the latest requester episode near resolved.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue
            if (r_ist - a_ist) < timedelta(hours=48):
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue
            if _match_requester(ack_src_now, ack_src_now, requester) and not str(created_src_now).startswith("PARSED_FROM_"):
                continue

            requester_timeline = [
                e for e in thread
                if e.sent_time and _req_match(e, requester) and _email_ist(e) and _email_ist(e) <= r_ist
            ]
            requester_timeline.sort(key=lambda e: e.sent_time)
            if not requester_timeline:
                continue

            # Use latest requester episode ending at resolved.
            end_idx = len(requester_timeline) - 1
            episode_start_idx = end_idx
            for j in range(end_idx, 0, -1):
                curr_t = _email_ist(requester_timeline[j])
                prev_t = _email_ist(requester_timeline[j - 1])
                if not curr_t or not prev_t:
                    continue
                if (curr_t - prev_t) > timedelta(hours=6):
                    break
                episode_start_idx = j - 1

            episode_slice = requester_timeline[episode_start_idx : end_idx + 1]
            non_ack_episode = [e for e in episode_slice if not _ack_like(e)]
            if not non_ack_episode:
                continue
            req_mail = non_ack_episode[0]
            req_ist = _email_ist(req_mail)
            if not req_ist:
                continue
            if req_ist <= a_ist + timedelta(hours=2):
                continue

            ack_mail = req_mail
            for e in episode_slice:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist < req_ist:
                    continue
                if e_ist > req_ist + timedelta(minutes=60):
                    break
                if _ack_like(e):
                    ack_mail = e
                    break
            ack_ist = _email_ist(ack_mail)
            if not ack_ist or not (req_ist <= ack_ist <= r_ist):
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            if new_created == c and new_ack == a:
                continue

            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; LateEpisodeRebaseGuard"

        # Requester-ack ownership guard (narrow):
        # If resolved is from requester but ack is not, align ack to first
        # requester reply after created within the row thread.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            thread = state.get("thread") or []
            requester = state.get("requester") or ""
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue
            if _match_requester(ack_src_now, ack_src_now, requester):
                continue

            requester_after_created = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and _email_ist(e) and _email_ist(e) >= c_ist
                and _email_ist(e) <= r_ist
            ]
            requester_after_created.sort(key=lambda e: e.sent_time)
            if not requester_after_created:
                continue

            ack_pick = requester_after_created[0]
            ack_pick_ist = _email_ist(ack_pick)
            if not ack_pick_ist:
                continue

            new_ack = _format_time(ack_pick.sent_time)
            if not new_ack:
                continue

            new_created = c
            # Only rebase created in stale parsed-anchor cases.
            if (
                isinstance(created_src_now, str)
                and created_src_now.startswith("PARSED_FROM_")
                and (ack_pick_ist - c_ist) >= timedelta(hours=24)
            ):
                new_created = new_ack

            changed = (new_ack != a) or (new_created != c)
            if not changed:
                continue

            row_vals["Actual Response Date & Time"] = new_ack
            row_vals["Created Date & Time"] = new_created
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                if new_created != c:
                    debug_rows[list_index]["CreatedSource"] = ack_pick.sender_email or ack_pick.sender_name
                debug_rows[list_index]["AckSource"] = ack_pick.sender_email or ack_pick.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterAckOwnershipGuard"

        # Generic stale-created guard for ESS-only rows:
        # if Created stayed on an older non-requester episode while Ack/Resolved
        # are requester-owned, rebase Created to the requester episode near Ack.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester or not subject_norm:
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            if "ESS-only; no non-ESS request" not in notes_now:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue
            if (a_ist - c_ist) < timedelta(hours=12):
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            if _match_requester(created_src_now, created_src_now, requester):
                continue
            if not _match_requester(ack_src_now, ack_src_now, requester):
                continue
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue

            row_tokens = _match_tokens(subject_norm)
            if not row_tokens:
                continue

            requester_before_ack = []
            for e in thread:
                if not e.sent_time or not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist or e_ist > a_ist:
                    continue
                # Keep this near the active episode only.
                if e_ist < (a_ist - timedelta(days=2)):
                    continue
                mail_subj = normalize_subject(e.subject or "")
                mail_tokens = _match_tokens(mail_subj)
                sim = _token_overlap_score(row_tokens, mail_tokens) if mail_tokens else 0.0
                contains = (subject_norm in mail_subj or mail_subj in subject_norm) if mail_subj else False
                if sim >= 0.40 or contains:
                    requester_before_ack.append(e)

            requester_before_ack.sort(key=lambda e: e.sent_time)
            if not requester_before_ack:
                continue

            # Prefer a non-ack requester message as Created anchor.
            non_ack_before_ack = [e for e in requester_before_ack if not _ack_like(e)]
            if not non_ack_before_ack:
                continue
            req_mail = non_ack_before_ack[0]

            req_ist = _email_ist(req_mail)
            if not req_ist or req_ist > a_ist:
                continue

            new_created = _format_time(req_mail.sent_time)
            if not new_created or new_created == c:
                continue
            row_vals["Created Date & Time"] = new_created
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; StaleCreatedRequesterEpisodeGuard"

        # Generic requester subject-episode ownership guard:
        # when resolved belongs to requester but ack source does not, re-anchor
        # Created/Ack to the first requester episode that matches this row subject.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester or not subject_norm:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue
            if (r_ist - a_ist) < timedelta(hours=8):
                continue

            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue
            if _match_requester(ack_src_now, ack_src_now, requester):
                continue

            row_tokens = _match_tokens(subject_norm)
            if not row_tokens:
                continue

            window_end = min(r_ist + timedelta(minutes=5), c_ist + timedelta(days=10))
            requester_timeline = []
            for e in thread:
                if not e.sent_time or not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist or e_ist < c_ist or e_ist > window_end:
                    continue
                mail_subj = normalize_subject(e.subject or "")
                mail_tokens = _match_tokens(mail_subj)
                sim = _token_overlap_score(row_tokens, mail_tokens) if mail_tokens else 0.0
                contains = (subject_norm in mail_subj or mail_subj in subject_norm) if mail_subj else False
                if sim >= 0.35 or contains:
                    requester_timeline.append(e)

            requester_timeline.sort(key=lambda e: e.sent_time)
            if not requester_timeline:
                continue

            requester_non_ack = [e for e in requester_timeline if not _ack_like(e)]
            if not requester_non_ack:
                continue
            req_mail = requester_non_ack[0]
            req_ist = _email_ist(req_mail)
            if not req_ist:
                continue
            if req_ist <= a_ist + timedelta(hours=2):
                continue
            if (r_ist - req_ist) > timedelta(days=2):
                continue

            ack_mail = req_mail
            for e in requester_timeline:
                e_ist = _email_ist(e)
                if not e_ist or e_ist < req_ist:
                    continue
                if e_ist > req_ist + timedelta(minutes=60):
                    break
                if _ack_like(e):
                    ack_mail = e
                    break

            ack_ist = _email_ist(ack_mail)
            if not ack_ist or ack_ist > (r_ist + timedelta(minutes=5)):
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            if new_created == c and new_ack == a:
                continue

            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterSubjectEpisodeGuard"

        # Mixed requester-ownership guard (global):
        # If Created/Ack were taken from the same non-requester source but Resolved
        # belongs to requester, re-anchor Created/Ack to requester episode near Resolved.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester or not subject_norm:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue
            if (r_ist - a_ist) < timedelta(hours=4):
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue
            if _match_requester(ack_src_now, ack_src_now, requester):
                continue
            if not created_src_now or not ack_src_now:
                continue
            if _requester_key(created_src_now) != _requester_key(ack_src_now):
                continue

            row_tokens = _match_tokens(subject_norm)
            if not row_tokens:
                continue

            r_soft = r_ist + timedelta(minutes=5)
            requester_timeline = []
            for e in thread:
                if not e.sent_time or not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist or e_ist > r_soft:
                    continue
                mail_subj = normalize_subject(e.subject or "")
                mail_tokens = _match_tokens(mail_subj)
                sim = _token_overlap_score(row_tokens, mail_tokens) if mail_tokens else 0.0
                contains = (subject_norm in mail_subj or mail_subj in subject_norm) if mail_subj else False
                if sim >= 0.30 or contains:
                    requester_timeline.append(e)

            requester_timeline.sort(key=lambda e: e.sent_time)
            if not requester_timeline:
                continue

            end_idx = len(requester_timeline) - 1
            start_idx = end_idx
            for j in range(end_idx, 0, -1):
                curr_t = _email_ist(requester_timeline[j])
                prev_t = _email_ist(requester_timeline[j - 1])
                if not curr_t or not prev_t:
                    continue
                if (curr_t - prev_t) > timedelta(hours=8):
                    break
                start_idx = j - 1
            episode = requester_timeline[start_idx : end_idx + 1]
            if not episode:
                continue

            non_ack_episode = [e for e in episode if not _ack_like(e)]
            if not non_ack_episode:
                continue
            req_mail = non_ack_episode[0]
            req_ist = _email_ist(req_mail)
            if not req_ist:
                continue
            if req_ist <= c_ist + timedelta(minutes=30):
                continue
            if req_ist < (a_ist - timedelta(minutes=5)):
                continue
            if (r_ist - req_ist) > timedelta(days=2):
                continue

            ack_mail = req_mail
            for e in episode:
                e_ist = _email_ist(e)
                if not e_ist or e_ist < req_ist:
                    continue
                if e_ist > req_ist + timedelta(minutes=60):
                    break
                if _ack_like(e):
                    ack_mail = e
                    break
            ack_ist = _email_ist(ack_mail)
            if not ack_ist or ack_ist > r_soft:
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            if new_created == c and new_ack == a:
                continue

            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; MixedRequesterEpisodeGuard"

        # Resolved-window reanchor guard (global):
        # If Created/Ack are from same non-requester source but Resolved is requester,
        # anchor Created/Ack from requester activity close to Resolved timestamp.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester or not subject_norm:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            if not created_src_now or not ack_src_now:
                continue
            if _requester_key(created_src_now) != _requester_key(ack_src_now):
                continue
            if _match_requester(ack_src_now, ack_src_now, requester):
                continue
            if not _match_requester(resolved_src_now, resolved_src_now, requester):
                continue

            row_tokens = _match_tokens(subject_norm)
            if not row_tokens:
                continue

            win_start = r_ist - timedelta(hours=48)
            win_end = r_ist + timedelta(minutes=5)
            requester_window = []
            for e in thread:
                if not e.sent_time or not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist or e_ist < win_start or e_ist > win_end:
                    continue
                mail_subj = normalize_subject(e.subject or "")
                mail_tokens = _match_tokens(mail_subj)
                sim = _token_overlap_score(row_tokens, mail_tokens) if mail_tokens else 0.0
                contains = (subject_norm in mail_subj or mail_subj in subject_norm) if mail_subj else False
                if sim >= 0.25 or contains:
                    requester_window.append(e)

            requester_window.sort(key=lambda e: e.sent_time)
            if not requester_window:
                continue

            requester_non_ack = [e for e in requester_window if not _ack_like(e)]
            if not requester_non_ack:
                continue
            req_mail = requester_non_ack[0]
            req_ist = _email_ist(req_mail)
            if not req_ist:
                continue
            if req_ist < (c_ist - timedelta(minutes=5)):
                continue

            ack_mail = req_mail
            for e in requester_window:
                e_ist = _email_ist(e)
                if not e_ist or e_ist < req_ist:
                    continue
                if e_ist > req_ist + timedelta(minutes=60):
                    break
                if _ack_like(e):
                    ack_mail = e
                    break
            ack_ist = _email_ist(ack_mail)
            if not ack_ist or ack_ist > win_end:
                continue

            new_created = _format_time(req_mail.sent_time)
            new_ack = _format_time(ack_mail.sent_time)
            if not new_created or not new_ack:
                continue
            if new_created == c and new_ack == a:
                continue

            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = req_mail.sender_email or req_mail.sender_name
                debug_rows[list_index]["AckSource"] = ack_mail.sender_email or ack_mail.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ResolvedWindowReanchorGuard"

        # Requester-span ack-like guard (global):
        # For ACK NOT FOUND + requester span(ack-like), prefer latest requester
        # non-ack in thread so reminder mails do not anchor old timestamps.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue
            if list_index < len(debug_rows):
                notes_now = (debug_rows[list_index].get("Notes", "") or "").lower()
                if "requester follow-up (no in-between request)" in notes_now:
                    continue
            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester:
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            if (
                "requester span(ack-like)" not in notes_l
                and "requester span(all-ack->ess)" not in notes_l
            ):
                continue
            if str(ack_src_now).strip().upper() != "ACK NOT FOUND":
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            r_dt = _parse_time_str(r)
            if not (c_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            r_ist = _to_ist(r_dt)

            requester_msgs = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and _email_ist(e)
            ]
            requester_msgs.sort(key=lambda e: e.sent_time)
            if not requester_msgs:
                continue

            episodes = []
            current = [requester_msgs[0]]
            for e in requester_msgs[1:]:
                prev_t = _email_ist(current[-1])
                cur_t = _email_ist(e)
                if not prev_t or not cur_t:
                    continue
                if (cur_t - prev_t) > timedelta(hours=8):
                    episodes.append(current)
                    current = [e]
                else:
                    current.append(e)
            episodes.append(current)
            if not episodes:
                continue

            latest_episode = episodes[-1]
            latest_non_ack = [e for e in latest_episode if not _ack_like(e)]
            if not latest_non_ack:
                continue
            latest_req = latest_non_ack[-1]
            latest_ist = _email_ist(latest_req)
            if not latest_ist:
                continue
            if latest_ist <= c_ist:
                continue
            if latest_ist < (r_ist - timedelta(days=2)):
                continue

            baseline_date = state.get("baseline_created_date")
            if baseline_date:
                if abs((latest_ist.date() - baseline_date).days) > 2:
                    continue
            else:
                # Keep conservative bound when baseline isn't available.
                if (latest_ist - c_ist) > timedelta(days=7):
                    continue

            t = _format_time(latest_req.sent_time)
            if not t:
                continue
            if latest_ist <= r_ist:
                continue
            row_vals["Actual Resolved Date & Time"] = t
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, resolved_col).value = t
            if list_index < len(debug_rows):
                who = latest_req.sender_email or latest_req.sender_name
                debug_rows[list_index]["ResolvedSource"] = who
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterSpanAckLikeGuard; ResolvedOnly"

        # Episode consistency guard (global):
        # For mixed-source/problem rows, pick one requester episode and derive all
        # three timestamps from that episode to avoid cross-episode mixing.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester or not subject_norm:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if not (c_ist <= a_ist <= r_ist):
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""

            mixed_sources = (
                created_src_now
                and ack_src_now
                and (_requester_key(created_src_now) == _requester_key(ack_src_now))
                and (not _match_requester(ack_src_now, ack_src_now, requester))
                and _match_requester(resolved_src_now, resolved_src_now, requester)
            )
            span_ack_missing = (
                "requester span" in (notes_now or "").lower()
                and str(ack_src_now).strip().upper() == "ACK NOT FOUND"
            )
            if not (mixed_sources or span_ack_missing):
                continue

            row_tokens = _match_tokens(subject_norm)
            if not row_tokens:
                continue

            requester_msgs = []
            for e in thread:
                if not e.sent_time or not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                mail_subj = normalize_subject(e.subject or "")
                mail_tokens = _match_tokens(mail_subj)
                sim = _token_overlap_score(row_tokens, mail_tokens) if mail_tokens else 0.0
                contains = (subject_norm in mail_subj or mail_subj in subject_norm) if mail_subj else False
                if sim >= 0.20 or contains:
                    requester_msgs.append(e)

            requester_msgs.sort(key=lambda e: e.sent_time)
            if not requester_msgs:
                continue

            episodes = []
            current = [requester_msgs[0]]
            for e in requester_msgs[1:]:
                prev_t = _email_ist(current[-1])
                cur_t = _email_ist(e)
                if not prev_t or not cur_t:
                    continue
                if (cur_t - prev_t) > timedelta(hours=8):
                    episodes.append(current)
                    current = [e]
                else:
                    current.append(e)
            episodes.append(current)
            if not episodes:
                continue

            baseline_date = state.get("baseline_created_date")
            chosen = None
            chosen_mode = "fallback"

            # For requester-span(ack-like) rows, always take the latest requester episode.
            # This prevents stale old episodes from winning via baseline.
            if span_ack_missing:
                chosen = episodes[-1]
                chosen_mode = "latest_span"

            # For mixed-source rows, prefer the episode nearest to resolved timestamp.
            if chosen is None and mixed_sources:
                near_ranked = []
                for ep in episodes:
                    ep_times = [(_email_ist(m), m) for m in ep if _email_ist(m)]
                    if not ep_times:
                        continue
                    min_abs = min(abs((t - r_ist).total_seconds()) for t, _ in ep_times)
                    before_flag = 0 if any(t <= r_ist for t, _ in ep_times) else 1
                    near_ranked.append((before_flag, min_abs, -_email_ist(ep[-1]).timestamp(), ep))
                if near_ranked:
                    near_ranked.sort()
                    chosen = near_ranked[0][3]
                    chosen_mode = "near_resolved"

            # Baseline-aware fallback for non-span/non-mixed rows.
            if chosen is None and baseline_date:
                def _episode_anchor(ep):
                    non_ack = [m for m in ep if not _ack_like(m)]
                    return non_ack[0] if non_ack else None
                ranked = []
                for ep in episodes:
                    anchor = _episode_anchor(ep)
                    if not anchor:
                        continue
                    anchor_t = _email_ist(anchor)
                    if not anchor_t:
                        continue
                    day_delta = abs((anchor_t.date() - baseline_date).days)
                    has_r = any(abs((_email_ist(m) - r_ist).total_seconds()) <= 21600 for m in ep if _email_ist(m))
                    ranked.append((day_delta, 0 if has_r else 1, -_email_ist(ep[-1]).timestamp(), ep))
                if ranked:
                    ranked.sort()
                    chosen = ranked[0][3]
                    chosen_mode = "baseline"

            if chosen is None:
                by_resolved = [
                    ep for ep in episodes
                    if any(abs((_email_ist(m) - r_ist).total_seconds()) <= 21600 for m in ep if _email_ist(m))
                ]
                chosen = by_resolved[-1] if by_resolved else episodes[-1]
                chosen_mode = "resolved_or_latest"

            if not chosen:
                continue

            chosen_non_ack = [m for m in chosen if not _ack_like(m)]
            if chosen_non_ack:
                ep_created = chosen_non_ack[0]
                ep_created_ist = _email_ist(ep_created)
                if not ep_created_ist:
                    continue

                ep_ack = ep_created
                for m in chosen:
                    m_ist = _email_ist(m)
                    if not m_ist or m_ist < ep_created_ist:
                        continue
                    if m_ist > ep_created_ist + timedelta(minutes=60):
                        break
                    if _ack_like(m):
                        ep_ack = m
                        break
                ep_ack_ist = _email_ist(ep_ack)
                if not ep_ack_ist:
                    continue

                non_ack_after_ack = [m for m in chosen if _email_ist(m) and _email_ist(m) >= ep_ack_ist and not _ack_like(m)]
                if not non_ack_after_ack:
                    continue
                ep_resolved = non_ack_after_ack[-1]
                ep_resolved_ist = _email_ist(ep_resolved)
                if not ep_resolved_ist:
                    continue

                if not (ep_created_ist <= ep_ack_ist <= ep_resolved_ist):
                    continue
            else:
                # All requester mails in this episode are ack-like. Keep this isolated
                # to requester-span rows with missing ACK source; collapse to latest
                # requester mail in the latest episode to avoid stale cross-episode picks.
                if not span_ack_missing:
                    continue
                ep_resolved = chosen[-1]
                ep_resolved_ist = _email_ist(ep_resolved)
                if not ep_resolved_ist:
                    continue
                ep_created = ep_resolved
                ep_ack = ep_resolved
                ep_created_ist = ep_resolved_ist
                ep_ack_ist = ep_resolved_ist
                chosen_mode = "latest_span_all_ack"

            # Safe bound: do not jump to a very distant day unexpectedly.
            # Do not apply this cap for latest-span rows.
            if baseline_date and chosen_mode not in ("latest_span", "latest_span_all_ack"):
                old_delta = abs((c_ist.date() - baseline_date).days)
                new_delta = abs((ep_created_ist.date() - baseline_date).days)
                if new_delta > 5 and new_delta > old_delta:
                    continue

            new_created = _format_time(ep_created.sent_time)
            new_ack = _format_time(ep_ack.sent_time)
            new_resolved = _format_time(ep_resolved.sent_time)
            if not new_created or not new_ack or not new_resolved:
                continue
            if new_created == c and new_ack == a and new_resolved == r:
                continue

            row_vals["Created Date & Time"] = new_created
            row_vals["Actual Response Date & Time"] = new_ack
            row_vals["Actual Resolved Date & Time"] = new_resolved
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = new_created
                ws.cell(row_idx, response_col).value = new_ack
                ws.cell(row_idx, resolved_col).value = new_resolved
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = ep_created.sender_email or ep_created.sender_name
                debug_rows[list_index]["AckSource"] = ep_ack.sender_email or ep_ack.sender_name
                debug_rows[list_index]["ResolvedSource"] = ep_resolved.sender_email or ep_resolved.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; EpisodeConsistencyGuard[{chosen_mode}]"

        # Final requester-span(ack-like) fallback (global):
        # If a row still sits on an old timestamp with ACK NOT FOUND, move it to
        # latest requester mail in the same expanded thread.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester:
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            # Keep this guard generic but isolated to ESS-only span-like rows.
            if "ess-only; no non-ess request" not in notes_l:
                continue

            row_vals = automation_rows[list_index]
            c = row_vals.get("Created Date & Time")
            c_dt = _parse_time_str(c)
            c_ist = _to_ist(c_dt) if c_dt else None
            if not c_ist:
                continue
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            r_ist = _to_ist(r_dt) if r_dt else None

            row_tokens = _match_tokens(subject_norm)
            requester_msgs = [
                e for e in thread
                if e.sent_time
                and _req_match(e, requester)
                and _email_ist(e)
                and (
                    not row_tokens
                    or (
                        (lambda s_norm, s_tokens: (
                            (_token_overlap_score(row_tokens, s_tokens) >= 0.45)
                            or (subject_norm and (subject_norm in s_norm or s_norm in subject_norm))
                        ))(
                            normalize_subject(e.subject or ""),
                            _match_tokens(normalize_subject(e.subject or "")),
                        )
                    )
                )
            ]
            if not requester_msgs:
                continue
            requester_non_ack = [e for e in requester_msgs if not _ack_like(e)]
            requester_non_ack_exists = bool(requester_non_ack)
            used_ess_non_ack = False
            latest_req = max(requester_non_ack, key=lambda e: e.sent_time) if requester_non_ack else max(requester_msgs, key=lambda e: e.sent_time)
            latest_ist = _email_ist(latest_req)
            if not latest_ist:
                continue
            # Only move when there is a clear stale gap from either created/resolved.
            ref_ist = c_ist if not r_ist else max(c_ist, r_ist)
            if (latest_ist - ref_ist) < timedelta(hours=6):
                continue

            t = _format_time(latest_req.sent_time)
            if not t:
                continue
            if requester_non_ack_exists:
                if r_ist and latest_ist <= r_ist:
                    continue
                row_vals["Actual Resolved Date & Time"] = t
            else:
                # All requester mails are ack-like in this span row. Keep timestamps
                # consistent without promoting update-chasing requester mails.
                # Prefer latest ESS non-ack reply before that requester tail.
                ess_non_ack = []
                for e in thread:
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if not _ess_sender(e):
                        continue
                    if _ack_like(e):
                        continue
                    if e_ist < c_ist:
                        continue
                    if e_ist > latest_ist:
                        continue
                    if row_tokens:
                        s_norm = normalize_subject(e.subject or "")
                        s_tokens = _match_tokens(s_norm)
                        if s_tokens:
                            score = _token_overlap_score(row_tokens, s_tokens)
                            contains = subject_norm and (subject_norm in s_norm or s_norm in subject_norm)
                            if score < 0.45 and not contains:
                                continue
                    ess_non_ack.append(e)
                ess_non_ack.sort(key=lambda e: e.sent_time)

                if ess_non_ack:
                    best = ess_non_ack[-1]
                    best_t = _format_time(best.sent_time)
                    if not best_t:
                        continue
                    if r_ist and _email_ist(best) and _email_ist(best) <= r_ist:
                        continue
                    row_vals["Actual Resolved Date & Time"] = best_t
                    t = best_t
                    latest_req = best
                    used_ess_non_ack = True
                else:
                    a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
                    a_ist = _to_ist(a_dt) if a_dt else None
                    if c_ist == latest_ist and a_ist == latest_ist and r_ist == latest_ist:
                        continue
                    row_vals["Created Date & Time"] = t
                    row_vals["Actual Response Date & Time"] = t
                    row_vals["Actual Resolved Date & Time"] = t
            row_idx = state.get("row_index")
            if row_idx:
                if requester_non_ack_exists or used_ess_non_ack:
                    ws.cell(row_idx, resolved_col).value = t
                else:
                    ws.cell(row_idx, created_col).value = t
                    ws.cell(row_idx, response_col).value = t
                    ws.cell(row_idx, resolved_col).value = t
            if list_index < len(debug_rows):
                who = latest_req.sender_email or latest_req.sender_name
                if used_ess_non_ack:
                    debug_rows[list_index]["ResolvedSource"] = who
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterSpanAllAckGuard; ResolvedFromEssNonAck"
                elif requester_non_ack_exists:
                    debug_rows[list_index]["ResolvedSource"] = who
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterSpanAckLikeFinalGuard; ResolvedOnly"
                else:
                    debug_rows[list_index]["CreatedSource"] = who
                    debug_rows[list_index]["AckSource"] = who
                    debug_rows[list_index]["ResolvedSource"] = who
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RequesterSpanAllAckGuard"

        # Final resolved non-ack guard (global):
        # If resolved currently points to an ack-like/update-like requester reply
        # (e.g., "thank you for the update/information"), move resolved to the latest
        # prior non-ack requester reply in the same subject episode.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue
            if list_index < len(debug_rows):
                notes_now = (debug_rows[list_index].get("Notes", "") or "").lower()
                if "requester follow-up (no in-between request)" in notes_now:
                    continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester:
                continue

            row_vals = automation_rows[list_index]
            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            if not (c_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt) if a_dt else None
            r_ist = _to_ist(r_dt)

            row_tokens = _match_tokens(subject_norm)
            requester_msgs = []
            for e in thread:
                if not e.sent_time or not _req_match(e, requester):
                    continue
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if row_tokens:
                    s_norm = normalize_subject(e.subject or "")
                    s_tokens = _match_tokens(s_norm)
                    if s_tokens:
                        score = _token_overlap_score(row_tokens, s_tokens)
                        contains = subject_norm and (subject_norm in s_norm or s_norm in subject_norm)
                        if score < 0.45 and not contains:
                            continue
                requester_msgs.append(e)

            if not requester_msgs:
                continue

            # Find the message that likely produced current resolved timestamp.
            resolved_hits = []
            for e in requester_msgs:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if abs((e_ist - r_ist).total_seconds()) <= 180:
                    resolved_hits.append(e)
            if resolved_hits:
                resolved_mail = max(resolved_hits, key=lambda e: e.sent_time)
            else:
                # Minute-level sheet timestamps can miss seconds; use nearest requester
                # mail at/before resolved (soft +5m tolerance).
                prior_msgs = [
                    e for e in requester_msgs
                    if _email_ist(e) and _email_ist(e) <= (r_ist + timedelta(minutes=5))
                ]
                if not prior_msgs:
                    continue
                resolved_mail = max(prior_msgs, key=lambda e: e.sent_time)
            if not (_ack_like(resolved_mail) or _ack_like_text_fallback(resolved_mail)):
                continue

            fallback = [
                e for e in requester_msgs
                if not _ack_like(e)
                and _email_ist(e)
                and _email_ist(e) <= (r_ist + timedelta(minutes=5))
            ]
            pick = max(fallback, key=lambda e: e.sent_time) if fallback else None
            pick_ist = _email_ist(pick) if pick else None
            pick_src = (pick.sender_email or pick.sender_name) if pick else ""
            used_quoted = False
            if not pick_ist:
                quoted_ist = _extract_quoted_requester_reply_ist(
                    resolved_mail,
                    requester,
                    subject_norm,
                    c_ist,
                    r_ist,
                )
                if not quoted_ist:
                    continue
                pick_ist = quoted_ist
                pick_src = "PARSED_FROM_QUOTED_REPLY"
                used_quoted = True
            # Safety: do not jump too far back in time on ambiguous long chains.
            if (r_ist - pick_ist) > timedelta(days=14):
                continue

            t = _format_time(pick_ist)
            if not t:
                continue
            reanchored = False
            if a_ist and pick_ist < a_ist:
                row_vals["Actual Response Date & Time"] = t
                if c_ist > pick_ist:
                    row_vals["Created Date & Time"] = t
                reanchored = True
            row_vals["Actual Resolved Date & Time"] = t
            row_idx = state.get("row_index")
            if row_idx:
                if reanchored:
                    ws.cell(row_idx, created_col).value = row_vals.get("Created Date & Time")
                    ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                ws.cell(row_idx, resolved_col).value = t
            if list_index < len(debug_rows):
                if reanchored:
                    debug_rows[list_index]["CreatedSource"] = pick_src
                    debug_rows[list_index]["AckSource"] = pick_src
                debug_rows[list_index]["ResolvedSource"] = pick_src
                extra = "; ResolvedAckLikeQuotedGuard" if used_quoted else ""
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ResolvedAckLikeGuard{extra}{'; ResolvedAckLikeReanchor' if reanchored else ''}"

        # Quoted request-anchor guard (isolated):
        # If created was retained due unreliable response anchor and response exists,
        # allow a strictly-bounded quoted request timestamp from the same chain to
        # re-anchor Created only.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue
            if list_index < len(debug_rows):
                notes_now = (debug_rows[list_index].get("Notes", "") or "").lower()
                if "requester follow-up (no in-between request)" in notes_now:
                    continue

            base_thread = state.get("thread") or []
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            thread = _expanded_thread(subject_norm, base_thread, requester)
            if not thread or not requester:
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            row_vals = automation_rows[list_index]
            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            if not (c_dt and a_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            if a_ist <= c_ist:
                continue

            source_mismatch = (
                bool(created_src_now)
                and bool(requester)
                and not _match_requester(created_src_now, created_src_now, requester)
            )
            ack_from_requester = bool(requester) and bool(ack_src_now) and _match_requester(ack_src_now, ack_src_now, requester)
            span_missing_non_ess = "ess-only; no non-ess request" in notes_l
            large_created_gap = (a_ist - c_ist) >= timedelta(hours=8)
            needs_quoted_anchor = (
                "created retained (response anchor unreliable)" in notes_l
                or "created_clamped_to_first" in (created_src_now.lower() if isinstance(created_src_now, str) else "")
                or (span_missing_non_ess and source_mismatch and large_created_gap)
                or (source_mismatch and ack_from_requester and large_created_gap)
            )
            if (
                not needs_quoted_anchor
            ):
                continue

            quoted_candidates = []
            for e in thread:
                if not e.sent_time:
                    continue
                if not _req_match(e, requester):
                    continue
                q_ist = _extract_quoted_request_before_ist(e, subject_norm, a_ist)
                if not q_ist:
                    continue
                quoted_candidates.append(q_ist)

            if not quoted_candidates:
                continue

            q_pick = max(quoted_candidates)
            # Safety bounds: keep within same practical episode.
            if q_pick <= c_ist:
                continue
            if q_pick >= a_ist:
                continue
            if (a_ist - q_pick) > timedelta(hours=24):
                continue

            t = _format_time(q_pick)
            if not t:
                continue
            row_vals["Created Date & Time"] = t
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = t
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = "PARSED_FROM_QUOTED_REQUEST"
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; QuotedRequestAnchorGuard"

        if True:
            # Live request anchor guard (global):
            # If a row drifted into ESS-only/parsed anchoring, but a live non-ESS
            # requester mail exists in the same subject episode before ack/resolved,
            # re-anchor Created to that live request.
            for state in row_states:
                list_index = state.get("list_index")
                if list_index is None or list_index >= len(automation_rows):
                    continue
                if state.get("is_dep_req") or state.get("is_dep_succ"):
                    continue

                row_vals = automation_rows[list_index]
                requester = state.get("requester") or ""
                subject_norm = (state.get("subject_norm") or "").lower()
                if not subject_norm:
                    continue

                notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
                notes_l = (notes_now or "").lower()
                created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
                ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
                resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
                baseline_date = state.get("baseline_created_date")
                baseline_mid = None
                if baseline_date:
                    try:
                        baseline_mid = _to_ist(datetime(baseline_date.year, baseline_date.month, baseline_date.day))
                    except Exception:
                        baseline_mid = None
                no_requester_recovery = (
                    "norequesterthreadrecovery" in notes_l
                    or "no ess or requester replies" in notes_l
                )
                parsed_created = isinstance(created_src_now, str) and created_src_now.startswith("PARSED_FROM_")
                trigger = (
                    parsed_created
                    or "dateanchorafter" in notes_l
                    or "dateanchormissing" in notes_l
                    or no_requester_recovery
                    or "ambiguousresolvedbyrequester" in notes_l
                )
                if not trigger:
                    continue

                c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
                a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
                r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
                c_ist = _to_ist(c_dt) if c_dt else None
                a_ist = _to_ist(a_dt) if a_dt else None
                r_ist = _to_ist(r_dt) if r_dt else None
                baseline_drift = 0
                if baseline_date and c_ist:
                    try:
                        baseline_drift = abs((c_ist.date() - baseline_date).days)
                    except Exception:
                        baseline_drift = 0
                mixed_owner_drift = False
                if requester and c_ist and a_ist and (a_ist - c_ist) >= timedelta(hours=2):
                    try:
                        mixed_owner_drift = (
                            bool(ack_src_now)
                            and bool(resolved_src_now)
                            and (not _match_requester(str(ack_src_now), str(ack_src_now), requester))
                            and _match_requester(str(resolved_src_now), str(resolved_src_now), requester)
                        )
                    except Exception:
                        mixed_owner_drift = False
                # Keep mixed-owner recovery narrow: it should only run when drift is
                # clearly large, otherwise it can pull from a wrong late episode.
                if mixed_owner_drift and (baseline_drift < 3) and (not parsed_created):
                    mixed_owner_drift = False
                # Keep this guard isolated: do not rewrite stable rows.
                if (
                    not parsed_created
                    and not no_requester_recovery
                    and not mixed_owner_drift
                    and baseline_drift <= 2
                ):
                    continue
                # Recovery mode keeps existing early-window behavior for the original
                # no-requester case only. Mixed-owner has its own tighter path below.
                recovery_mode = no_requester_recovery
                upper_ist = a_ist or r_ist
                if not c_ist:
                    continue
                if not upper_ist:
                    upper_ist = c_ist + timedelta(days=7)
                if recovery_mode:
                    upper_ist = max(upper_ist, c_ist + timedelta(days=7))
                if (upper_ist - c_ist) < timedelta(hours=18) and not recovery_mode:
                    continue

                base_thread = state.get("thread") or []
                thread = _expanded_thread(
                    subject_norm,
                    base_thread,
                    requester,
                    include_non_ess=True,
                    reference_ist=upper_ist,
                )
                row_tokens = _match_tokens(subject_norm)
                candidates = []
                for e in (thread or []):
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if e_ist > (upper_ist + timedelta(minutes=5)):
                        continue
                    if e_ist < (upper_ist - timedelta(days=14)):
                        continue
                    if baseline_date and abs((e_ist.date() - baseline_date).days) > 5:
                        continue
                    if e_ist < (c_ist - timedelta(hours=2)):
                        continue
                    if _ess_sender(e):
                        continue
                    if _system_like_sender(e):
                        continue
                    if row_tokens:
                        s_norm = normalize_subject(e.subject or "")
                        s_tokens = _match_tokens(s_norm)
                        score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                        contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                        if score < 0.45 and not contains:
                            continue
                    candidates.append(e)

                pick = None
                pick_ist = None
                pick_src = ""
                allow_global_fallback = not mixed_owner_drift
                def _baseline_rank_mail(e):
                    e_ist = _email_ist(e)
                    if not e_ist or not baseline_date:
                        return (9999, 999999999)
                    day_delta = abs((e_ist.date() - baseline_date).days)
                    prox = abs((e_ist - (baseline_mid or e_ist)).total_seconds())
                    return (day_delta, prox)

                def _baseline_rank_dt(dt_ist):
                    if not dt_ist or not baseline_date:
                        return (9999, 999999999)
                    day_delta = abs((dt_ist.date() - baseline_date).days)
                    prox = abs((dt_ist - (baseline_mid or dt_ist)).total_seconds())
                    return (day_delta, prox)

                if candidates:
                    candidates.sort(key=lambda e: e.sent_time)
                    if baseline_date:
                        pick = min(candidates, key=lambda e: (_baseline_rank_mail(e), -_email_ist(e).timestamp()))
                    elif mixed_owner_drift:
                        local_end = c_ist + timedelta(hours=72)
                        local_candidates = [
                            e for e in candidates
                            if _email_ist(e) and _email_ist(e) >= (c_ist - timedelta(minutes=15)) and _email_ist(e) <= local_end
                        ]
                        if local_candidates:
                            pick = min(local_candidates, key=lambda e: abs((_email_ist(e) - c_ist).total_seconds()))
                    elif recovery_mode:
                        early_window_end = c_ist + timedelta(hours=72)
                        early_candidates = [
                            e for e in candidates
                            if _email_ist(e) and _email_ist(e) >= c_ist and _email_ist(e) <= early_window_end
                        ]
                        pick = early_candidates[0] if early_candidates else candidates[-1]
                    else:
                        pick = candidates[-1]
                    pick_ist = _email_ist(pick)
                    pick_src = pick.sender_email or pick.sender_name

                # Fallback-1: use requester pool across all emails for this subject episode.
                if (not pick_ist) and allow_global_fallback:
                    requester_pool = _requester_pool(subject_norm, requester, upper_ist, day_window=21)
                    fallback_pool = []
                    for e in requester_pool:
                        e_ist = _email_ist(e)
                        if not e_ist:
                            continue
                        if e_ist > (upper_ist + timedelta(minutes=5)):
                            continue
                        if e_ist < (upper_ist - timedelta(days=14)):
                            continue
                        if baseline_date and abs((e_ist.date() - baseline_date).days) > 5:
                            continue
                        if _ack_like(e) or _ack_like_text_fallback(e):
                            continue
                        fallback_pool.append(e)
                    if fallback_pool:
                        if baseline_date:
                            pick = min(fallback_pool, key=lambda e: (_baseline_rank_mail(e), -_email_ist(e).timestamp()))
                        else:
                            pick = max(fallback_pool, key=lambda e: e.sent_time)
                        pick_ist = _email_ist(pick)
                        pick_src = pick.sender_email or pick.sender_name

                # Fallback-2: parse quoted non-ESS request from requester emails.
                if (not pick_ist) and allow_global_fallback:
                    parsed_candidates = []
                    for e in (thread or []):
                        if not getattr(e, "sent_time", None):
                            continue
                        e_ist = _email_ist(e)
                        if not e_ist:
                            continue
                        if e_ist > (upper_ist + timedelta(minutes=5)):
                            continue
                        if not _req_match(e, requester):
                            continue
                        try:
                            parsed_dt = _extract_request_time_from_email(
                                e,
                                ess_team,
                                max_dt=e.sent_time,
                                subject_norm=subject_norm,
                            )
                        except Exception:
                            parsed_dt = None
                        if not parsed_dt:
                            continue
                        parsed_ist = _to_ist(parsed_dt)
                        if parsed_ist > (upper_ist + timedelta(minutes=5)):
                            continue
                        if parsed_ist < (upper_ist - timedelta(days=14)):
                            continue
                        if baseline_date and abs((parsed_ist.date() - baseline_date).days) > 5:
                            continue
                        parsed_candidates.append(parsed_ist)
                    if parsed_candidates:
                        if baseline_date:
                            pick_ist = min(parsed_candidates, key=lambda dt: (_baseline_rank_dt(dt), -dt.timestamp()))
                        else:
                            parsed_candidates.sort()
                            pick_ist = parsed_candidates[-1]
                        pick_src = "PARSED_FROM_QUOTED_REQUEST"

                if not pick_ist:
                    continue
                if mixed_owner_drift:
                    if pick_ist < (c_ist - timedelta(minutes=15)) or pick_ist > (c_ist + timedelta(hours=72)):
                        continue
                if baseline_date:
                    try:
                        current_delta = abs((c_ist.date() - baseline_date).days)
                        pick_delta = abs((pick_ist.date() - baseline_date).days)
                    except Exception:
                        current_delta = 9999
                        pick_delta = 9999
                    # Do not worsen day drift when a baseline is known.
                    if pick_delta > current_delta and pick_delta > 2:
                        continue
                if abs((pick_ist - c_ist).total_seconds()) <= 120:
                    continue
                if pick_ist <= (c_ist + timedelta(minutes=15)):
                    continue

                t = _format_time(pick.sent_time) if pick else _format_time(pick_ist)
                if not t:
                    continue
                row_vals["Created Date & Time"] = t
                if a_ist and a_ist < pick_ist:
                    row_vals["Actual Response Date & Time"] = t
                if r_ist and r_ist < pick_ist:
                    row_vals["Actual Resolved Date & Time"] = t

                row_idx = state.get("row_index")
                if row_idx:
                    ws.cell(row_idx, created_col).value = row_vals.get("Created Date & Time")
                    ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                    ws.cell(row_idx, resolved_col).value = row_vals.get("Actual Resolved Date & Time")
                if list_index < len(debug_rows):
                    who = pick_src
                    debug_rows[list_index]["CreatedSource"] = who
                    if a_ist and a_ist < pick_ist:
                        debug_rows[list_index]["AckSource"] = who
                    if r_ist and r_ist < pick_ist:
                        debug_rows[list_index]["ResolvedSource"] = who
                    if no_requester_recovery:
                        suffix = "; LiveRequestAnchorGuardNoRequesterRecovery"
                    elif mixed_owner_drift:
                        suffix = "; LiveRequestAnchorGuardMixedOwner"
                    else:
                        suffix = "; LiveRequestAnchorGuard"
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}{suffix}"

        if True:
            # ESS continuation guard (global):
            # If the row is ESS-initiated/span-like and the same consultant keeps
            # replying with no external request in-between, align to consultant tail.
            for state in row_states:
                list_index = state.get("list_index")
                if list_index is None or list_index >= len(automation_rows):
                    continue
                if state.get("is_dep_req") or state.get("is_dep_succ"):
                    continue

                row_vals = automation_rows[list_index]
                requester = state.get("requester") or ""
                subject_norm = (state.get("subject_norm") or "").lower()
                if not requester or not subject_norm:
                    continue

                notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
                notes_l = (notes_now or "").lower()
                created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
                ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
                resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
                mixed_source_pattern = (
                    isinstance(created_src_now, str)
                    and created_src_now.startswith("PARSED_FROM_")
                    and bool(ack_src_now)
                    and bool(requester)
                    and not _match_requester(ack_src_now, ack_src_now, requester)
                    and _match_requester(resolved_src_now, resolved_src_now, requester)
                )
                if (
                    "ess-only; no non-ess request" not in notes_l
                    and "requester follow-up (no in-between request)" not in notes_l
                    and not mixed_source_pattern
                ):
                    continue

                c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
                a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
                r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
                if not (c_dt and a_dt and r_dt):
                    continue
                c_ist = _to_ist(c_dt)
                a_ist = _to_ist(a_dt)
                r_ist = _to_ist(r_dt)
                if not (
                    (c_ist == a_ist and a_ist == r_ist)
                    or (mixed_source_pattern and (r_ist - c_ist) >= timedelta(hours=12))
                ):
                    continue

                base_thread = state.get("thread") or []
                thread = _expanded_thread(
                    subject_norm,
                    base_thread,
                    requester,
                    include_non_ess=True,
                    reference_ist=r_ist,
                )
                requester_pool = _requester_pool(subject_norm, requester, r_ist, day_window=21)
                merged_msgs = []
                for e in (thread or []):
                    merged_msgs.append(e)
                if not merged_msgs:
                    continue
                dedup = {}
                for e in merged_msgs:
                    dedup[(getattr(e, "subject", ""), getattr(e, "sender_email", ""), getattr(e, "sender_name", ""), getattr(e, "sent_time", None))] = e
                merged_msgs = list(dedup.values())
                merged_msgs.sort(key=lambda e: e.sent_time if getattr(e, "sent_time", None) else datetime.max)

                consultant_msgs = []
                for e in merged_msgs:
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if not _req_match(e, requester):
                        continue
                    if _ack_like(e) or _ack_like_text_fallback(e):
                        continue
                    consultant_msgs.append(e)
                consultant_msgs.sort(key=lambda e: e.sent_time)
                if len(consultant_msgs) < (1 if mixed_source_pattern else 2):
                    continue

                first_ist = _email_ist(consultant_msgs[0])
                latest = consultant_msgs[-1]
                latest_ist = _email_ist(latest)
                if not first_ist or not latest_ist:
                    continue

                non_ess_between = False
                row_tokens = _match_tokens(subject_norm)
                for e in emails:
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if e_ist <= first_ist or e_ist > latest_ist:
                        continue
                    if row_tokens:
                        s_norm = normalize_subject(getattr(e, "subject", "") or "")
                        s_tokens = _match_tokens(s_norm)
                        score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                        contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                        if score < 0.45 and not contains:
                            continue
                    if _ess_sender(e):
                        continue
                    if _system_like_sender(e):
                        continue
                    if _req_match(e, requester):
                        continue
                    non_ess_between = True
                    break
                if non_ess_between:
                    continue

                if (latest_ist - c_ist) > timedelta(days=5):
                    continue

                t = _format_time(latest.sent_time)
                if not t:
                    continue

                if c_ist == a_ist and a_ist == r_ist:
                    row_vals["Created Date & Time"] = t
                    row_vals["Actual Response Date & Time"] = t
                    row_vals["Actual Resolved Date & Time"] = t
                    mode = "AllThree"
                else:
                    if latest_ist <= (r_ist + timedelta(minutes=3)):
                        continue
                    if latest_ist >= c_ist:
                        row_vals["Actual Response Date & Time"] = t
                        row_vals["Actual Resolved Date & Time"] = t
                        mode = "ResponseResolved"
                    else:
                        continue

                row_idx = state.get("row_index")
                if row_idx:
                    ws.cell(row_idx, created_col).value = row_vals.get("Created Date & Time")
                    ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                    ws.cell(row_idx, resolved_col).value = row_vals.get("Actual Resolved Date & Time")
                if list_index < len(debug_rows):
                    who = latest.sender_email or latest.sender_name
                    if mode == "AllThree":
                        debug_rows[list_index]["CreatedSource"] = who
                        debug_rows[list_index]["AckSource"] = who
                    else:
                        debug_rows[list_index]["AckSource"] = who
                    debug_rows[list_index]["ResolvedSource"] = who
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ESSContinuationGuard[{mode}]"

        # Consultant-window environment guard:
        # Derive environment from consultant replies only, bounded by the selected
        # Created..Resolved timeframe, to avoid cross-request contamination in
        # long mixed threads.
        if env_col:
            subj_explicit_env_re = re.compile(
                r"\b(prod|prd|production|fcp|bip|uat|fct|biu|qa|fcq|biq|dev|development|fcd|bid)\b",
                flags=re.IGNORECASE,
            )
            for state in row_states:
                list_index = state.get("list_index")
                if list_index is None or list_index >= len(automation_rows):
                    continue
                if state.get("is_dep_req") or state.get("is_dep_succ"):
                    continue

                row_vals = automation_rows[list_index]
                requester = state.get("requester") or ""
                if not requester:
                    continue

                description = state.get("description") or ""
                subject_text = _subject_for_description(description)
                # Keep subject-explicit env untouched.
                if subject_text and subj_explicit_env_re.search(subject_text):
                    continue

                c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
                r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
                if not (c_dt and r_dt):
                    continue
                c_ist = _to_ist(c_dt)
                r_ist = _to_ist(r_dt)
                if r_ist < c_ist:
                    continue

                base_thread = state.get("thread") or []
                subject_norm = (state.get("subject_norm") or "").lower()
                thread = _expanded_thread(subject_norm, base_thread, requester)
                if not thread:
                    continue

                window_start = c_ist - timedelta(minutes=5)
                window_end = r_ist + timedelta(minutes=5)
                consultant_parts = []
                for e in thread:
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if e_ist < window_start or e_ist > window_end:
                        continue
                    if not _req_match(e, requester):
                        continue
                    if getattr(e, "body", None):
                        consultant_parts.append(e.body)
                    if getattr(e, "body_html", None):
                        consultant_parts.append(e.body_html)
                if not consultant_parts:
                    continue

                env_window = resolve_environment(subject_text, "\n".join(consultant_parts))
                if not env_window:
                    continue
                env_current = row_vals.get("Environment") or ""
                if env_window == env_current:
                    continue

                row_vals["Environment"] = env_window
                row_idx = state.get("row_index")
                if row_idx:
                    ws.cell(row_idx, env_col).value = env_window
                if list_index < len(debug_rows):
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; EnvConsultantWindowGuard"

        # Continuation all-three lock:
        # For ESS-only continuation/requester-span rows, if created is already chosen
        # correctly but response/resolved were later moved by post-ack passes, keep all
        # three equal to created.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            if "ess-only; no non-ess request" not in notes_l:
                continue
            if not (
                "requester span" in notes_l
                or "requester follow-up (no in-between request)" in notes_l
                or "requester follow-up(top-only)" in notes_l
                or "ess-only continuation(top requester)" in notes_l
                or "esscontinuationguard[allthree]" in notes_l
            ):
                continue
            if not (
                "resolvedafterackpost" in notes_l
                or "resolvedafterackrelated" in notes_l
                or "resolvedafterack" in notes_l
                or "resolvedwithin48hafterack" in notes_l
            ):
                continue

            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not c_dt:
                continue
            if a_dt and r_dt and _to_ist(a_dt) == _to_ist(c_dt) and _to_ist(r_dt) == _to_ist(c_dt):
                continue

            # Allow same-consultant and teammate continuation tails:
            # in ESS-only continuation rows, latest actionable reply can be by
            # requester or ESS teammate. Keep this generic and do not require
            # requester-only source matching here.
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            res_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            # Lock only when created is requester-owned; otherwise this can clobber
            # legitimate request->ack->resolve sequences.
            if requester and created_src_now:
                try:
                    if not _match_requester(str(created_src_now), str(created_src_now), requester):
                        continue
                except Exception:
                    continue
            elif requester:
                continue

            # Extra safety: do not force all-three if a non-ESS request exists in the
            # active episode or if requester has actionable (non-ack) follow-ups.
            subject_norm = (state.get("subject_norm") or "").lower()
            base_thread = state.get("thread") or []
            thread = _expanded_thread(
                subject_norm,
                base_thread,
                requester,
                include_non_ess=True,
                reference_ist=_to_ist(r_dt or a_dt or c_dt),
            )
            if thread:
                upper_ist = _to_ist(r_dt or a_dt or c_dt)
                lower_ist = _to_ist(c_dt) - timedelta(minutes=2)
                has_external_between = False
                requester_actionable_after_created = False
                for e in thread:
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if e_ist < lower_ist or e_ist > (upper_ist + timedelta(minutes=5)):
                        continue
                    if (not _ess_sender(e)) and (not _system_like_sender(e)):
                        has_external_between = True
                        break
                    if _req_match(e, requester) and e_ist > (_to_ist(c_dt) + timedelta(minutes=2)):
                        if (not _ack_like(e)) and (not _ack_like_text_fallback(e)):
                            requester_actionable_after_created = True
                            break
                if has_external_between or requester_actionable_after_created:
                    continue

            row_vals["Actual Response Date & Time"] = c
            row_vals["Actual Resolved Date & Time"] = c

            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, response_col).value = c
                ws.cell(row_idx, resolved_col).value = c

            if list_index < len(debug_rows):
                src = debug_rows[list_index].get("CreatedSource", "") or ack_src_now or res_src_now
                if src:
                    debug_rows[list_index]["AckSource"] = src
                    debug_rows[list_index]["ResolvedSource"] = src
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ContinuationAllThreeFinalLock"

        # Parsed-gap reanchor guard:
        # If Created came from parsed quote and response/resolved are much later,
        # re-anchor Created to the latest real non-ESS request before response.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            if not subject_norm:
                continue

            created_src_now = debug_rows[list_index].get("CreatedSource", "") if list_index < len(debug_rows) else ""
            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            if not (isinstance(created_src_now, str) and created_src_now.startswith("PARSED_")):
                continue
            if not (
                "monotonicguard" in notes_l
                or "resolvedafterack" in notes_l
                or "resolvedwithin48hafterack" in notes_l
            ):
                continue

            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            cutoff_dt = a_dt or r_dt
            if not (c_dt and cutoff_dt):
                continue
            c_ist = _to_ist(c_dt)
            cutoff_ist = _to_ist(cutoff_dt)
            if cutoff_ist <= c_ist:
                continue
            old_gap = cutoff_ist - c_ist
            if old_gap < timedelta(hours=12):
                continue

            base_thread = state.get("thread") or []
            thread = _expanded_thread(
                subject_norm,
                base_thread,
                requester,
                include_non_ess=True,
                reference_ist=cutoff_ist,
            )
            if not thread:
                continue

            row_tokens = _match_tokens(subject_norm)
            candidates = []
            for e in thread:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist > (cutoff_ist + timedelta(minutes=3)):
                    continue
                if _ess_sender(e) or _system_like_sender(e):
                    continue
                if row_tokens:
                    s_norm = normalize_subject(getattr(e, "subject", "") or "")
                    s_tokens = _match_tokens(s_norm)
                    score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                    contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                    if score < 0.45 and not contains:
                        continue
                candidates.append(e)

            if not candidates:
                continue
            pick = max(candidates, key=lambda e: e.sent_time)
            pick_ist = _email_ist(pick)
            if not pick_ist:
                continue
            if pick_ist <= (c_ist + timedelta(minutes=15)):
                continue
            new_gap = cutoff_ist - pick_ist
            if new_gap < timedelta(0) or new_gap > timedelta(days=7):
                continue
            if new_gap >= (old_gap - timedelta(minutes=5)):
                continue

            t = _format_time(pick.sent_time)
            if not t:
                continue
            row_vals["Created Date & Time"] = t
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = t
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = pick.sender_email or pick.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; ParsedGapReanchorGuard"

        # Baseline stale-created guard (safe + narrow):
        # When Created drifts >1 day away from ServiceNow baseline on risky/ambiguous
        # rows, re-anchor Created to the nearest valid request episode around baseline.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            baseline_date = state.get("baseline_created_date")
            if not baseline_date:
                continue

            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            subject_norm = (state.get("subject_norm") or "").lower()
            if not requester or not subject_norm:
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            risky = (
                "norequesterthreadrecovery" in notes_l
                or "no ess or requester replies" in notes_l
                or "ambiguousresolvedbyrequester" in notes_l
                or "score:" in notes_l
                or "altunion:" in notes_l
                or "dateanchormissing" in notes_l
                or "dateanchorafter" in notes_l
            )
            if not risky:
                continue

            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            if not c_dt:
                continue
            c_ist = _to_ist(c_dt)
            if abs((c_ist.date() - baseline_date).days) <= 1:
                continue

            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            a_ist = _to_ist(a_dt) if a_dt else None
            r_ist = _to_ist(r_dt) if r_dt else None

            baseline_mid = _to_ist(datetime(baseline_date.year, baseline_date.month, baseline_date.day))
            window_start = baseline_mid - timedelta(hours=18)
            window_end = baseline_mid + timedelta(hours=72)

            base_thread = state.get("thread") or []
            thread = _expanded_thread(
                subject_norm,
                base_thread,
                requester,
                include_non_ess=True,
                reference_ist=baseline_mid,
            )
            if not thread:
                continue

            row_tokens = _match_tokens(subject_norm)
            req_candidates = []
            for e in thread:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist < window_start or e_ist > window_end:
                    continue
                if _ess_sender(e) or _system_like_sender(e):
                    continue
                if row_tokens:
                    s_norm = normalize_subject(getattr(e, "subject", "") or "")
                    s_tokens = _match_tokens(s_norm)
                    score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                    contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                    if score < 0.45 and not contains:
                        continue
                req_candidates.append(e)

            pick_ist = None
            pick_src = ""
            if req_candidates:
                req_candidates.sort(key=lambda e: e.sent_time)
                pick = req_candidates[0]
                pick_ist = _email_ist(pick)
                pick_src = pick.sender_email or pick.sender_name
            else:
                # Quoted fallback from requester messages in the same episode window.
                quoted_candidates = []
                for e in thread:
                    if not e.sent_time or not _req_match(e, requester):
                        continue
                    e_ist = _email_ist(e)
                    if not e_ist:
                        continue
                    if e_ist < (window_start - timedelta(hours=6)) or e_ist > (window_end + timedelta(hours=6)):
                        continue
                    q_ist = _extract_quoted_request_before_ist(e, subject_norm, a_ist or r_ist or e_ist)
                    if not q_ist:
                        continue
                    if q_ist < window_start or q_ist > window_end:
                        continue
                    quoted_candidates.append(q_ist)
                if quoted_candidates:
                    quoted_candidates.sort()
                    pick_ist = quoted_candidates[-1]
                    pick_src = "PARSED_FROM_QUOTED_REQUEST"

            if not pick_ist:
                continue
            if abs((pick_ist - c_ist).total_seconds()) <= 120:
                continue

            t = _format_time(pick_ist)
            if not t:
                continue

            row_vals["Created Date & Time"] = t
            if a_ist and a_ist < pick_ist:
                row_vals["Actual Response Date & Time"] = t
            if r_ist and r_ist < pick_ist:
                row_vals["Actual Resolved Date & Time"] = t

            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = row_vals.get("Created Date & Time")
                ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                ws.cell(row_idx, resolved_col).value = row_vals.get("Actual Resolved Date & Time")
            if list_index < len(debug_rows):
                debug_rows[list_index]["CreatedSource"] = pick_src
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BaselineStaleCreatedGuard"

        # Risk-row episode lock (isolated, opt-in):
        # For clearly problematic rows, freeze all three timestamps from one
        # consultant/request episode to avoid cross-episode mixing.
        # Disabled by default because it can over-correct stable rows.
        locked_list_indexes = set()
        enable_risk_episode_lock = os.getenv("ENABLE_RISK_EPISODE_LOCK", "0") == "1"
        if enable_risk_episode_lock:
            for state in row_states:
                list_index = state.get("list_index")
                if list_index is None or list_index >= len(automation_rows):
                    continue
                if state.get("is_dep_req") or state.get("is_dep_succ"):
                    continue

                row_vals = automation_rows[list_index]
                requester = state.get("requester") or ""
                subject_norm = (state.get("subject_norm") or "").lower()
                if not requester or not subject_norm:
                    continue

                notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
                notes_l = (notes_now or "").lower()
                baseline_date = state.get("baseline_created_date")
                ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
                resolved_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""

                c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
                a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
                r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
                if not (c_dt and a_dt and r_dt):
                    continue
                c_ist = _to_ist(c_dt)
                a_ist = _to_ist(a_dt)
                r_ist = _to_ist(r_dt)

                baseline_drift = 0
                if baseline_date:
                    try:
                        baseline_drift = abs((c_ist.date() - baseline_date).days)
                    except Exception:
                        baseline_drift = 0
                source_mixed = False
                if requester and ack_src_now and resolved_src_now:
                    try:
                        source_mixed = (
                            (not _match_requester(str(ack_src_now), str(ack_src_now), requester))
                            and _match_requester(str(resolved_src_now), str(resolved_src_now), requester)
                        )
                    except Exception:
                        source_mixed = False
                # Keep this pass on clearly divergent rows only; 12h was too sensitive
                # and pulled many stable rows into expensive episode locking.
                large_gap = (a_ist - c_ist) > timedelta(hours=24)

                strong_risk_markers = (
                    "norequesterthreadrecovery",
                    "no ess or requester replies",
                    "ambiguousresolvedbyrequester",
                )
                weak_risk_markers = (
                    "dateanchormissing",
                    "dateanchorafter",
                )
                strong_risky = any(m in notes_l for m in strong_risk_markers)
                weak_risky = any(m in notes_l for m in weak_risk_markers)
                # Only allow weak marker rows when there is a strong drift signal.
                if not strong_risky and not (weak_risky and (baseline_drift >= 3 or source_mixed or large_gap)):
                    continue
                # Keep lock pass narrow for performance and safety.
                if baseline_drift <= 2 and (not source_mixed) and (not large_gap):
                    continue

                base_thread = state.get("thread") or []
                ref_ist = c_ist
                if baseline_date:
                    try:
                        ref_ist = _to_ist(datetime(baseline_date.year, baseline_date.month, baseline_date.day))
                    except Exception:
                        ref_ist = c_ist
                thread = _expanded_thread(
                    subject_norm,
                    base_thread,
                    requester,
                    include_non_ess=True,
                    reference_ist=ref_ist,
                )
                if not thread:
                    continue

                row_tokens = _match_tokens(subject_norm)
                subject_fit_cache = {}

                def _subject_fit(email_obj):
                    k = id(email_obj)
                    if k in subject_fit_cache:
                        return subject_fit_cache[k]
                    if not row_tokens:
                        subject_fit_cache[k] = True
                        return True
                    s_norm = normalize_subject(getattr(email_obj, "subject", "") or "")
                    s_tokens = _match_tokens(s_norm)
                    score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                    contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                    ok = score >= 0.45 or contains
                    subject_fit_cache[k] = ok
                    return ok

                consultant_msgs = [
                    e for e in thread
                    if getattr(e, "sent_time", None)
                    and _req_match(e, requester)
                    and _email_ist(e)
                    and _subject_fit(e)
                ]
                consultant_msgs.sort(key=lambda e: e.sent_time)
                if not consultant_msgs:
                    continue

                episodes = []
                current = [consultant_msgs[0]]
                for e in consultant_msgs[1:]:
                    prev_ist = _email_ist(current[-1])
                    now_ist = _email_ist(e)
                    if prev_ist and now_ist and (now_ist - prev_ist) > timedelta(hours=24):
                        episodes.append(current)
                        current = [e]
                    else:
                        current.append(e)
                episodes.append(current)

                anchor_ist = a_ist or r_ist or c_ist
                if baseline_date:
                    try:
                        anchor_ist = _to_ist(datetime(baseline_date.year, baseline_date.month, baseline_date.day))
                    except Exception:
                        pass

                best_episode = None
                best_rank = None
                for ep in episodes:
                    non_ack_ep = [e for e in ep if not _ack_like(e) and not _ack_like_text_fallback(e)]
                    seed = non_ack_ep[0] if non_ack_ep else ep[0]
                    seed_ist = _email_ist(seed)
                    if not seed_ist:
                        continue
                    if baseline_date:
                        day_delta = abs((seed_ist.date() - baseline_date).days)
                    else:
                        day_delta = abs((seed_ist - anchor_ist).total_seconds()) / 86400.0
                    prox = abs((seed_ist - anchor_ist).total_seconds())
                    ack_penalty = 1 if not non_ack_ep else 0
                    rank = (day_delta, prox, ack_penalty)
                    if best_rank is None or rank < best_rank:
                        best_rank = rank
                        best_episode = ep

                if not best_episode:
                    continue

                non_ack_best = [e for e in best_episode if not _ack_like(e) and not _ack_like_text_fallback(e)]
                response_mail = non_ack_best[0] if non_ack_best else best_episode[0]
                resolved_mail = non_ack_best[-1] if non_ack_best else best_episode[-1]
                response_ist = _email_ist(response_mail)
                if not response_ist:
                    continue

                non_ess_reqs = [
                    e for e in thread
                    if getattr(e, "sent_time", None)
                    and _email_ist(e)
                    and _subject_fit(e)
                    and not _ess_sender(e)
                    and not _system_like_sender(e)
                    and _email_ist(e) <= response_ist
                    and _email_ist(e) >= (response_ist - timedelta(days=7))
                ]
                if baseline_date:
                    by_day = [e for e in non_ess_reqs if _email_ist(e).date() == baseline_date]
                    if by_day:
                        non_ess_reqs = by_day
                created_mail = max(non_ess_reqs, key=lambda e: e.sent_time) if non_ess_reqs else response_mail

                t_c = _format_time(created_mail.sent_time)
                t_a = _format_time(response_mail.sent_time)
                t_r = _format_time(resolved_mail.sent_time)
                if not (t_c and t_a and t_r):
                    continue

                row_vals["Created Date & Time"] = t_c
                row_vals["Actual Response Date & Time"] = t_a
                row_vals["Actual Resolved Date & Time"] = t_r
                row_idx = state.get("row_index")
                if row_idx:
                    ws.cell(row_idx, created_col).value = t_c
                    ws.cell(row_idx, response_col).value = t_a
                    ws.cell(row_idx, resolved_col).value = t_r
                if list_index < len(debug_rows):
                    debug_rows[list_index]["CreatedSource"] = created_mail.sender_email or created_mail.sender_name
                    debug_rows[list_index]["AckSource"] = response_mail.sender_email or response_mail.sender_name
                    debug_rows[list_index]["ResolvedSource"] = resolved_mail.sender_email or resolved_mail.sender_name
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; RiskEpisodeLock"
                locked_list_indexes.add(list_index)

        # Final monotonic safety:
        # ensure Created <= Ack <= Resolved after all guards.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if list_index in locked_list_indexes:
                continue
            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            c = row_vals.get("Created Date & Time")
            a = row_vals.get("Actual Response Date & Time")
            r = row_vals.get("Actual Resolved Date & Time")
            c_dt = _parse_time_str(c)
            a_dt = _parse_time_str(a)
            r_dt = _parse_time_str(r)
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)

            changed = False
            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            res_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""

            if a_ist < c_ist:
                repaired_created = False
                # Try safe requester/non-ESS re-anchor before collapsing ack to created.
                base_thread = state.get("thread") or []
                subject_norm = (state.get("subject_norm") or "").lower()
                thread = _expanded_thread(subject_norm, base_thread, requester, include_non_ess=True, reference_ist=a_ist)
                if thread:
                    candidate_reqs = []
                    for e in thread:
                        e_ist = _email_ist(e)
                        if not e_ist:
                            continue
                        if e_ist > (a_ist + timedelta(minutes=3)):
                            continue
                        if e_ist < (a_ist - timedelta(days=14)):
                            continue
                        if _ess_sender(e) or _system_like_sender(e):
                            continue
                        candidate_reqs.append(e)
                    if candidate_reqs:
                        pick_req = max(candidate_reqs, key=lambda e: e.sent_time)
                        pick_req_ist = _email_ist(pick_req)
                        if pick_req_ist and pick_req_ist <= a_ist:
                            new_c = _format_time(pick_req.sent_time)
                            if new_c:
                                row_vals["Created Date & Time"] = new_c
                                c = new_c
                                c_dt = _parse_time_str(c)
                                c_ist = _to_ist(c_dt) if c_dt else c_ist
                                repaired_created = True
                                changed = True
                                if list_index < len(debug_rows):
                                    debug_rows[list_index]["CreatedSource"] = pick_req.sender_email or pick_req.sender_name
                                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; MonotonicCreatedRepair"
                if a_ist < c_ist:
                    a = c
                    a_dt = _parse_time_str(a)
                    a_ist = _to_ist(a_dt) if a_dt else c_ist
                    row_vals["Actual Response Date & Time"] = a
                    changed = True
            if r_ist < a_ist:
                prefer_back_anchor = (
                    "requester span(all-ack->ess)" in (notes_now or "").lower()
                    or "requester span(all-ack->requester-fallback)" in (notes_now or "").lower()
                    or (
                        str(ack_src_now).strip().upper() == "ACK NOT FOUND"
                        and c == a
                        and res_src_now
                        and requester
                        and not _match_requester(res_src_now, res_src_now, requester)
                    )
                )

                if prefer_back_anchor:
                    # Keep older non-ack resolved episode and re-anchor created/ack back.
                    a = r
                    row_vals["Actual Response Date & Time"] = a
                    a_dt = _parse_time_str(a)
                    a_ist = _to_ist(a_dt) if a_dt else r_ist
                    if c_ist > a_ist:
                        c = a
                        row_vals["Created Date & Time"] = c
                        c_dt = _parse_time_str(c)
                        c_ist = _to_ist(c_dt) if c_dt else a_ist
                    changed = True
                    if list_index < len(debug_rows):
                        debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; MonotonicBackAnchorGuard"
                else:
                    # Before collapsing resolved to ack, try a safe requester-based
                    # repair in the same thread episode.
                    base_thread = state.get("thread") or []
                    subject_norm = (state.get("subject_norm") or "").lower()
                    thread = _expanded_thread(subject_norm, base_thread, requester)
                    repaired = False
                    if thread and requester:
                        repair_candidates = [
                            e for e in thread
                            if e.sent_time
                            and _req_match(e, requester)
                            and _email_ist(e)
                            and _email_ist(e) >= a_ist
                            and _email_ist(e) <= (a_ist + timedelta(hours=48))
                            and not _ack_like(e)
                            and not _ack_like_text_fallback(e)
                        ]
                        repair_candidates.sort(key=lambda e: e.sent_time)
                        if repair_candidates:
                            pick = repair_candidates[0]
                            repaired_r = _format_time(pick.sent_time)
                            if repaired_r:
                                r = repaired_r
                                row_vals["Actual Resolved Date & Time"] = r
                                repaired = True
                                changed = True
                                if list_index < len(debug_rows):
                                    debug_rows[list_index]["ResolvedSource"] = pick.sender_email or pick.sender_name
                                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; MonotonicResolvedRepair"
                    if not repaired:
                        r = a
                        row_vals["Actual Resolved Date & Time"] = r
                        changed = True

            if not changed:
                continue
            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, created_col).value = row_vals.get("Created Date & Time")
                ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                ws.cell(row_idx, resolved_col).value = row_vals.get("Actual Resolved Date & Time")
            if list_index < len(debug_rows):
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; MonotonicGuard"

        # Mixed-owner episode clamp (isolated):
        # When Created is early, Response/Resolved are very late, and ownership is
        # split (Ack non-requester, Resolved requester), clamp to requester's local
        # episode near Created instead of a later drifted episode.
        for state in row_states:
            list_index = state.get("list_index")
            if list_index is None or list_index >= len(automation_rows):
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            if not requester:
                continue

            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            if not (c_dt and a_dt and r_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt)
            if (a_ist - c_ist) <= timedelta(hours=18):
                continue

            notes_now = debug_rows[list_index].get("Notes", "") if list_index < len(debug_rows) else ""
            notes_l = (notes_now or "").lower()
            if not (
                "monotonicguard" in notes_l
                or "liverequestanchorguardmixedowner" in notes_l
            ):
                continue

            ack_src_now = debug_rows[list_index].get("AckSource", "") if list_index < len(debug_rows) else ""
            res_src_now = debug_rows[list_index].get("ResolvedSource", "") if list_index < len(debug_rows) else ""
            if not (ack_src_now and res_src_now):
                continue
            try:
                ack_is_req = _match_requester(str(ack_src_now), str(ack_src_now), requester)
                res_is_req = _match_requester(str(res_src_now), str(res_src_now), requester)
            except Exception:
                continue
            if ack_is_req or (not res_is_req):
                continue

            subject_norm = (state.get("subject_norm") or "").lower()
            base_thread = state.get("thread") or []
            thread = _expanded_thread(
                subject_norm,
                base_thread,
                requester,
                include_non_ess=True,
                reference_ist=r_ist,
            )
            if not thread:
                continue

            win_start = c_ist - timedelta(minutes=5)
            win_end = min(c_ist + timedelta(hours=72), r_ist + timedelta(minutes=5))
            req_local = []
            for e in thread:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist < win_start or e_ist > win_end:
                    continue
                if not _req_match(e, requester):
                    continue
                if _ack_like(e) or _ack_like_text_fallback(e):
                    continue
                req_local.append(e)
            if not req_local:
                continue

            req_local.sort(key=lambda e: e.sent_time)
            response_pick = req_local[0]
            response_pick_ist = _email_ist(response_pick)
            if not response_pick_ist:
                continue
            # Apply only when materially earlier than current response.
            if response_pick_ist >= (a_ist - timedelta(minutes=5)):
                continue

            resolved_local = [
                e for e in req_local
                if _email_ist(e) >= response_pick_ist
                and _email_ist(e) <= (response_pick_ist + timedelta(hours=72))
            ]
            resolved_pick = resolved_local[-1] if resolved_local else response_pick

            t_a = _format_time(response_pick.sent_time)
            t_r = _format_time(resolved_pick.sent_time)
            if not (t_a and t_r):
                continue
            row_vals["Actual Response Date & Time"] = t_a
            row_vals["Actual Resolved Date & Time"] = t_r

            row_idx = state.get("row_index")
            if row_idx:
                ws.cell(row_idx, response_col).value = t_a
                ws.cell(row_idx, resolved_col).value = t_r
            if list_index < len(debug_rows):
                debug_rows[list_index]["AckSource"] = response_pick.sender_email or response_pick.sender_name
                debug_rows[list_index]["ResolvedSource"] = resolved_pick.sender_email or resolved_pick.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; MixedOwnerEpisodeClamp"

        # Blue-gap final audit:
        # - Mark rows blue when Created->Response gap > 16 minutes.
        # - Before marking, try one safe local re-anchor for missed ack.
        # - If re-anchor fixes the gap, clear blue.
        blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        clear_fill = PatternFill(fill_type=None)

        def _is_blue_cell_fill(cell):
            try:
                rgb = (cell.fill.start_color.rgb or "").upper()
                return rgb.endswith("BDD7EE")
            except Exception:
                return False

        def _row_has_blue_fill(row_idx):
            for col in range(1, ws.max_column + 1):
                if _is_blue_cell_fill(ws.cell(row_idx, col)):
                    return True
            return False

        def _set_row_fill(row_idx, fill_obj):
            for col in range(1, ws.max_column + 1):
                ws.cell(row_idx, col).fill = fill_obj

        for state in row_states:
            list_index = state.get("list_index")
            row_idx = state.get("row_index")
            if list_index is None or list_index >= len(automation_rows) or not row_idx:
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue

            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            if not requester:
                continue

            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            if not (c_dt and a_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt) if r_dt else None

            # Safe local re-anchor only for rows currently >16m.
            if (a_ist - c_ist) > timedelta(minutes=16):
                subject_norm = (state.get("subject_norm") or "").lower()
                base_thread = state.get("thread") or []
                thread = _expanded_thread(
                    subject_norm,
                    base_thread,
                    requester,
                    include_non_ess=True,
                    reference_ist=a_ist,
                )
                if thread:
                    local_req = []
                    for e in thread:
                        e_ist = _email_ist(e)
                        if not e_ist:
                            continue
                        if not _req_match(e, requester):
                            continue
                        if e_ist <= c_ist:
                            continue
                        if e_ist > (c_ist + timedelta(minutes=16)):
                            continue
                        local_req.append(e)
                    if local_req:
                        local_req.sort(key=lambda e: e.sent_time)
                        ack_pick = local_req[0]
                        ack_pick_ist = _email_ist(ack_pick)
                        if ack_pick_ist and ack_pick_ist < a_ist:
                            t_a = _format_time(ack_pick.sent_time)
                            if t_a:
                                row_vals["Actual Response Date & Time"] = t_a
                                a_dt = _parse_time_str(t_a)
                                a_ist = _to_ist(a_dt) if a_dt else a_ist
                                if r_ist and r_ist < a_ist:
                                    row_vals["Actual Resolved Date & Time"] = t_a
                                    r_ist = a_ist
                                ws.cell(row_idx, response_col).value = row_vals.get("Actual Response Date & Time")
                                ws.cell(row_idx, resolved_col).value = row_vals.get("Actual Resolved Date & Time")
                                if list_index < len(debug_rows):
                                    debug_rows[list_index]["AckSource"] = ack_pick.sender_email or ack_pick.sender_name
                                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BlueGapReanchor"

            # Final blue decision after re-anchor.
            a_dt2 = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            if not a_dt2:
                continue
            a_ist2 = _to_ist(a_dt2)
            is_blue_gap = (a_ist2 - c_ist) > timedelta(minutes=16)
            has_blue = _row_has_blue_fill(row_idx)
            if is_blue_gap:
                if not has_blue:
                    _set_row_fill(row_idx, blue_fill)
                if list_index < len(debug_rows):
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BlueGap>16m"
            else:
                if has_blue:
                    _set_row_fill(row_idx, clear_fill)
                if list_index < len(debug_rows):
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BlueCleared"

        # Blue-only strict realign pass (isolated):
        # Process only rows already marked blue, so stricter logic cannot disturb
        # correctly filled rows.
        for state in row_states:
            list_index = state.get("list_index")
            row_idx = state.get("row_index")
            if list_index is None or list_index >= len(automation_rows) or not row_idx:
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue
            if not _row_has_blue_fill(row_idx):
                continue

            row_vals = automation_rows[list_index]
            requester = state.get("requester") or ""
            if not requester:
                continue

            c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
            a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            r_dt = _parse_time_str(row_vals.get("Actual Resolved Date & Time"))
            if not (c_dt and a_dt):
                continue
            c_ist = _to_ist(c_dt)
            a_ist = _to_ist(a_dt)
            r_ist = _to_ist(r_dt) if r_dt else a_ist
            old_gap = a_ist - c_ist
            if old_gap <= timedelta(minutes=16):
                _set_row_fill(row_idx, clear_fill)
                if list_index < len(debug_rows):
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BlueCleared"
                continue

            subject_norm = (state.get("subject_norm") or "").lower()
            base_thread = state.get("thread") or []
            thread = _expanded_thread(
                subject_norm,
                base_thread,
                requester,
                include_non_ess=True,
                reference_ist=a_ist,
            )
            if not thread:
                continue

            row_tokens = _match_tokens(subject_norm)
            baseline_date = state.get("baseline_created_date")
            win_start = c_ist - timedelta(minutes=5)
            win_end = c_ist + timedelta(hours=96)
            strict_candidates = []
            for e in thread:
                e_ist = _email_ist(e)
                if not e_ist:
                    continue
                if e_ist < win_start or e_ist > win_end:
                    continue
                if not _req_match(e, requester):
                    continue
                if _ack_like(e) or _ack_like_text_fallback(e):
                    continue
                if baseline_date and abs((e_ist.date() - baseline_date).days) > 2:
                    continue
                if row_tokens:
                    s_norm = normalize_subject(getattr(e, "subject", "") or "")
                    s_tokens = _match_tokens(s_norm)
                    score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                    contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                    if score < 0.45 and not contains:
                        continue
                strict_candidates.append(e)
            if not strict_candidates:
                continue

            strict_candidates.sort(key=lambda e: e.sent_time)
            response_pick = None
            for e in strict_candidates:
                if _email_ist(e) >= c_ist:
                    response_pick = e
                    break
            if not response_pick:
                response_pick = strict_candidates[0]
            response_pick_ist = _email_ist(response_pick)
            if not response_pick_ist:
                continue

            resolved_pool = [
                e for e in strict_candidates
                if _email_ist(e) >= response_pick_ist
                and _email_ist(e) <= (response_pick_ist + timedelta(hours=72))
            ]
            resolved_pick = resolved_pool[-1] if resolved_pool else response_pick
            resolved_pick_ist = _email_ist(resolved_pick)
            if not resolved_pick_ist:
                continue

            new_gap = response_pick_ist - c_ist
            # Apply only if strictly improves (or fully resolves) the blue gap.
            if new_gap >= old_gap:
                continue

            t_a = _format_time(response_pick.sent_time)
            t_r = _format_time(resolved_pick.sent_time)
            if not (t_a and t_r):
                continue
            row_vals["Actual Response Date & Time"] = t_a
            row_vals["Actual Resolved Date & Time"] = t_r
            ws.cell(row_idx, response_col).value = t_a
            ws.cell(row_idx, resolved_col).value = t_r
            if list_index < len(debug_rows):
                debug_rows[list_index]["AckSource"] = response_pick.sender_email or response_pick.sender_name
                debug_rows[list_index]["ResolvedSource"] = resolved_pick.sender_email or resolved_pick.sender_name
                debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BlueStrictEpisodeAlign"

            # Re-evaluate blue after strict align.
            a_dt2 = _parse_time_str(row_vals.get("Actual Response Date & Time"))
            if not a_dt2:
                continue
            a_ist2 = _to_ist(a_dt2)
            if (a_ist2 - c_ist) <= timedelta(minutes=16):
                _set_row_fill(row_idx, clear_fill)
                if list_index < len(debug_rows):
                    debug_rows[list_index]["Notes"] = f"{debug_rows[list_index].get('Notes','')}; BlueClearedStrict"

        # Duplicate same-time de-collision (isolated):
        # If multiple rows for the same requester ended up with identical
        # created/response/resolved, move only secondary rows to another valid
        # requester episode without touching non-duplicate rows.
        duplicate_groups = {}
        for state in row_states:
            list_index = state.get("list_index")
            row_idx = state.get("row_index")
            if list_index is None or list_index >= len(automation_rows) or not row_idx:
                continue
            if state.get("is_dep_req") or state.get("is_dep_succ"):
                continue
            row_vals = automation_rows[list_index]
            req = _requester_key(state.get("requester") or "")
            c = row_vals.get("Created Date & Time") or ""
            a = row_vals.get("Actual Response Date & Time") or ""
            r = row_vals.get("Actual Resolved Date & Time") or ""
            if not (req and c and a and r):
                continue
            key = (req, c, a, r)
            duplicate_groups.setdefault(key, []).append(state)

        for dkey, group in duplicate_groups.items():
            if len(group) < 2:
                continue

            # Keep one stable anchor row untouched:
            # prefer non-ambiguous match; otherwise earliest row index.
            anchor_state = None
            for s in group:
                li = s.get("list_index")
                notes = debug_rows[li].get("Notes", "") if li is not None and li < len(debug_rows) else ""
                notes_l = (notes or "").lower()
                ambiguous = ("match=score:" in notes_l) or ("ambiguous" in notes_l)
                if not ambiguous:
                    anchor_state = s
                    break
            if anchor_state is None:
                anchor_state = min(group, key=lambda x: x.get("row_index") or 10**9)
            anchor_subject_norm = (anchor_state.get("subject_norm") or "").lower()
            anchor_tokens = _match_tokens(anchor_subject_norm)

            used_response_ists = set()
            anchor_li = anchor_state.get("list_index")
            anchor_a_dt = _parse_time_str(automation_rows[anchor_li].get("Actual Response Date & Time")) if anchor_li is not None else None
            if anchor_a_dt:
                used_response_ists.add(_to_ist(anchor_a_dt).replace(second=0, microsecond=0))

            for s in group:
                li = s.get("list_index")
                ri = s.get("row_index")
                if li is None or li >= len(automation_rows) or not ri:
                    continue
                if s is anchor_state:
                    continue

                row_vals = automation_rows[li]
                requester = s.get("requester") or ""
                subject_norm = (s.get("subject_norm") or "").lower()
                # Same consultant + same-subject family only.
                if anchor_tokens:
                    row_tokens_for_group = _match_tokens(subject_norm)
                    sim_group = _token_overlap_score(anchor_tokens, row_tokens_for_group) if row_tokens_for_group else 0.0
                    contains_group = bool(
                        anchor_subject_norm and subject_norm and (
                            anchor_subject_norm in subject_norm or subject_norm in anchor_subject_norm
                        )
                    )
                    if sim_group < 0.40 and not contains_group:
                        continue
                a_dt = _parse_time_str(row_vals.get("Actual Response Date & Time"))
                c_dt = _parse_time_str(row_vals.get("Created Date & Time"))
                if not (requester and subject_norm and a_dt):
                    continue
                a_ist = _to_ist(a_dt)
                c_ist = _to_ist(c_dt) if c_dt else None
                baseline_date = s.get("baseline_created_date")
                created_src_now = debug_rows[li].get("CreatedSource", "") if li < len(debug_rows) else ""

                base_thread = s.get("thread") or []
                thread = _expanded_thread(
                    subject_norm,
                    base_thread,
                    requester,
                    include_non_ess=True,
                    reference_ist=a_ist,
                )
                if not thread:
                    continue

                row_tokens = _match_tokens(subject_norm)
                def _collect_candidates(max_days: int, enforce_baseline: bool):
                    out = []
                    for e in thread:
                        e_ist = _email_ist(e)
                        if not e_ist:
                            continue
                        if not _req_match(e, requester):
                            continue
                        if _ack_like(e) or _ack_like_text_fallback(e):
                            continue
                        if abs((e_ist - a_ist).total_seconds()) <= 60:
                            continue
                        if abs((e_ist - a_ist).total_seconds()) > (max_days * 24 * 3600):
                            continue
                        if enforce_baseline and baseline_date and abs((e_ist.date() - baseline_date).days) > 3:
                            continue
                        if row_tokens:
                            s_norm = normalize_subject(getattr(e, "subject", "") or "")
                            s_tokens = _match_tokens(s_norm)
                            score = _token_overlap_score(row_tokens, s_tokens) if s_tokens else 0.0
                            contains = bool(subject_norm and s_norm and (subject_norm in s_norm or s_norm in subject_norm))
                            if score < 0.45 and not contains:
                                continue
                        out.append(e)
                    return out

                candidates = _collect_candidates(max_days=14, enforce_baseline=True)
                if not candidates:
                    candidates = _collect_candidates(max_days=45, enforce_baseline=False)
                if not candidates:
                    continue

                # Pick nearest candidate not already used by anchored duplicates.
                candidates.sort(key=lambda e: abs((_email_ist(e) - a_ist).total_seconds()))
                pick = None
                for e in candidates:
                    tkey = _email_ist(e).replace(second=0, microsecond=0)
                    if tkey in used_response_ists:
                        continue
                    pick = e
                    break
                if not pick:
                    continue

                t = _format_time(pick.sent_time)
                if not t:
                    continue
                row_vals["Actual Response Date & Time"] = t
                row_vals["Actual Resolved Date & Time"] = t
                pick_ist = _email_ist(pick)
                # If created was parsed and sits after the selected episode, align it.
                if (
                    pick_ist
                    and c_ist
                    and pick_ist < c_ist
                    and isinstance(created_src_now, str)
                    and created_src_now.startswith("PARSED_FROM_")
                ):
                    row_vals["Created Date & Time"] = t
                    ws.cell(ri, created_col).value = t
                ws.cell(ri, response_col).value = t
                ws.cell(ri, resolved_col).value = t
                used_response_ists.add(_email_ist(pick).replace(second=0, microsecond=0))
                if li < len(debug_rows):
                    who = pick.sender_email or pick.sender_name
                    if (
                        pick_ist
                        and c_ist
                        and pick_ist < c_ist
                        and isinstance(created_src_now, str)
                        and created_src_now.startswith("PARSED_FROM_")
                    ):
                        debug_rows[li]["CreatedSource"] = who
                    debug_rows[li]["AckSource"] = who
                    debug_rows[li]["ResolvedSource"] = who
                    debug_rows[li]["Notes"] = f"{debug_rows[li].get('Notes','')}; DuplicateDecollision"


    fill_result = fill_template(
        template_path=template_path,
        output_path=output_dir / "Support_Tracker_DEC_25_Incident_Business_done_filled.xlsx",
        row_resolver=resolve_row,
        logger=logger,
        post_process=_sequence_ordering_pass,
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
