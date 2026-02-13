import os
import sys
import re
from datetime import datetime, timedelta
from pathlib import Path

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
        if len(consultant_after) >= 2:
            first_dt = _to_ist(consultant_after[0].sent_time)
            if first_dt <= ack_ist + timedelta(minutes=grace_minutes) and _is_ack_like_reply(consultant_after[0]):
                return consultant_after[1]
        return consultant_after[0]

    def _find_unique_requester_date_thread(subject_norm, requester, date_tokens, iface_tokens=None):
        if not subject_norm or not requester or not date_tokens:
            return None
        iface_tokens = iface_tokens or set()
        subj_tokens = _match_tokens(subject_norm)
        if not subj_tokens:
            return None
        hits = []
        for key, thread in threads.items():
            if not _thread_has_requester(thread, requester):
                continue
            key_tokens = _match_tokens(key)
            if not key_tokens:
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
        best = None
        for key, thread in threads.items():
            if not _thread_has_requester(thread, requester):
                continue
            key_tokens = _match_tokens(key)
            if not key_tokens:
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
            min_delta = None
            for e in consultant_replies:
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
                    baseline_mid = _to_ist(datetime(
                        baseline_created_date.year,
                        baseline_created_date.month,
                        baseline_created_date.day,
                    ))
                    def _rank(e):
                        sent = _to_ist(e.sent_time)
                        day_delta = abs((sent.date() - baseline_created_date).days)
                        return (day_delta, abs(sent - baseline_mid))
                    best = min(requester_replies, key=_rank)
                    best_day_delta = abs((_to_ist(best.sent_time).date() - baseline_created_date).days)
                    if best_day_delta <= 2 and best_day_delta < current_delta:
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
        if thread is None:
            thread, match_note = find_thread(
                subject_norm,
                requester,
                date_tokens=date_tokens,
                prefer_consultant_date=explicit_marker,
                iface_tokens=iface_hint_tokens,
                baseline_date=baseline_created_date,
            )
        # When a stale date marker existed, refine thread choice using requester-reply
        # proximity to the row's original ServiceNow baseline date.
        if thread and requester and baseline_created_date and stale_anchor:
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

        consultant_body_text = ""
        if thread and requester:
            consultant_bodies = []
            for e in thread:
                if _match_requester(e.sender_name, e.sender_email, requester):
                    if e.body:
                        consultant_bodies.append(e.body)
            consultant_body_text = "\n".join(consultant_bodies)

        # Environment: subject -> consultant replies -> full thread -> description (if no thread)
        env = resolve_environment(subject_text, consultant_body_text)
        if not env:
            if thread:
                all_body_text = "\n".join((e.body or "") for e in thread)
                env = resolve_environment(subject_text, all_body_text)
            else:
                env = resolve_environment(subject_text, description or "")
        interface_code = resolve_interface_code(description)
        service_request = resolve_service_request(category_type)
        incident_type = resolve_incident_type(category_type, description)

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
            if not row_dr_ids:
                # Fallback: extract DR IDs from the row description if subject match failed
                row_dr_ids |= _extract_dr_ids(description or "")
                row_dr_ids |= _extract_dr_ids(subject_text or "")

            for dr in row_dr_ids:
                bucket = deployment_index.get(dr)
                if not bucket:
                    continue
                req_candidates = bucket["request"]
                succ_candidates = bucket["success"]
                if not req_candidates or not succ_candidates:
                    continue

                # Require environment match (PROD/UAT) when present on the row.
                if row_env:
                    req_candidates = [c for c in req_candidates if _dep_env(c["subject_key"]) == row_env]
                    succ_candidates = [c for c in succ_candidates if _dep_env(c["subject_key"]) == row_env]
                    if not req_candidates or not succ_candidates:
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
                    req_thread = thread if thread else min(req_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]
                    succ_thread = min(succ_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]
                else:
                    succ_thread = thread if thread else min(succ_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]
                    req_thread = min(req_candidates, key=lambda c: min(e.sent_time for e in c["thread"]))["thread"]

                req_email = min(req_thread, key=lambda e: e.sent_time)
                succ_email = min(succ_thread, key=lambda e: e.sent_time)
                times = TimeResult(
                    _format_time(req_email.sent_time),
                    _format_time(req_email.sent_time),
                    _format_time(succ_email.sent_time),
                )
                debug = TimeDebug(
                    req_email.sender_email or req_email.sender_name,
                    req_email.sender_email or req_email.sender_name,
                    succ_email.sender_email or succ_email.sender_name,
                    f"DeploymentPair DR={dr}; Match={match_note}",
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
                            latest_all = consultant_all[-1]
                            if _to_ist(latest_all.sent_time) == _to_ist(pick.sent_time):
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
                consultant_replies = [
                    e for e in thread
                    if _match_requester(e.sender_name, e.sender_email, requester)
                ]
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
                consultant_replies = [
                    e for e in thread
                    if _match_requester(e.sender_name, e.sender_email, requester)
                ]
                consultant_replies.sort(key=lambda e: e.sent_time)
                if len(consultant_replies) >= 2:
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
                    consultant_replies = [
                        e for e in thread
                        if e.sent_time and _match_requester(e.sender_name, e.sender_email, requester)
                    ]
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
            if consultant_replies:
                ack_dt = _parse_time_str(times.response)
                res_dt = _parse_time_str(times.resolved)
                if ack_dt:
                    enable_postack_fallback_30m = os.getenv("POSTACK_FALLBACK_30M", "1") == "1"
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
                "requester": requester,
                "subject_norm": subject_norm,
                "date_tokens": date_tokens,
                "explicit_marker": explicit_marker,
                "date_anchor_missing": date_anchor_missing,
                "date_anchor_after": date_anchor_after,
                "stale_anchor": stale_anchor,
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
        if not created_col or not response_col or not resolved_col:
            return

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
                if not e.sent_time:
                    continue
                try:
                    sent_ist = _to_ist(e.sent_time)
                except Exception:
                    continue
                if window_start <= sent_ist <= window_end:
                    window_thread.append(e)

            consultant_in_window = [
                e for e in window_thread
                if _match_requester(e.sender_name, e.sender_email, requester)
            ]
            candidate_in_window = consultant_in_window

            ess_only_no_request = "ESS-only; no non-ESS request" in debug_notes
            if not candidate_in_window and ess_only_no_request and not row_has_ids:
                candidate_in_window = [
                    e for e in window_thread
                    if _is_ess_sender(e, ess_team)
                ]

            if not candidate_in_window:
                continue

            pick = min(
                candidate_in_window,
                key=lambda e: abs(_to_ist(e.sent_time) - expected_center),
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
            enable_postack_fallback_30m = os.getenv("POSTACK_FALLBACK_30M", "1") == "1"
            anchor_date = _anchor_date(date_tokens)
            if anchor_date:
                anchor_start = _to_ist(datetime(anchor_date.year, anchor_date.month, anchor_date.day))
                window_end = max(window_end, anchor_start + timedelta(hours=36))
            consultant_after = [
                e for e in consultant_replies
                if _to_ist(e.sent_time) > ack_ist and _to_ist(e.sent_time) <= window_end
                and not _is_ack_like_reply(e)
            ]
            # Augment with global scan across all emails to catch the next
            # consultant reply even if the thread grouping missed it.
            subj_norm = state.get("subject_norm") or ""
            subj_tokens = _match_tokens(subj_norm)
            subj_inc_set = _inc_tokens(subj_norm)
            min_score = 0.72
            if explicit_marker and anchor_date:
                min_score = 0.60
            if len(subj_tokens) >= 3 and len(subj_norm) >= 10:
                global_candidates = []
                for e in emails:
                    if not _match_requester(e.sender_name, e.sender_email, requester):
                        continue
                    if not e.sent_time:
                        continue
                    if _is_ack_like_reply(e):
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
                    sent_ist = _to_ist(e.sent_time)
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
                    if _to_ist(e.sent_time).date() == anchor_date
                ]
                if on_anchor:
                    consultant_after = on_anchor

            consultant_after.sort(key=lambda e: e.sent_time)
            fallback_30m_used = False

            if not consultant_after and enable_postack_fallback_30m:
                fallback_after = [
                    e for e in consultant_replies
                    if _to_ist(e.sent_time) >= (ack_ist + timedelta(minutes=30))
                    and _to_ist(e.sent_time) <= window_end
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
