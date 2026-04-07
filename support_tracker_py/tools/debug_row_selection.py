import argparse
import csv
import json
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path

from src.rules.subject_normalizer import (
    extract_subject_from_description,
    normalize_subject,
    normalize_subject_for_match,
)
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


def _load_csv_rows(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = []
        for idx, row in enumerate(reader, start=2):
            row["_line"] = idx
            rows.append(row)
        return rows


def _get_col(row: dict, *names: str) -> str:
    for name in names:
        if name in row and row[name] not in (None, ""):
            return str(row[name]).strip()
    lowered = {str(k).strip().lower(): v for k, v in row.items()}
    for name in names:
        value = lowered.get(name.lower())
        if value not in (None, ""):
            return str(value).strip()
    return ""


def _match_rows(rows: list[dict], needles: list[str]) -> list[dict]:
    out = []
    for row in rows:
        desc = _get_col(row, "Description")
        if any(n.lower() in desc.lower() for n in needles):
            out.append(row)
    return out


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
        with path.open("rb") as handle:
            msg = BytesParser(policy=policy.default).parse(handle)
        plain, html = _extract_body(msg)
        sender_name = ""
        sender_email = ""
        try:
            addrs = getaddresses([msg.get("from", "")])
            if addrs:
                sender_name, sender_email = addrs[0]
        except Exception:
            pass
        sent_time = None
        try:
            sent_time = parsedate_to_datetime(msg.get("date", ""))
        except Exception:
            sent_time = None
        return DebugEmail(
            subject=str(msg.get("subject", "") or ""),
            sender_name=sender_name or "",
            sender_email=(sender_email or "").lower(),
            sent_time=sent_time,
            body=plain,
            body_html=html,
            path=path,
        )
    except Exception:
        return None


def _fmt(dt: datetime | None) -> str:
    if not dt:
        return "-"
    try:
        dt = _to_ist(dt)
    except Exception:
        pass
    return dt.strftime("%d-%m-%Y %H:%M")


def _parse_cell_dt(value: str) -> datetime | None:
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


def _token_overlap(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    if inter <= 0:
        return 0.0
    return inter / max(len(a), len(b))


def _match_tokens(text: str) -> set[str]:
    text = normalize_subject_for_match(text or "")
    return {t for t in text.split() if t}


def _interface_tokens(text: str) -> set[str]:
    import re

    return {t.lower() for t in re.findall(r"\b[a-z]{1,5}\d{2,}\b", text or "", flags=re.IGNORECASE)}


def _inc_tokens(text: str) -> set[str]:
    import re

    return {t.lower() for t in re.findall(r"\binc\d{6,}\b", text or "", flags=re.IGNORECASE)}


def _row_family_subject(row: dict) -> str:
    return normalize_subject(extract_subject_from_description(_get_col(row, "Description")))


def _row_triplet(row: dict) -> tuple[str, str, str]:
    return (
        _get_col(row, "Created Date & Time"),
        _get_col(row, "Actual Response Date & Time"),
        _get_col(row, "Actual Resolved Date & Time"),
    )


def _candidate_emails(emails: list[DebugEmail], family_subject: str) -> list[tuple[float, str, DebugEmail]]:
    family_tokens = _match_tokens(family_subject)
    family_inc = _inc_tokens(family_subject)
    family_iface = _interface_tokens(family_subject)
    out = []
    for email in emails:
        subj_norm = normalize_subject(email.subject or "")
        subj_tokens = _match_tokens(subj_norm)
        overlap = _token_overlap(family_tokens, subj_tokens)
        contains = bool(family_subject and subj_norm and (family_subject in subj_norm or subj_norm in family_subject))
        inc_hit = bool(family_inc and _inc_tokens(subj_norm) and not family_inc.isdisjoint(_inc_tokens(subj_norm)))
        iface_hit = bool(family_iface and _interface_tokens(subj_norm) and not family_iface.isdisjoint(_interface_tokens(subj_norm)))
        score = overlap
        if contains:
            score += 0.20
        if inc_hit:
            score += 0.25
        if iface_hit:
            score += 0.15
        if score >= 0.45:
            out.append((score, subj_norm, email))
    out.sort(key=lambda x: (x[0], _fmt(x[2].sent_time)), reverse=True)
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--debug-csv", required=True)
    parser.add_argument("--output-csv", required=True)
    parser.add_argument("--eml-dir", required=True)
    parser.add_argument("--subject", action="append", required=True)
    args = parser.parse_args()

    debug_rows = _load_csv_rows(Path(args.debug_csv))
    output_rows = _load_csv_rows(Path(args.output_csv))
    matched_debug = _match_rows(debug_rows, args.subject)
    matched_output = _match_rows(output_rows, args.subject)
    ess_team = []
    config_path = Path("/app/config/ess_team.json")
    try:
        ess_team = [
            str(v).strip().lower()
            for v in json.loads(config_path.read_text(encoding="utf-8"))
            if str(v).strip()
        ]
    except Exception:
        ess_team = []

    emails = []
    for path in Path(args.eml_dir).rglob("*.eml"):
        msg = _parse_eml(path)
        if msg:
            emails.append(msg)

    output_by_line = {int(r["_line"]): r for r in matched_output}

    for row in matched_debug:
        line_no = int(row["_line"])
        requester = _get_col(row, "Requester")
        family_subject = _row_family_subject(row)
        debug_triplet = _row_triplet(row)
        output_triplet = _row_triplet(output_by_line.get(line_no, {}))

        print("=" * 100)
        print(f"line={line_no} requester={requester}")
        print(f"desc={_get_col(row, 'Description')}")
        print(f"family_subject={family_subject}")
        print(f"debug_triplet={' / '.join(v or '-' for v in debug_triplet)}")
        print(f"output_triplet={' / '.join(v or '-' for v in output_triplet)}")
        print(f"notes={_get_col(row, 'Notes')}")

        response_dt = _parse_cell_dt(output_triplet[1])
        resolved_dt = _parse_cell_dt(output_triplet[2])

        candidates = _candidate_emails(emails, family_subject)
        family_buckets: dict[str, list[tuple[float, DebugEmail]]] = defaultdict(list)
        for score, subj_norm, email in candidates:
            family_buckets[subj_norm].append((score, email))

        print("-" * 100)
        print("TOP SUBJECT FAMILIES")
        ranked_families = sorted(
            family_buckets.items(),
            key=lambda kv: (
                max(s for s, _ in kv[1]),
                len(kv[1]),
                kv[0],
            ),
            reverse=True,
        )
        for idx, (subj_norm, items) in enumerate(ranked_families[:8], start=1):
            consultant_hits = sum(
                1
                for _, email in items
                if requester and _match_requester(email.sender_name, email.sender_email, requester)
            )
            print(f"  {idx}. count={len(items)} consultant_hits={consultant_hits} subject={subj_norm}")

        print("-" * 100)
        print("MATCHED EMAILS")
        for idx, (score, subj_norm, email) in enumerate(candidates[:30], start=1):
            try:
                kind = _classify_reply_kind(email)
            except Exception:
                kind = "unknown"
            ess = _is_ess_sender(email, ess_team)
            req_hit = requester and _match_requester(email.sender_name, email.sender_email, requester)
            flags = []
            if req_hit:
                flags.append("requester")
            if ess:
                flags.append("ess")
            if response_dt and email.sent_time:
                try:
                    if _to_ist(email.sent_time).strftime("%d-%m-%Y %H:%M") == response_dt.strftime("%d-%m-%Y %H:%M"):
                        flags.append("matches_output_response")
                except Exception:
                    pass
            if resolved_dt and email.sent_time:
                try:
                    if _to_ist(email.sent_time).strftime("%d-%m-%Y %H:%M") == resolved_dt.strftime("%d-%m-%Y %H:%M"):
                        flags.append("matches_output_resolved")
                except Exception:
                    pass
            print(
                f"  {idx}. {_fmt(email.sent_time)} | score={score:.2f} | kind={kind} | {';'.join(flags) or '-'}"
            )
            print(f"     from={email.sender_email or email.sender_name}")
            print(f"     subj={subj_norm}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
