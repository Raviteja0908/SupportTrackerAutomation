import argparse
import csv
from pathlib import Path

from src.rules.subject_normalizer import normalize_subject

from debug_unique_quote_seed import (
    _analyze_unique_shape,
    _load_ess_team,
    _parse_eml,
    _quoted_from_line_is_ess,
)


_SHORT_VARIANT_STOPWORDS = {
    "re", "fw", "fwd", "aw", "wg", "sv",
    "in", "on", "at", "to", "of", "for", "not",
    "and", "or", "the", "a", "an", "is", "it",
    "by", "as", "if", "no", "we", "us",
    "es", "api", "sap", "uat", "fct",
}


def _match_tokens(text: str) -> set[str]:
    import re

    if not text:
        return set()
    return {tok for tok in re.sub(r"[^a-z0-9]+", " ", text.lower()).split() if tok}


def _subject_short_variant_tokens(text: str) -> set[str]:
    norm = normalize_subject(text or "")
    return {
        tok
        for tok in _match_tokens(norm)
        if len(tok) <= 3 and tok.isalpha() and tok not in _SHORT_VARIANT_STOPWORDS
    }


def _load_csv_rows(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _matching_rows(rows: list[dict], subject: str) -> list[tuple[int, dict]]:
    wanted = normalize_subject(subject)
    out = []
    for idx, row in enumerate(rows, start=2):
        desc = row.get("Description") or ""
        if normalize_subject(desc) == wanted:
            out.append((idx, row))
    return out


def _scan_matching_emls(eml_dir: Path, wanted_norms: set[str]) -> dict[str, list]:
    buckets = {norm: [] for norm in wanted_norms}
    for path in eml_dir.rglob("*.eml"):
        email_obj = _parse_eml(path)
        if not email_obj:
            continue
        norm = normalize_subject(email_obj.subject)
        if norm in buckets:
            buckets[norm].append(email_obj)
    for matches in buckets.values():
        matches.sort(key=lambda e: e.sent_time.isoformat() if e.sent_time else "")
    return buckets


def main():
    parser = argparse.ArgumentParser(description="Bundle current row/debug/seed traces for multiple subjects.")
    parser.add_argument("--output-csv", required=True)
    parser.add_argument("--debug-csv", required=True)
    parser.add_argument("--eml-dir", required=True)
    parser.add_argument("--subject", action="append", required=True)
    parser.add_argument("--config-dir", default="/app/config")
    args = parser.parse_args()

    output_rows = _load_csv_rows(Path(args.output_csv))
    debug_rows = _load_csv_rows(Path(args.debug_csv))
    eml_dir = Path(args.eml_dir)
    ess_team = _load_ess_team(Path(args.config_dir))
    wanted_norms = {normalize_subject(subject) for subject in args.subject}
    eml_matches = _scan_matching_emls(eml_dir, wanted_norms)

    for subject in args.subject:
        print("=" * 100)
        print(f"subject={subject}")
        print(f"norm={normalize_subject(subject)}")
        print(f"short_variants={sorted(_subject_short_variant_tokens(subject))}")

        rows = _matching_rows(output_rows, subject)
        if not rows:
            print("output_rows=<none>")
        else:
            print("output_rows:")
            for line_no, row in rows:
                dbg = debug_rows[line_no - 2] if (line_no - 2) < len(debug_rows) else {}
                print(
                    f"  line={line_no} | requester={row.get('Requester')} | "
                    f"created={row.get('Created Date & Time')} | "
                    f"ack={row.get('Actual Response Date & Time')} | "
                    f"resolved={row.get('Actual Resolved Date & Time')}"
                )
                print(
                    "    sources="
                    f"{dbg.get('CreatedSource')} | {dbg.get('AckSource')} | {dbg.get('ResolvedSource')}"
                )
                print(f"    notes={dbg.get('Notes')}")

        matches = eml_matches.get(normalize_subject(subject), [])
        if not matches:
            print("matching_emls=<none>")
            continue

        print(f"matching_emls={len(matches)}")
        for email_obj in matches:
            analysis = _analyze_unique_shape(email_obj, ess_team)
            flags = analysis["flags"]
            print("-" * 100)
            print(f"path={email_obj.path}")
            print(f"sent={email_obj.sent_time} | from={email_obj.sender_name}")
            print(
                "reply_flags="
                f"kind={flags['kind']} | direct={flags['direct_resolution']} | "
                f"ack_candidate={flags['ack_candidate']} | thanks={flags['thanks_info']} | "
                f"nonfinal={flags['nonfinal_followup']} | real_reply={flags['substantive_reply']}"
            )
            print(
                f"first_quoted_is_ess={analysis['first_is_ess']} | "
                f"lower_non_ess={analysis['lower_non_ess']} | "
                f"paired_request={analysis['paired_request']} | "
                f"predicted_shape={analysis['predicted']} | why={analysis['why']}"
            )
            print("quoted_blocks:")
            for idx, (from_line, sent_ist, subj) in enumerate(analysis["blocks"][:8], start=1):
                subj_short = sorted(_subject_short_variant_tokens(subj or ""))
                overlap = bool(set(subj_short) & _subject_short_variant_tokens(subject))
                print(
                    f"  {idx}. ess={_quoted_from_line_is_ess(from_line, ess_team)} | "
                    f"sent={sent_ist} | short_overlap={overlap} | from={from_line} | subj={subj}"
                )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
