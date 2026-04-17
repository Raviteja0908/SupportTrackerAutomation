import argparse
import csv
import re
import subprocess
import sys
from pathlib import Path

from src.rules.subject_normalizer import normalize_subject


def _load_rows(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _find_matching_rows(rows: list[dict], debug_rows: list[dict], description: str):
    wanted = normalize_subject(description)
    matches = []
    for line_no, row in enumerate(rows, start=2):
        desc = row.get("Description") or ""
        if normalize_subject(desc) == wanted:
            dbg = debug_rows[line_no - 2] if (line_no - 2) < len(debug_rows) else {}
            matches.append((line_no, row, dbg))
    return matches


def _candidate_subjects(description: str) -> list[str]:
    desc = (description or "").strip()
    candidates = []

    def add(value: str):
        value = (value or "").strip()
        if not value:
            return
        if value not in candidates:
            candidates.append(value)

    add(desc)

    for marker in ("RE:", "FW:", "FWD:"):
        match = re.search(rf"(?i)\b{re.escape(marker)}\s*.+", desc)
        if match:
            add(match.group(0).strip())

    arrow_parts = [part.strip() for part in re.split(r"\s*(?:--?>|=>)\s*", desc) if part.strip()]
    for part in arrow_parts:
        if re.fullmatch(r"\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?", part):
            continue
        add(part)

    if arrow_parts:
        tail = arrow_parts[-1]
        if not re.fullmatch(r"\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?", tail):
            add(tail)

    return candidates


def _run_subtool(args: list[str]) -> str:
    proc = subprocess.run(args, capture_output=True, text=True)
    if proc.returncode == 0:
        return proc.stdout.strip()
    return (proc.stdout + "\n" + proc.stderr).strip()


def _run_seed_trace(eml_dir: Path, config_dir: Path, description: str) -> str:
    tool = Path(__file__).with_name("debug_unique_quote_seed.py")
    for subject in _candidate_subjects(description):
        out = _run_subtool(
            [
                sys.executable,
                str(tool),
                "--eml-dir",
                str(eml_dir),
                "--subject",
                subject,
                "--config-dir",
                str(config_dir),
            ]
        )
        if out and "No matching EMLs found." not in out:
            return f"seed_subject={subject}\n{out}"
    return "seed_trace=No matching EMLs found."


def _run_occurrence_trace(output_csv: Path, debug_csv: Path, eml_dir: Path, description: str) -> str:
    tool = Path(__file__).with_name("debug_occurrence_family.py")
    for subject in _candidate_subjects(description):
        out = _run_subtool(
            [
                sys.executable,
                str(tool),
                "--debug-csv",
                str(debug_csv),
                "--output-csv",
                str(output_csv),
                "--eml-dir",
                str(eml_dir),
                "--subject",
                subject,
            ]
        )
        if out and "family_subject_norm=" in out:
            return f"occ_subject={subject}\n{out}"
    return "occurrence_trace=No matching occurrence trace found."


def main():
    parser = argparse.ArgumentParser(description="Focused selection/occurrence diagnostics for exact failing cases.")
    parser.add_argument("--output-csv", required=True)
    parser.add_argument("--debug-csv", required=True)
    parser.add_argument("--eml-dir", required=True)
    parser.add_argument("--config-dir", default="/app/config")
    parser.add_argument(
        "--case",
        action="append",
        required=True,
        help="Format: description|expected_created|expected_ack|expected_resolved",
    )
    args = parser.parse_args()

    output_rows = _load_rows(Path(args.output_csv))
    debug_rows = _load_rows(Path(args.debug_csv))
    eml_dir = Path(args.eml_dir)
    config_dir = Path(args.config_dir)

    for raw_case in args.case:
        parts = raw_case.split("|")
        description = parts[0].strip()
        expected_created = parts[1].strip() if len(parts) > 1 else ""
        expected_ack = parts[2].strip() if len(parts) > 2 else ""
        expected_resolved = parts[3].strip() if len(parts) > 3 else ""

        print("=" * 120)
        print(f"case={description}")
        print(
            f"expected={expected_created or '-'} / "
            f"{expected_ack or '-'} / {expected_resolved or '-'}"
        )

        matches = _find_matching_rows(output_rows, debug_rows, description)
        if not matches:
            print("row=<not found>")
        for line_no, row, dbg in matches:
            actual = (
                row.get("Created Date & Time") or "",
                row.get("Actual Response Date & Time") or "",
                row.get("Actual Resolved Date & Time") or "",
            )
            print(
                f"line={line_no} | requester={row.get('Requester')} | "
                f"actual={actual[0] or '-'} / {actual[1] or '-'} / {actual[2] or '-'}"
            )
            print(
                "sources="
                f"{dbg.get('CreatedSource') or '-'} | "
                f"{dbg.get('AckSource') or '-'} | "
                f"{dbg.get('ResolvedSource') or '-'}"
            )
            print(f"notes={dbg.get('Notes') or '-'}")
            if expected_created or expected_ack or expected_resolved:
                print(
                    "compare="
                    f"created={'OK' if actual[0] == expected_created else 'DIFF'} | "
                    f"ack={'OK' if actual[1] == expected_ack else 'DIFF'} | "
                    f"resolved={'OK' if actual[2] == expected_resolved else 'DIFF'}"
                )

        print("-" * 120)
        print("SEED TRACE")
        print(_run_seed_trace(eml_dir, config_dir, description))
        print("-" * 120)
        print("OCCURRENCE TRACE")
        print(_run_occurrence_trace(Path(args.output_csv), Path(args.debug_csv), eml_dir, description))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
