#!/usr/bin/env python3
import argparse
import csv
import glob
import re
from email.utils import parsedate_to_datetime


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def requester_match(requester: str, from_line: str) -> bool:
    r = norm(requester)
    f = norm(from_line)
    if not r or not f:
        return False
    if r in f:
        return True
    toks = [t for t in re.split(r"[^a-z0-9]+", r) if len(t) >= 3]
    return all(t in f for t in toks) if toks else False


def parse_headers(path: str):
    try:
        head = open(path, "rb").read(20000).decode("utf-8", "ignore")
    except Exception:
        return None
    subj = ""
    frm = ""
    dline = ""
    for line in head.splitlines():
        if line.lower().startswith("subject:") and not subj:
            subj = line.split(":", 1)[1].strip()
        elif line.lower().startswith("from:") and not frm:
            frm = line.split(":", 1)[1].strip()
        elif (line.lower().startswith("date:") or line.lower().startswith("sent:")) and not dline:
            dline = line.split(":", 1)[1].strip()
        if subj and frm and dline:
            break
    dt = None
    if dline:
        try:
            dt = parsedate_to_datetime(dline)
        except Exception:
            dt = None
    return subj, frm, dt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--eml-root", required=True)
    ap.add_argument("--debug-csv", required=True)
    ap.add_argument("--pattern", default="SF005")
    args = ap.parse_args()

    pattern = args.pattern.lower()

    rows = []
    with open(args.debug_csv, newline="", encoding="utf-8", errors="ignore") as f:
        r = csv.reader(f)
        for row in r:
            if not row:
                continue
            desc = row[0].strip()
            if desc.lower().startswith("description"):
                continue
            subj = row[2].strip() if len(row) > 2 else ""
            if pattern in desc.lower() or pattern in subj.lower():
                requester = row[1].strip() if len(row) > 1 else ""
                rows.append((desc, requester, subj))

    # Pre-scan EMLs for the subject pattern and cache headers.
    eml_files = glob.glob(f"{args.eml_root}/**/*.eml", recursive=True)
    eml_headers = []
    for f in eml_files:
        hdr = parse_headers(f)
        if not hdr:
            continue
        subj, frm, dt = hdr
        if not subj or not frm or not dt:
            continue
        if pattern not in subj.lower():
            continue
        eml_headers.append((f, subj, frm, dt))

    # Group rows by (subject, requester) and apply occurrence logic.
    groups = {}
    for desc, requester, subj in rows:
        key = (norm(subj) or norm(desc), norm(requester))
        groups.setdefault(key, []).append((desc, requester, subj))

    for key, group_rows in groups.items():
        base_subj, requester = key
        candidates = []
        for f, subj, frm, dt in eml_headers:
            if requester_match(requester, frm):
                candidates.append((dt, f, frm, subj))
        candidates.sort(key=lambda x: x[0])
        total_rows = len(group_rows)
        print(f"\n== {group_rows[0][2] or group_rows[0][0]} | requester={requester} rows={total_rows} ==")
        if not candidates:
            print("NO CANDIDATES")
            continue
        for i, row in enumerate(group_rows):
            if total_rows <= 1:
                pick = candidates[-1]
            else:
                pick = candidates[min(i, len(candidates) - 1)]
            dt = pick[0]
            print(f"row#{i+1} pick={dt} from={pick[2]}")
            print(f"  would_fill: Created={dt} Response={dt} Resolved={dt}")


if __name__ == "__main__":
    main()
