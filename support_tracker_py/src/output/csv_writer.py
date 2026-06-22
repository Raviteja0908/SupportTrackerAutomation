import csv
from pathlib import Path


def write_csv(path: Path, rows, headers):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        import logging
        logging.warning(f"write_csv called with empty rows list for {path}")
        return
    try:
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                writer.writerow(row)
    except Exception as exc:
        import logging
        logging.error(f"Failed to write CSV to {path}: {exc}")
        raise
