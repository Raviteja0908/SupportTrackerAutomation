import json
import os
from pathlib import Path


def load_json_list(path: Path):
    if not path.exists():
        return set()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return {str(x).strip().lower() for x in data if str(x).strip()}
    except Exception:
        return set()


def load_subject_exclusions(path: Path):
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return [str(x).strip().lower() for x in data if str(x).strip()]
    except Exception:
        return []


def load_aspose_license(logger):
    """Aspose license not used in readpst + EML pipeline."""
    logger.log("[INFO] Aspose license not used (readpst + EML pipeline).")
