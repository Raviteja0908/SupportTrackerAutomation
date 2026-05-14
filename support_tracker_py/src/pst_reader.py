from pathlib import Path
import os
import shutil
import subprocess

from .eml_reader import read_eml_directory


def read_pst_emails(pst_path, logger, eml_root: Path):
    pst_path = Path(pst_path)
    eml_dir = eml_root / pst_path.stem
    eml_dir.mkdir(parents=True, exist_ok=True)

    if not _readpst_available():
        logger.log("[ERROR] readpst not available in container.")
        return []

    if _env_truthy("SKIP_PST_CONVERT"):
        logger.log("[INFO] SKIP_PST_CONVERT enabled; using existing EMLs only.")
        return read_eml_directory(eml_dir, logger)

    existing_any = next(eml_dir.rglob("*.eml"), None)
    if existing_any is not None:
        if _env_truthy("SKIP_PST_CONVERT"):
            logger.log(f"[INFO] Using existing EMLs in {eml_dir}")
            return read_eml_directory(eml_dir, logger)
        logger.log(f"[INFO] Removing existing EML cache before fresh PST conversion: {eml_dir}")
        shutil.rmtree(eml_dir)
        eml_dir.mkdir(parents=True, exist_ok=True)

    work_dir = eml_root / "__pst_work__"
    work_dir.mkdir(parents=True, exist_ok=True)
    local_pst = work_dir / (pst_path.stem + pst_path.suffix)
    try:
        shutil.copy2(pst_path, local_pst)
        pst_to_read = local_pst
    except Exception as exc:
        logger.log(f"[WARNING] Failed to copy PST locally: {exc}")
        logger.log("[WARNING] Falling back to reading PST directly from input mount.")
        pst_to_read = pst_path

    logger.log(f"[INFO] Converting PST to EML: {pst_to_read} -> {eml_dir}")
    extra_args = _readpst_extra_args()
    cmd = ["readpst", "-r", "-e", "-o", str(eml_dir)] + extra_args + [str(pst_to_read)]
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
    except Exception as exc:
        logger.log(f"[ERROR] readpst failed: {exc}")
        try:
            if hasattr(exc, "stderr") and exc.stderr:
                logger.log(f"[ERROR] readpst stderr: {exc.stderr.strip()}")
        except Exception:
            pass
        raise RuntimeError(f"readpst failed for {pst_path}") from exc

    emails = read_eml_directory(eml_dir, logger)
    if not emails:
        raise RuntimeError(f"readpst produced zero emails for {pst_path}")
    return emails


def _readpst_available() -> bool:
    return shutil.which("readpst") is not None


def _env_truthy(name: str) -> bool:
    val = (os.getenv(name) or "").strip().lower()
    return val in ("1", "true", "yes", "y", "on")


def _readpst_extra_args():
    # Allow tuning readpst without code changes
    raw = (os.getenv("READPST_ARGS") or "").strip()
    if not raw:
        return []
    return raw.split()
