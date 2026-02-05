from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from ..output.run_logger import MarkingReason


EXPECTED_HEADERS = [
    "Service No",
    "Environment",
    "Module",
    "Issue Type",
    "Category",
    "Priority",
    "Created Date & Time",
    "Description",
    "Status",
    "Requester",
    "Consultant",
    "Location/ Branch",
    "Actual Response Date & Time",
    "Actual Resolved Date & Time",
    "SLA Start Date & Time",
    "SLA Target End Date & Time (Response Time)",
    "SLA Target End Date & Time (Resolution Time)",
    "SLA Met (Response)? (Auto)",
    "SLA Met (Resolution)? (Auto)",
    "Time Delay (days) (Response)",
    "Time Delay (days) (Resolution)",
    "Created Date (YYMM DD) (Auto)",
    "Target Response Date (YYMM DD) (Auto)",
    "Target Resolved Date (YYMM DD) (Auto)",
    "Actual Resolved Date (YYMM DD) (Auto)",
    "Comments",
    "Category Type",
    "ServiceRequest/Incident?",
    "ServiceRequest/Incident type?",
    "Issue occurred in",
    "Interface Code",
    "Time spent in minutes",
    "Time spent in hours",
    "Time spent in seconds",
]


def _normalize_header(text: str) -> str:
    if text is None:
        return ""
    return " ".join(str(text).replace("\n", " ").split()).strip().lower()


@dataclass
class FillResult:
    filled_count: int
    maintenance_count: int
    unknown_count: int


def fill_template(template_path, output_path, row_resolver, logger):
    template_path = Path(template_path)
    output_path = Path(output_path)

    wb = load_workbook(template_path)
    ws = wb.active

    header_row = _find_header_row(ws, logger)
    col_map = _build_col_map(ws, header_row)

    if "description" not in col_map:
        raise RuntimeError("Description column not found in template.")

    comments_col = col_map.get("comments")

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    filled = 0
    maintenance = 0
    unknown = 0

    for row in range(header_row + 1, ws.max_row + 1):
        description = ws.cell(row, col_map["description"]).value
        service_no = ws.cell(row, col_map.get("service no", 1)).value

        if _is_row_empty(description, service_no):
            continue

        row_context = _build_row_context(ws, row, col_map)

        desc_text = str(row_context.get("Description", "") or "")
        if "maintenance" in desc_text.lower():
            _mark_row(ws, row, red_fill)
            _write_comment(ws, row, comments_col, MarkingReason.maintenance)
            maintenance += 1
            continue

        resolved = row_resolver(row_context)
        mark_blue = resolved.pop("_MarkBlue", False)
        required_missing, reason = _is_unknown(resolved)
        if required_missing:
            _mark_row(ws, row, yellow_fill)
            _write_comment(ws, row, comments_col, reason or MarkingReason.unknown)
            unknown += 1
            continue

        if mark_blue:
            _mark_row(ws, row, blue_fill)
            _write_comment(ws, row, comments_col, MarkingReason.blue)

        _write_values(ws, row, col_map, resolved)
        filled += 1

    safe_path = _resolve_output_path(output_path)
    try:
        wb.save(safe_path)
        logger.log(f"[INFO] Excel saved: {safe_path}")
    except PermissionError:
        alt_path = _resolve_output_path(output_path, force_suffix=True)
        wb.save(alt_path)
        logger.log(f"[WARNING] Output locked, saved to: {alt_path}")

    return FillResult(filled, maintenance, unknown)


def _resolve_output_path(path: Path, force_suffix: bool = False) -> Path:
    path = Path(path)
    if force_suffix:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return path.with_name(path.stem + "_" + stamp + path.suffix)
    return path


def _find_header_row(ws, logger):
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

    if best_hits < 5:
        logger.log("[WARNING] Header row detection weak; using row 1.")
        return 1

    logger.log(f"[INFO] Header row detected at {best_row} with {best_hits} matches.")
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


def _build_row_context(ws, row, col_map):
    context = {}
    for key, col in col_map.items():
        value = ws.cell(row, col).value
        context[_title_case(key)] = value if value is not None else ""
    return context


def _title_case(key: str) -> str:
    return " ".join(p.capitalize() for p in key.split())


def _is_row_empty(description, service_no):
    if description is None and service_no is None:
        return True
    if str(description).strip() == "" and str(service_no).strip() == "":
        return True
    return False


def _write_values(ws, row, col_map, values):
    for key, value in values.items():
        col = col_map.get(_normalize_header(key))
        if not col:
            continue
        ws.cell(row, col).value = value


def _write_comment(ws, row, comments_col, text):
    if not comments_col:
        return
    ws.cell(row, comments_col).value = text


def _mark_row(ws, row, fill):
    for col in range(1, ws.max_column + 1):
        ws.cell(row, col).fill = fill


def _is_unknown(resolved_values):
    created = resolved_values.get("Created Date & Time", "")
    resolved = resolved_values.get("Actual Resolved Date & Time", "")
    sr = resolved_values.get("ServiceRequest/Incident?", "")
    sr_type = resolved_values.get("ServiceRequest/Incident type?", "")

    missing = []
    if not created:
        missing.append("Created Date & Time")
    if not resolved:
        missing.append("Actual Resolved Date & Time")
    if not sr:
        missing.append("ServiceRequest/Incident?")
    if not sr_type:
        missing.append("ServiceRequest/Incident type?")

    if missing:
        return True, "Missing: " + ", ".join(missing)

    return False, ""
