"""Parser for Viedoc EDC Configuration Reports.

Extracts study metadata, version information, and all target specification tables
using the lowest header row in each sheet.
"""

from datetime import datetime
from pathlib import Path
import re
from typing import Any, List, Optional
import openpyxl

from src.models import (
    ParsedSheet,
    ParsedViedocConfig,
    SHEET_CONFIGS,
    SheetConfig,
    StudyMetadata,
)

MONTH_NAMES = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
]


def format_header_date(dt: Optional[datetime] = None) -> str:
    """Formats date in standard DD/MMM/YYYY uppercase format (e.g., 05/OCT/2026)."""
    if dt is None:
        dt = datetime.now()
    day = f"{dt.day:02d}"
    month = MONTH_NAMES[dt.month - 1]
    year = f"{dt.year:04d}"
    return f"{day}/{month}/{year}"


def sanitize_filename(name: str) -> str:
    """Removes characters forbidden in Windows filenames while preserving valid delimiters."""
    sanitized = re.sub(r'[\\/*?:"<>|]', "_", name)
    sanitized = re.sub(r"[ \t]+", " ", sanitized).strip()
    return sanitized


def generate_output_filename(
    study_code: str,
    qc_form_name: str,
    design_version: str,
    gen_date: Optional[datetime] = None,
) -> str:
    """Generates output filename using schema:
    STUDYNAME_YYYYMMDD_name of qc form_design version of ecrf.xlsx
    """
    if gen_date is None:
        gen_date = datetime.now()
    date_str = gen_date.strftime("%Y%m%d")
    raw_name = f"{study_code}_{date_str}_{qc_form_name}_{design_version}.xlsx"
    return sanitize_filename(raw_name)


def extract_metadata(wb: openpyxl.Workbook) -> StudyMetadata:
    """Extracts study metadata from 'Design settings' and 'Design info' sheets."""
    meta = StudyMetadata()

    if "Design settings" in wb.sheetnames:
        ws = wb["Design settings"]
        for r in range(1, ws.max_row + 1):
            key_cell = ws.cell(r, 1).value
            val_cell = ws.cell(r, 2).value
            if key_cell is None or val_cell is None:
                continue
            key = str(key_cell).strip().lower()
            val = str(val_cell).strip()

            if key == "internal description":
                meta.internal_description = val
            elif key == "study name":
                meta.study_name = val
            elif key == "study description":
                meta.study_description = val
            elif key == "protocol name":
                meta.protocol_name = val
            elif key == "protocol version":
                meta.protocol_version = val
            elif key == "version":
                meta.version = val
            elif key == "revised version":
                meta.revised_version = val

    if "Design info" in wb.sheetnames:
        ws = wb["Design info"]
        for r in range(1, ws.max_row + 1):
            key_cell = ws.cell(r, 1).value
            val_cell = ws.cell(r, 2).value
            if key_cell is None or val_cell is None:
                continue
            key = str(key_cell).strip().lower()
            val = str(val_cell).strip()

            if "printed date/time" in key:
                meta.printed_date_time = val
            elif "systemversion" in key:
                meta.system_version = val

    return meta


def is_help_or_info_row(row_vals: List[Any]) -> bool:
    """Detects whether a row contains informational/help text rather than study specification data."""
    non_empty = [v for v in row_vals if v is not None and str(v).strip() != ""]
    if not non_empty:
        return False

    if len(non_empty) <= 2:
        first_str = str(non_empty[0]).strip().lower()
        help_prefixes = (
            "see the ",
            "see also ",
            "note!",
            "note:",
            "min/max repeat",
            "repeating number",
            'the "',
            "the '",
            "the additional languages",
            "if recurrence",
            "please note",
        )
        if any(first_str.startswith(p) for p in help_prefixes):
            return True
        if (
            "sheet for details" in first_str
            or "for detailed" in first_str
            or "for details about" in first_str
        ):
            return True

    return False


def is_active_alert_row(headers: List[str], row_vals: List[Any]) -> bool:
    """Filters Viedoc Alert rows per SOI Section 4.2.B:
    Only alert rows where 'Send message when TRUE = X' or 'Send message when FALSE = X'
    (or 'Send message when TRUE/FALSE' is active) require QC.
    """
    header_map = {h.strip().lower(): idx for idx, h in enumerate(headers) if h}

    # Check Action type (exclude Tracker / Repeating if not sending message)
    send_cols = [
        idx
        for h_low, idx in header_map.items()
        if "send message when" in h_low
    ]
    if not send_cols:
        return True

    for idx in send_cols:
        if idx < len(row_vals) and row_vals[idx] is not None:
            val = str(row_vals[idx]).strip().upper()
            if val in ("X", "TRUE", "FALSE", "YES", "1"):
                return True
    return False


def parse_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, cfg: SheetConfig) -> ParsedSheet:
    """Extracts the lowest header row and all subsequent data rows from a sheet."""
    header_row_idx = cfg.lowest_header_row
    max_col = ws.max_column

    raw_headers: List[str] = []
    for c in range(1, max_col + 1):
        val = ws.cell(header_row_idx, c).value
        raw_headers.append(str(val).strip() if val is not None else "")

    while raw_headers and not raw_headers[-1]:
        raw_headers.pop()

    num_cols = len(raw_headers)
    headers = raw_headers

    data_rows: List[List[Any]] = []
    for r in range(cfg.data_start_row, ws.max_row + 1):
        row_vals: List[Any] = []
        is_empty = True
        for c in range(1, num_cols + 1):
            cell_val = ws.cell(r, c).value
            if cell_val is not None and str(cell_val).strip() != "":
                is_empty = False
            row_vals.append(cell_val)

        if not is_empty and not is_help_or_info_row(row_vals):
            if cfg.qc_sheet_name == "Viedoc - Alerts" and not is_active_alert_row(headers, row_vals):
                continue
            data_rows.append(row_vals)

    return ParsedSheet(
        qc_sheet_name=cfg.qc_sheet_name,
        viedoc_sheet_name=cfg.viedoc_sheet_name,
        lowest_headers=headers,
        rows=data_rows,
    )


def parse_viedoc_report(file_path: str | Path) -> ParsedViedocConfig:
    """Ingests a Viedoc Configuration Report Excel workbook and extracts
    metadata and all configured specification sheets.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Viedoc Configuration Report not found: {path}")

    wb = openpyxl.load_workbook(str(path), data_only=True)
    try:
        metadata = extract_metadata(wb)
        sheets = {}

        for cfg in SHEET_CONFIGS:
            if cfg.viedoc_sheet_name in wb.sheetnames:
                ws = wb[cfg.viedoc_sheet_name]
                parsed = parse_sheet(ws, cfg)
                sheets[cfg.qc_sheet_name] = parsed
            else:
                sheets[cfg.qc_sheet_name] = ParsedSheet(
                    qc_sheet_name=cfg.qc_sheet_name,
                    viedoc_sheet_name=cfg.viedoc_sheet_name,
                    lowest_headers=[],
                    rows=[],
                )

        return ParsedViedocConfig(
            file_path=path,
            metadata=metadata,
            sheets=sheets,
        )
    finally:
        wb.close()


def update_header_footer_placeholders(
    raw_text: str,
    study_code: str,
    design_version: str,
    header_date: str,
) -> str:
    """Replaces <Code>, <X.X>, and <DD/MMM/YYYY> placeholders in Excel headers/footers."""
    if not raw_text:
        return ""
    text = raw_text.replace("<Code>", study_code)
    text = text.replace("<X.X>", design_version)
    text = text.replace("<DD/MMM/YYYY>", header_date)
    text = re.sub(r"&K[0-9A-Fa-f\+\-]+", "", text)
    return text