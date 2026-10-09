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

# Month abbreviations in English for GCP-standard header date formatting
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
    # Replace illegal characters: \ / : * ? " < > |
    sanitized = re.sub(r'[\\/*?:"<>|]', '_', name)
    # Collapse multiple spaces, trim whitespace
    sanitized = re.sub(r'[ \t]+', ' ', sanitized).strip()
    return sanitized


def generate_output_filename(
    study_code: str,
    qc_form_name: str,
    design_version: str,
    gen_date: Optional[datetime] = None,
) -> str:
    """Generates output filename using schema:

    STUDYNAME_YYYYMMDD_name of qc form_design version of ecrf.xlsx
    Example: SurfSubQ ML45551_20261005_Validation Check QC Form_20.0.xlsx
    """
    if gen_date is None:
        gen_date = datetime.now()
    date_str = gen_date.strftime("%Y%m%d")
    raw_name = f"{study_code}_{date_str}_{qc_form_name}_{design_version}.xlsx"
    return sanitize_filename(raw_name)


def extract_metadata(wb: openpyxl.Workbook) -> StudyMetadata:
    """Extracts study metadata from 'Design settings' and 'Design info' sheets."""
    meta = StudyMetadata()

    # 1. Parse 'Design settings'
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

    # 2. Parse 'Design info'
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


def parse_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, cfg: SheetConfig) -> ParsedSheet:
    """Extracts the lowest header row and all subsequent data rows from a sheet."""
    # Find lowest header row
    header_row_idx = cfg.lowest_header_row
    max_col = ws.max_column

    # Extract headers and determine valid column span (stop at trailing empty headers)
    raw_headers: List[str] = []
    for c in range(1, max_col + 1):
        val = ws.cell(header_row_idx, c).value
        raw_headers.append(str(val).strip() if val is not None else "")

    # Trim trailing empty header strings
    while raw_headers and not raw_headers[-1]:
        raw_headers.pop()

    num_cols = len(raw_headers)
    headers = raw_headers

    # Extract data rows starting from cfg.data_start_row
    data_rows: List[List[Any]] = []
    for r in range(cfg.data_start_row, ws.max_row + 1):
        row_vals: List[Any] = []
        is_empty = True
        for c in range(1, num_cols + 1):
            cell_val = ws.cell(r, c).value
            if cell_val is not None and str(cell_val).strip() != "":
                is_empty = False
            row_vals.append(cell_val)

        # Do not include completely empty rows
        if not is_empty:
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
                # Sheet might be absent in some mini exports; record warning/empty
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
    """Replaces <Code>, <X.X>, and <DD/MMM/YYYY> placeholders in Excel headers/footers

    and strips legacy color codes (e.g. &KFF0000) so populated headers use clean black text.
    """
    if not raw_text:
        return ""
    text = raw_text.replace("<Code>", study_code)
    text = text.replace("<X.X>", design_version)
    text = text.replace("<DD/MMM/YYYY>", header_date)
    # Strip any color formatting codes (e.g. &KFF0000, &K01+000)
    text = re.sub(r"&K[0-9A-Fa-f\+\-]+", "", text)
    return text
