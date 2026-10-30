"""Configuration Report Diff Engine for Viedoc EDC.

Compares two Viedoc Configuration Reports (baseline vs new version) across all sheets,
detects added, modified, and removed elements across study settings and specifications,
highlights changes in-place with yellow fills, rich text diffs, and comments,
inserts an initial 'Change Log' sheet detailing all changes, sets yellow tab colors
for modified sheets, enforces normal view, and excludes QC columns and headers/footers.
"""

from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.delta_engine import (
    DATA_ALIGNMENT,
    DATA_BORDER,
    DATA_FONT,
    DIFF_NEW_FONT,
    DIFF_OLD_FONT,
    HEADER_FILL,
    HEADER_FONT,
    REMOVED_FILL,
    REMOVED_FONT,
    YELLOW_FILL,
    YELLOW_TAB_COLOR,
    clear_sheet_header_footer,
    normalize_value,
)
from src.generator import adjust_column_widths, set_cell_value_safe, set_sheet_view_normal
from src.parser import extract_metadata, sanitize_filename


CONFIG_REPORT_SHEET_DEFS: Dict[str, Dict[str, Any]] = {
    "Design info": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Items"],
        "ignore_keys": {"printed date/time"},
    },
    "Design settings": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Items"],
    },
    "Global-Designer settings": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Field defaults"],
    },
    "Global-Medical coding": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Form", "Item"],
    },
    "Global-Data mappings": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Data Mapping Name"],
    },
    "Global-Reference data scopes": {
        "header_row": 3,
        "data_start": 4,
        "keys": ["Scope name", "Factor #", "Variable #", "Variable form", "Variable name", "Variable Type.Name"],
    },
    "Global-Reports configuration": {
        "multi_section": True,
    },
    "Settings-Selection View": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["No"],
    },
    "Settings-Subject Id Gen": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Subject ID format"],
    },
    "Settings-SDV": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Items"],
    },
    "Settings-Miscellaneous": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Items"],
    },
    "Settings-Alerts": {
        "header_row": 3,
        "data_start": 4,
        "keys": ["Internal description of alert", "Action type"],
    },
    "Settings-Subject Status": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Status"],
    },
    "Settings-RTSM": {
        "header_row": 1,
        "data_start": 2,
        "keys": [],
    },
    "Settings-eLearning": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Title"],
    },
    "Roles": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Role ID"],
    },
    "Forms": {
        "header_row": 3,
        "data_start": 4,
        "keys": ["Id"],
    },
    "Items and Groups": {
        "header_row": 3,
        "data_start": 4,
        "keys": ["Form ID", "Item group ID", "Item ID"],
    },
    "Study workflow-Events": {
        "header_row": 3,
        "data_start": 4,
        "keys": ["Study event ID"],
    },
    "Study workflow-Activities": {
        "header_row": 3,
        "data_start": 4,
        "keys": ["Study event ID", "Activity ID"],
    },
    "Study workflow-Forms": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Event ID", "Activity ID", "Form ID"],
    },
    "Event reminder": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Event ID"],
    },
    "Viedoc Me reminder": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Event ID", "Activity ID"],
    },
    "Functions and Conditions": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["Event ID", "Activity ID", "Form ID", "Item group ID", "Item ID", "Type"],
    },
    "Data checks": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["OID"],
    },
    "Code lists": {
        "header_row": 2,
        "data_start": 3,
        "keys": ["OID", "Code value"],
    },
}


def detect_sheet_structure(
    ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str
) -> Tuple[int, int, List[str], List[str], Set[str]]:
    """Determines header row, data start row, header list, key columns, and ignore keys."""
    if sheet_name in CONFIG_REPORT_SHEET_DEFS and not CONFIG_REPORT_SHEET_DEFS[sheet_name].get("multi_section"):
        cfg = CONFIG_REPORT_SHEET_DEFS[sheet_name]
        header_row = cfg["header_row"]
        data_start = cfg["data_start"]
        key_names = cfg.get("keys", [])
        ignore_keys = {k.lower() for k in cfg.get("ignore_keys", set())}
    else:
        # Fallback: scan rows 1..4 for first row with >= 2 non-empty values
        header_row = 1
        for r in range(1, min(5, ws.max_row + 1)):
            non_empty = [
                ws.cell(r, c).value
                for c in range(1, ws.max_column + 1)
                if ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip()
            ]
            if len(non_empty) >= 2:
                header_row = r
                break
        data_start = header_row + 1
        key_names = []
        ignore_keys = set()

    headers = [
        str(ws.cell(header_row, c).value).strip() if ws.cell(header_row, c).value is not None else ""
        for c in range(1, ws.max_column + 1)
    ]
    while headers and not headers[-1]:
        headers.pop()

    if not key_names and headers:
        key_names = [headers[0]]

    return header_row, data_start, headers, key_names, ignore_keys


def generate_config_diff_filename(
    study_code: str,
    old_version: str,
    new_version: str,
    gen_date: Optional[datetime] = None,
) -> str:
    """Generates filename using schema:
    [StudyCode]_[YYYYMMDD]_Configuration Report Diff_[OldVersion]_to_[NewVersion].xlsx
    """
    if gen_date is None:
        gen_date = datetime.now()
    date_str = gen_date.strftime("%Y%m%d")
    raw_name = f"{study_code}_{date_str}_Configuration Report Diff_{old_version}_to_{new_version}.xlsx"
    return sanitize_filename(raw_name)


def parse_config_sheet_records(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    sheet_name: str,
) -> Tuple[int, int, List[str], Dict[Tuple[str, ...], Dict[str, Any]], Dict[Tuple[str, ...], List[Any]]]:
    """Extracts specification dictionary and raw row values keyed by primary key."""
    header_row, data_start, headers, key_names, ignore_keys = detect_sheet_structure(ws, sheet_name)
    header_to_idx = {h: i for i, h in enumerate(headers)}
    key_indices = [header_to_idx[k] for k in key_names if k in header_to_idx]

    records: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    full_rows: Dict[Tuple[str, ...], List[Any]] = {}

    for r in range(data_start, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
        if not any(v is not None and str(v).strip() != "" for v in row_vals):
            continue

        if key_indices:
            key_parts = [
                normalize_value(row_vals[idx]) if idx < len(row_vals) else ""
                for idx in key_indices
            ]
        else:
            key_parts = [normalize_value(row_vals[0])] if row_vals else [str(r)]

        record_key = tuple(key_parts)
        key_desc = " / ".join(p for p in key_parts if p)

        # Check if row is in ignore_keys
        if any(normalize_value(part).lower() in ignore_keys for part in key_parts):
            continue

        spec_dict = {}
        for h, idx in header_to_idx.items():
            val = row_vals[idx] if idx < len(row_vals) else None
            spec_dict[h] = val

        records[record_key] = spec_dict
        full_rows[record_key] = row_vals

    return header_row, data_start, headers, records, full_rows


KNOWN_REPORT_SECTIONS = {"visibility", "dashboard", "demographics", "adverse events", "custom reports"}


def parse_report_config_sections(
    ws: openpyxl.worksheet.worksheet.Worksheet,
) -> List[Dict[str, Any]]:
    """Parses a multi-section sheet like 'Global-Reports configuration' into distinct sections.

    Each section consists of a section title, a column header row, and data rows.
    """
    sections = []
    r = 2
    while r <= ws.max_row:
        cell_val = ws.cell(r, 1).value
        other_cells = [ws.cell(r, c).value for c in range(2, ws.max_column + 1)]
        val_str = str(cell_val).strip() if cell_val is not None else ""
        is_known = val_str.lower() in KNOWN_REPORT_SECTIONS
        is_sec_title = (
            val_str != ""
            and not any(v is not None and str(v).strip() != "" for v in other_cells)
            and (is_known or (r < ws.max_row and not val_str.isdigit() and val_str != "#"))
        )
        if is_sec_title:
            sec_name = val_str
            header_row_idx = r + 1
            headers = [
                str(ws.cell(header_row_idx, c).value).strip() if ws.cell(header_row_idx, c).value is not None else ""
                for c in range(1, ws.max_column + 1)
            ]
            while headers and not headers[-1]:
                headers.pop()

            data_rows = []
            next_r = header_row_idx + 1
            while next_r <= ws.max_row:
                next_val = ws.cell(next_r, 1).value
                next_others = [ws.cell(next_r, c).value for c in range(2, ws.max_column + 1)]
                next_str = str(next_val).strip() if next_val is not None else ""
                is_next_sec = (
                    next_str != ""
                    and not any(v is not None and str(v).strip() != "" for v in next_others)
                    and (next_str.lower() in KNOWN_REPORT_SECTIONS or (not next_str.isdigit() and next_str != "#"))
                )
                if is_next_sec:
                    break
                row_vals = [ws.cell(next_r, c).value for c in range(1, len(headers) + 1)]
                data_rows.append((next_r, row_vals))
                next_r += 1

            sections.append({
                "name": sec_name,
                "title_row": r,
                "header_row": header_row_idx,
                "headers": headers,
                "data_rows": data_rows,
            })
            r = next_r
        else:
            r += 1

    return sections


def diff_global_reports_configuration(
    ws_old: openpyxl.worksheet.worksheet.Worksheet,
    ws_out: openpyxl.worksheet.worksheet.Worksheet,
    old_version: str,
    new_version: str,
    sheet_name: str = "Global-Reports configuration",
) -> Tuple[List[Dict[str, Any]], bool]:
    """Compares multi-section 'Global-Reports configuration' sheet between baseline and new version.

    Identifies sections (Visibility, Dashboard, Demographics, Adverse Events, Custom Reports, etc.),
    compares data rows within each section, and applies in-place highlights, rich text diffs, and comments.
    """
    old_sections_list = parse_report_config_sections(ws_old)
    new_sections_list = parse_report_config_sections(ws_out)

    old_sections = {s["name"].lower(): s for s in old_sections_list}
    new_sections = {s["name"].lower(): s for s in new_sections_list}

    diff_entries: List[Dict[str, Any]] = []
    sheet_has_changes = False

    for s_new in new_sections_list:
        sec_name = s_new["name"]
        sec_key = sec_name.lower()
        headers = s_new["headers"]
        header_to_idx = {h: i for i, h in enumerate(headers)}

        if sec_key not in old_sections:
            # Entire section added in new version
            for r_new, row_vals in s_new["data_rows"]:
                has_data = any(v is not None and str(v).strip() != "" for v in row_vals)
                if has_data:
                    sheet_has_changes = True
                    for c_idx in range(1, len(headers) + 1):
                        ws_out.cell(r_new, c_idx).fill = YELLOW_FILL
                    rk = str(row_vals[0]).strip() if (row_vals and row_vals[0] is not None and str(row_vals[0]).strip() != "") else "(All Records)"
                    record_key_desc = f"{sec_name} / {rk}" if sec_key != "visibility" else f"{sec_name}"
                    diff_entries.append({
                        "sheet_name": sheet_name,
                        "change_type": "NEW",
                        "record_key": record_key_desc,
                        "changed_field": "(Entire Record)",
                        "previous_value": "N/A",
                        "current_value": f"Added in v{new_version}",
                    })
            continue

        s_old = old_sections[sec_key]
        old_headers = s_old["headers"]

        # Parse old rows for this section
        old_records: Dict[str, Tuple[Dict[str, Any], List[Any]]] = {}
        for idx, (r_old, row_vals_old) in enumerate(s_old["data_rows"]):
            if sec_key == "visibility":
                rk = "Visibility"
            else:
                rk = str(row_vals_old[0]).strip() if (row_vals_old and row_vals_old[0] is not None and str(row_vals_old[0]).strip() != "") else str(idx + 1)
            rec_dict = {h: row_vals_old[i] if i < len(row_vals_old) else None for i, h in enumerate(old_headers)}
            old_records[rk] = (rec_dict, row_vals_old)

        seen_keys: Set[str] = set()
        for idx, (r_new, row_vals_new) in enumerate(s_new["data_rows"]):
            if sec_key == "visibility":
                rk = "Visibility"
                record_key_desc = "Visibility"
            else:
                rk = str(row_vals_new[0]).strip() if (row_vals_new and row_vals_new[0] is not None and str(row_vals_new[0]).strip() != "") else str(idx + 1)
                record_key_desc = f"{sec_name} / {rk}"

            seen_keys.add(rk)

            if rk not in old_records:
                has_data = any(v is not None and str(v).strip() != "" for v in row_vals_new)
                if has_data:
                    sheet_has_changes = True
                    for c_idx in range(1, len(headers) + 1):
                        ws_out.cell(r_new, c_idx).fill = YELLOW_FILL
                    diff_entries.append({
                        "sheet_name": sheet_name,
                        "change_type": "NEW",
                        "record_key": record_key_desc,
                        "changed_field": "(Entire Record)",
                        "previous_value": "N/A",
                        "current_value": f"Added in v{new_version}",
                    })
            else:
                old_rec_dict, row_vals_old = old_records[rk]
                row_modified = False

                for h_name in headers:
                    if not h_name or h_name == "#":
                        continue
                    if h_name in old_rec_dict:
                        old_v = old_rec_dict[h_name]
                        col_idx = header_to_idx[h_name]
                        new_v = row_vals_new[col_idx] if col_idx < len(row_vals_new) else None

                        if normalize_value(old_v) != normalize_value(new_v):
                            row_modified = True
                            sheet_has_changes = True
                            cell = ws_out.cell(r_new, col_idx + 1)
                            cell.fill = YELLOW_FILL

                            clean_old = str(old_v).rstrip("\r\n") if old_v is not None else ""
                            clean_new = str(new_v).rstrip("\r\n") if new_v is not None else ""

                            if clean_old and clean_new:
                                rt = CellRichText()
                                rt.append(TextBlock(DIFF_OLD_FONT, clean_old + "\n"))
                                rt.append(TextBlock(DIFF_NEW_FONT, clean_new))
                                cell.value = rt
                            elif clean_old and not clean_new:
                                rt = CellRichText()
                                rt.append(TextBlock(DIFF_OLD_FONT, clean_old))
                                cell.value = rt
                            else:
                                set_cell_value_safe(cell, new_v)

                            old_display = str(old_v) if old_v is not None else "(empty)"
                            cell.comment = Comment(
                                f"[v{new_version} Modified]\nPrevious: {old_display}",
                                "QC Automation",
                            )

                            diff_entries.append({
                                "sheet_name": sheet_name,
                                "change_type": "MODIFIED",
                                "record_key": record_key_desc,
                                "changed_field": h_name,
                                "previous_value": str(old_v) if old_v is not None else "",
                                "current_value": str(new_v) if new_v is not None else "",
                            })

                if row_modified:
                    for c_idx in range(1, len(headers) + 1):
                        ws_out.cell(r_new, c_idx).fill = YELLOW_FILL

        # Check for removed rows in this section
        for old_rk, (old_rec_dict, row_vals_old) in old_records.items():
            if old_rk not in seen_keys:
                check_vals = row_vals_old[1:] if sec_key != "visibility" else row_vals_old
                has_data = any(v is not None and str(v).strip() != "" for v in check_vals)
                if has_data:
                    sheet_has_changes = True
                    record_key_desc = f"{sec_name} / {old_rk}" if sec_key != "visibility" else f"{sec_name}"
                    append_row_idx = ws_out.max_row + 1
                    for c_idx, val in enumerate(row_vals_old, 1):
                        cell = ws_out.cell(append_row_idx, c_idx)
                        set_cell_value_safe(cell, val)
                        cell.fill = REMOVED_FILL
                        cell.font = REMOVED_FONT
                        cell.border = DATA_BORDER
                        cell.alignment = DATA_ALIGNMENT

                    first_cell = ws_out.cell(append_row_idx, 1)
                    first_cell.comment = Comment(
                        f"[Removed in v{new_version}]\nExisted in v{old_version}",
                        "QC Automation",
                    )
                    diff_entries.append({
                        "sheet_name": sheet_name,
                        "change_type": "REMOVED",
                        "record_key": record_key_desc,
                        "changed_field": "(Entire Record)",
                        "previous_value": f"Existed in v{old_version}",
                        "current_value": "Removed",
                    })

    # Also check if an entire section was removed from ws_old
    for s_old in old_sections_list:
        sec_name = s_old["name"]
        sec_key = sec_name.lower()
        if sec_key not in new_sections:
            for idx, (r_old, row_vals_old) in enumerate(s_old["data_rows"]):
                check_vals = row_vals_old[1:] if sec_key != "visibility" else row_vals_old
                has_data = any(v is not None and str(v).strip() != "" for v in check_vals)
                if has_data:
                    sheet_has_changes = True
                    rk = str(row_vals_old[0]).strip() if (row_vals_old and row_vals_old[0] is not None and str(row_vals_old[0]).strip() != "") else str(idx + 1)
                    record_key_desc = f"{sec_name} / {rk}" if sec_key != "visibility" else f"{sec_name}"
                    append_row_idx = ws_out.max_row + 1
                    for c_idx, val in enumerate(row_vals_old, 1):
                        cell = ws_out.cell(append_row_idx, c_idx)
                        set_cell_value_safe(cell, val)
                        cell.fill = REMOVED_FILL
                        cell.font = REMOVED_FONT
                        cell.border = DATA_BORDER
                        cell.alignment = DATA_ALIGNMENT

                    first_cell = ws_out.cell(append_row_idx, 1)
                    first_cell.comment = Comment(
                        f"[Removed in v{new_version}]\nExisted in v{old_version}",
                        "QC Automation",
                    )
                    diff_entries.append({
                        "sheet_name": sheet_name,
                        "change_type": "REMOVED",
                        "record_key": record_key_desc,
                        "changed_field": "(Entire Record)",
                        "previous_value": f"Existed in v{old_version}",
                        "current_value": "Removed",
                    })

    if sheet_has_changes:
        ws_out.sheet_properties.tabColor = YELLOW_TAB_COLOR

    return diff_entries, sheet_has_changes


def generate_configuration_report_diff(
    old_report_path: Path,
    new_report_path: Path,
    output_dir: Path,
    gen_date: Optional[datetime] = None,
) -> Tuple[Path, List[Dict[str, Any]]]:
    """Compares two Viedoc Configuration Reports and generates a diff workbook."""
    if not old_report_path.exists():
        raise FileNotFoundError(f"Prior Configuration Report not found: {old_report_path}")
    if not new_report_path.exists():
        raise FileNotFoundError(f"New Configuration Report not found: {new_report_path}")

    if gen_date is None:
        gen_date = datetime.now()

    # Extract study metadata
    wb_meta_new = openpyxl.load_workbook(str(new_report_path), data_only=True)
    try:
        new_meta = extract_metadata(wb_meta_new)
    finally:
        wb_meta_new.close()

    wb_meta_old = openpyxl.load_workbook(str(old_report_path), data_only=True)
    try:
        old_meta = extract_metadata(wb_meta_old)
    finally:
        wb_meta_old.close()

    study_code = new_meta.study_code or old_meta.study_code or "UnknownStudy"
    new_version = new_meta.full_version or "New"
    old_version = old_meta.full_version or "Old"

    output_filename = generate_config_diff_filename(
        study_code=study_code,
        old_version=old_version,
        new_version=new_version,
        gen_date=gen_date,
    )
    output_path = output_dir / output_filename

    # Load both workbooks
    wb_old = openpyxl.load_workbook(str(old_report_path), data_only=True)
    wb_out = openpyxl.load_workbook(str(new_report_path))

    all_diff_entries: List[Dict[str, Any]] = []

    try:
        for sheet_name in list(wb_out.sheetnames):
            ws_out = wb_out[sheet_name]
            set_sheet_view_normal(ws_out)
            clear_sheet_header_footer(ws_out)

            if sheet_name not in wb_old.sheetnames:
                # Entire sheet added
                ws_out.sheet_properties.tabColor = YELLOW_TAB_COLOR
                all_diff_entries.append({
                    "sheet_name": sheet_name,
                    "change_type": "NEW",
                    "record_key": "(All Records)",
                    "changed_field": "(Entire Sheet)",
                    "previous_value": "N/A",
                    "current_value": f"Added in v{new_version}",
                })
                continue

            ws_old = wb_old[sheet_name]

            if sheet_name == "Global-Reports configuration" or (
                sheet_name in CONFIG_REPORT_SHEET_DEFS and CONFIG_REPORT_SHEET_DEFS[sheet_name].get("multi_section")
            ):
                diffs, changed = diff_global_reports_configuration(
                    ws_old, ws_out, old_version, new_version, sheet_name=sheet_name
                )
                all_diff_entries.extend(diffs)
                continue

            # Parse old records
            _, _, old_headers, old_records, old_full_rows = parse_config_sheet_records(
                ws_old, sheet_name
            )

            # Detect new sheet structure
            header_row, data_start, headers, key_names, ignore_keys = detect_sheet_structure(
                ws_out, sheet_name
            )
            header_to_idx = {h: i for i, h in enumerate(headers)}
            key_indices = [header_to_idx[k] for k in key_names if k in header_to_idx]

            sheet_has_changes = False
            seen_keys: Set[Tuple[str, ...]] = set()

            for r in range(data_start, ws_out.max_row + 1):
                row_vals = [ws_out.cell(r, c).value for c in range(1, len(headers) + 1)]
                if not any(v is not None and str(v).strip() != "" for v in row_vals):
                    continue

                if key_indices:
                    key_parts = [
                        normalize_value(row_vals[idx]) if idx < len(row_vals) else ""
                        for idx in key_indices
                    ]
                else:
                    key_parts = [normalize_value(row_vals[0])] if row_vals else [str(r)]

                record_key = tuple(key_parts)
                key_desc = " / ".join(p for p in key_parts if p)

                # Skip ignored keys (e.g., printed date/time)
                if any(normalize_value(part).lower() in ignore_keys for part in key_parts):
                    continue

                seen_keys.add(record_key)

                if record_key not in old_records:
                    # NEW ROW
                    sheet_has_changes = True
                    for c_idx in range(1, len(headers) + 1):
                        cell = ws_out.cell(r, c_idx)
                        cell.fill = YELLOW_FILL

                    all_diff_entries.append({
                        "sheet_name": sheet_name,
                        "change_type": "NEW",
                        "record_key": key_desc,
                        "changed_field": "(Entire Record)",
                        "previous_value": "N/A",
                        "current_value": f"Added in v{new_version}",
                    })
                else:
                    # Check field differences
                    old_rec = old_records[record_key]
                    row_modified = False

                    for h_name, col_idx in header_to_idx.items():
                        if h_name in old_rec:
                            old_v = old_rec[h_name]
                            new_v = row_vals[col_idx] if col_idx < len(row_vals) else None

                            if normalize_value(old_v) != normalize_value(new_v):
                                row_modified = True
                                sheet_has_changes = True
                                cell = ws_out.cell(r, col_idx + 1)
                                cell.fill = YELLOW_FILL

                                clean_old = str(old_v).rstrip("\r\n") if old_v is not None else ""
                                clean_new = str(new_v).rstrip("\r\n") if new_v is not None else ""

                                if clean_old and clean_new:
                                    rt = CellRichText()
                                    rt.append(TextBlock(DIFF_OLD_FONT, clean_old + "\n"))
                                    rt.append(TextBlock(DIFF_NEW_FONT, clean_new))
                                    cell.value = rt
                                elif clean_old and not clean_new:
                                    rt = CellRichText()
                                    rt.append(TextBlock(DIFF_OLD_FONT, clean_old))
                                    cell.value = rt
                                else:
                                    set_cell_value_safe(cell, new_v)

                                old_display = str(old_v) if old_v is not None else "(empty)"
                                cell.comment = Comment(
                                    f"[v{new_version} Modified]\nPrevious: {old_display}",
                                    "QC Automation",
                                )

                                all_diff_entries.append({
                                    "sheet_name": sheet_name,
                                    "change_type": "MODIFIED",
                                    "record_key": key_desc,
                                    "changed_field": h_name,
                                    "previous_value": str(old_v) if old_v is not None else "",
                                    "current_value": str(new_v) if new_v is not None else "",
                                })

                    if row_modified:
                        for c_idx in range(1, len(headers) + 1):
                            ws_out.cell(r, c_idx).fill = YELLOW_FILL

            # Check for REMOVED rows in this sheet
            for old_key, old_vals in old_full_rows.items():
                if old_key not in seen_keys:
                    sheet_has_changes = True
                    old_key_desc = " / ".join(p for p in old_key if p)
                    append_row_idx = ws_out.max_row + 1

                    for c_idx, val in enumerate(old_vals, 1):
                        cell = ws_out.cell(append_row_idx, c_idx)
                        set_cell_value_safe(cell, val)
                        cell.fill = REMOVED_FILL
                        cell.font = REMOVED_FONT
                        cell.border = DATA_BORDER
                        cell.alignment = DATA_ALIGNMENT

                    first_cell = ws_out.cell(append_row_idx, 1)
                    first_cell.comment = Comment(
                        f"[Removed in v{new_version}]\nExisted in v{old_version}",
                        "QC Automation",
                    )

                    all_diff_entries.append({
                        "sheet_name": sheet_name,
                        "change_type": "REMOVED",
                        "record_key": old_key_desc,
                        "changed_field": "(Entire Record)",
                        "previous_value": f"Existed in v{old_version}",
                        "current_value": "Removed",
                    })

            if sheet_has_changes:
                ws_out.sheet_properties.tabColor = YELLOW_TAB_COLOR

        # Create initial Change Log sheet
        log_sheet_name = "Change Log"
        if log_sheet_name in wb_out.sheetnames:
            ws_log = wb_out[log_sheet_name]
            ws_log.delete_rows(1, ws_log.max_row)
        else:
            ws_log = wb_out.create_sheet(title=log_sheet_name, index=0)

        set_sheet_view_normal(ws_log)
        clear_sheet_header_footer(ws_log)

        log_headers = [
            "Sheet Name",
            "Change Type",
            "Record Key",
            "Changed Field",
            "Previous Value",
            "Current Value",
        ]

        for col_idx, h in enumerate(log_headers, 1):
            cell = ws_log.cell(1, col_idx, h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.border = DATA_BORDER
            cell.alignment = Alignment(vertical="center", horizontal="center")

        current_row = 2
        for entry in all_diff_entries:
            ws_log.cell(current_row, 1, entry["sheet_name"])
            ws_log.cell(current_row, 2, entry["change_type"])
            ws_log.cell(current_row, 3, entry["record_key"])
            ws_log.cell(current_row, 4, entry["changed_field"])
            ws_log.cell(current_row, 5, entry["previous_value"])
            ws_log.cell(current_row, 6, entry["current_value"])

            row_fill = YELLOW_FILL if entry["change_type"] in ("MODIFIED", "NEW") else REMOVED_FILL
            row_font = DATA_FONT if entry["change_type"] in ("MODIFIED", "NEW") else REMOVED_FONT

            for c in range(1, 7):
                cell = ws_log.cell(current_row, c)
                cell.font = row_font
                cell.border = DATA_BORDER
                cell.alignment = DATA_ALIGNMENT
                cell.fill = row_fill
            current_row += 1

        last_log_row = max(current_row - 1, 1)
        ws_log.auto_filter.ref = f"A1:F{last_log_row}"
        adjust_column_widths(ws_log, 6, last_log_row)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb_out.save(str(output_path))
        return output_path, all_diff_entries

    finally:
        wb_old.close()
        wb_out.close()
