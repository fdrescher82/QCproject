"""QC Form Generator for Viedoc EDC.

Populates GCP/SOP-compliant Excel QC forms from parsed Viedoc Configuration Reports,
handling page headers, cell formatting, column widths, autofilters, and QC columns.
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.models import (
    ParsedSheet,
    ParsedViedocConfig,
    SHEET_CONFIGS_BY_QC_NAME,
    SheetConfig,
)
from src.parser import (
    format_header_date,
    generate_output_filename,
    update_header_footer_placeholders,
)

# Standard template file names
DEFAULT_VALIDATION_TEMPLATE = "GCPS_DMF_012.8_A-CValidation Check QC Form_DRAFT.xlsx"
DEFAULT_DESIGN_TEMPLATE = "GCPS_DMF_026.5_A-EDesign and Specification QC Form_DRAFT.xlsm"

# Standard cell formatting for QC forms
DATA_FONT = Font(name="Arial", size=10, bold=False, color="000000")
BORDER_SIDE = Side(style="thin", color="D9D9D9")
DATA_BORDER = Border(
    left=BORDER_SIDE, right=BORDER_SIDE, top=BORDER_SIDE, bottom=BORDER_SIDE
)
DATA_ALIGNMENT = Alignment(vertical="top", wrap_text=True)

QC_STATUS_COLUMN_WIDTH = 18
QC_DATE_COLUMN_WIDTH = 15
QC_VALIDATOR_COLUMN_WIDTH = 16
QC_SUBJECT_COLUMN_WIDTH = 14
QC_COMMENT_COLUMN_WIDTH = 25


def adjust_column_widths(ws: openpyxl.worksheet.worksheet.Worksheet, total_cols: int, total_rows: int) -> None:
    """Intelligently adjusts column widths based on header and data content length."""
    sample_row_limit = min(total_rows, 50)

    for col_idx in range(1, total_cols + 1):
        col_letter = get_column_letter(col_idx)
        header_val = ws.cell(1, col_idx).value
        header_text = str(header_val) if header_val is not None else ""

        # Check special widths for trailing QC columns
        if "Validation Status" in header_text:
            ws.column_dimensions[col_letter].width = QC_STATUS_COLUMN_WIDTH
            continue
        elif "Validation Date" in header_text:
            ws.column_dimensions[col_letter].width = QC_DATE_COLUMN_WIDTH
            continue
        elif "Validator" in header_text:
            ws.column_dimensions[col_letter].width = QC_VALIDATOR_COLUMN_WIDTH
            continue
        elif "Subject ID" in header_text:
            ws.column_dimensions[col_letter].width = QC_SUBJECT_COLUMN_WIDTH
            continue
        elif "Comment" in header_text:
            ws.column_dimensions[col_letter].width = QC_COMMENT_COLUMN_WIDTH
            continue

        # Sample cells for width calculation
        max_len = len(header_text)
        for row_idx in range(2, sample_row_limit + 1):
            val = ws.cell(row_idx, col_idx).value
            if val is not None:
                first_line = str(val).split("\n")[0]
                max_len = max(max_len, len(first_line))

        # Restrict column width between 12 and 45 characters
        calculated_width = min(max(max_len + 3, 12), 45)
        ws.column_dimensions[col_letter].width = calculated_width


def update_sheet_header_footer(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    study_code: str,
    design_version: str,
    header_date: str,
) -> None:
    """Updates page header and footer placeholders in the worksheet and sets text color to black."""
    hf = ws.HeaderFooter

    for section in [hf.oddHeader, hf.evenHeader, hf.firstHeader]:
        if section is not None:
            section.color = None
            if section.left:
                section.left.color = None
                if section.left.text:
                    section.left.text = update_header_footer_placeholders(
                        section.left.text, study_code, design_version, header_date
                    )
            if section.center:
                section.center.color = None
                if section.center.text:
                    section.center.text = update_header_footer_placeholders(
                        section.center.text, study_code, design_version, header_date
                    )
            if section.right:
                section.right.color = None
                if section.right.text:
                    section.right.text = update_header_footer_placeholders(
                        section.right.text, study_code, design_version, header_date
                    )

    for section in [hf.oddFooter, hf.evenFooter, hf.firstFooter]:
        if section is not None:
            section.color = None
            if section.left:
                section.left.color = None
                if section.left.text:
                    section.left.text = update_header_footer_placeholders(
                        section.left.text, study_code, design_version, header_date
                    )
            if section.center:
                section.center.color = None
                if section.center.text:
                    section.center.text = update_header_footer_placeholders(
                        section.center.text, study_code, design_version, header_date
                    )
            if section.right:
                section.right.color = None
                if section.right.text:
                    section.right.text = update_header_footer_placeholders(
                        section.right.text, study_code, design_version, header_date
                    )


def populate_worksheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    parsed_sheet: ParsedSheet,
    cfg: SheetConfig,
    full_version: str,
) -> None:
    """Fills a single worksheet with parsed rows and adds empty QC trailing columns."""
    # 1. Clear any pre-existing dummy rows from row 2 upwards
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    total_cols = ws.max_column
    current_row = 2

    for row_vals in parsed_sheet.rows:
        col_offset = 1

        # If sheet has 'Design Version' as Col 1, insert version string
        if cfg.has_design_version_col:
            cell = ws.cell(current_row, col_offset, full_version)
            cell.font = DATA_FONT
            cell.border = DATA_BORDER
            cell.alignment = DATA_ALIGNMENT
            col_offset += 1

        # Insert Viedoc specification data
        for val in row_vals:
            cell = ws.cell(current_row, col_offset, val)
            cell.font = DATA_FONT
            cell.border = DATA_BORDER
            cell.alignment = DATA_ALIGNMENT
            col_offset += 1

        # Format trailing 5 QC columns as empty styled cells
        # (Subject ID, Validation Date, Validator, Validation Status, Comment)
        for _ in range(5):
            cell = ws.cell(current_row, col_offset, None)
            cell.font = DATA_FONT
            cell.border = DATA_BORDER
            cell.alignment = DATA_ALIGNMENT
            col_offset += 1

        current_row += 1

    last_row = max(current_row - 1, 1)

    # Update AutoFilter to cover all populated data rows
    last_col_letter = get_column_letter(total_cols)
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    # Auto-adjust column widths
    adjust_column_widths(ws, total_cols, last_row)


def populate_qc_form(
    parsed_config: ParsedViedocConfig,
    template_path: Path,
    output_path: Path,
    qc_form_type: str,
    gen_date: Optional[datetime] = None,
) -> Path:
    """Loads a QC template, populates all sheets for the given form type, and saves as .xlsx."""
    if not template_path.exists():
        raise FileNotFoundError(f"QC template not found: {template_path}")

    if gen_date is None:
        gen_date = datetime.now()

    formatted_date = format_header_date(gen_date)
    study_code = parsed_config.metadata.study_code
    full_version = parsed_config.metadata.full_version

    wb = openpyxl.load_workbook(str(template_path))
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            cfg = SHEET_CONFIGS_BY_QC_NAME.get(sheet_name)

            # Update page header/footer on every sheet
            update_sheet_header_footer(ws, study_code, full_version, formatted_date)

            if cfg and cfg.qc_form_type == qc_form_type:
                parsed_sheet = parsed_config.sheets.get(sheet_name)
                if parsed_sheet:
                    populate_worksheet(ws, parsed_sheet, cfg, full_version)

        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb.save(str(output_path))
        return output_path
    finally:
        wb.close()


def generate_initial_qc_forms(
    parsed_config: ParsedViedocConfig,
    templates_dir: Path,
    output_dir: Path,
    gen_date: Optional[datetime] = None,
) -> Tuple[Path, Path]:
    """Generates both the Validation Check QC Form and Design and Specification QC Form

    from a parsed Viedoc Configuration Report.
    """
    if gen_date is None:
        gen_date = datetime.now()

    study_code = parsed_config.metadata.study_code
    full_version = parsed_config.metadata.full_version

    # Locate templates
    val_template = templates_dir / DEFAULT_VALIDATION_TEMPLATE
    design_template = templates_dir / DEFAULT_DESIGN_TEMPLATE

    if not val_template.exists():
        raise FileNotFoundError(f"Validation template not found at {val_template}")
    if not design_template.exists():
        raise FileNotFoundError(f"Design template not found at {design_template}")

    # Generate output file paths adhering to naming schema
    val_output_name = generate_output_filename(
        study_code=study_code,
        qc_form_name="Validation Check QC Form",
        design_version=full_version,
        gen_date=gen_date,
    )
    design_output_name = generate_output_filename(
        study_code=study_code,
        qc_form_name="Design and Specification QC Form",
        design_version=full_version,
        gen_date=gen_date,
    )

    val_output_path = output_dir / val_output_name
    design_output_path = output_dir / design_output_name

    # Populate both workbooks
    populate_qc_form(
        parsed_config=parsed_config,
        template_path=val_template,
        output_path=val_output_path,
        qc_form_type="validation",
        gen_date=gen_date,
    )

    populate_qc_form(
        parsed_config=parsed_config,
        template_path=design_template,
        output_path=design_output_path,
        qc_form_type="design",
        gen_date=gen_date,
    )

    return val_output_path, design_output_path
