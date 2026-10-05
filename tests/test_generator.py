"""Unit and integration tests for Phase 2: QC Form Generator."""

from datetime import datetime
from pathlib import Path
import openpyxl
import pytest

from src.generator import generate_initial_qc_forms
from src.parser import parse_viedoc_report

TEMPLATES_DIR = Path(r"d:\QCproject\templates")
REPORT_20_PATH = Path(r"d:\QCproject\SurfSubQ ML45551_Configuration Report_20.0_20260302.xlsx")
TEST_OUTPUT_DIR = Path(r"d:\QCproject\output_test")


def test_generate_initial_qc_forms_end_to_end():
    """Generates both QC forms from v20.0 report and verifies exact sheet counts,

    headers, page setup, and cell contents.
    """
    assert REPORT_20_PATH.exists()
    assert TEMPLATES_DIR.exists()

    gen_date = datetime(2026, 10, 5)
    parsed_config = parse_viedoc_report(REPORT_20_PATH)

    val_path, design_path = generate_initial_qc_forms(
        parsed_config=parsed_config,
        templates_dir=TEMPLATES_DIR,
        output_dir=TEST_OUTPUT_DIR,
        gen_date=gen_date,
    )

    # 1. Verify files exist and naming schema matches
    assert val_path.exists()
    assert design_path.exists()
    assert val_path.name == "SurfSubQ ML45551_20261005_Validation Check QC Form_20.0.xlsx"
    assert design_path.name == "SurfSubQ ML45551_20261005_Design and Specification QC Form_20.0.xlsx"

    # 2. Inspect Validation Check QC Form
    wb_val = openpyxl.load_workbook(str(val_path), data_only=True)
    try:
        assert wb_val.sheetnames == ["Viedoc - Data Checks", "Viedoc - Alerts"]

        # Check 'Viedoc - Data Checks'
        ws_dc = wb_val["Viedoc - Data Checks"]
        assert ws_dc.max_row == 150  # 1 header + 149 data rows
        assert ws_dc.max_column == 13
        assert ws_dc.cell(1, 1).value == "OID"
        assert ws_dc.cell(1, 13).value == "Comment"
        # First data row
        assert ws_dc.cell(2, 1).value == "RC_AENUM_23_0_1"
        assert ws_dc.cell(2, 2).value == "AE"
        # QC columns must be empty
        for col_idx in range(9, 14):
            assert ws_dc.cell(2, col_idx).value is None
        # Page Header check
        hf_dc = ws_dc.HeaderFooter.oddHeader
        assert "SurfSubQ ML45551" in hf_dc.left.text
        assert "20.0" in hf_dc.center.text
        assert "05/OCT/2026" in hf_dc.center.text

        # Check 'Viedoc - Alerts'
        ws_al = wb_val["Viedoc - Alerts"]
        assert ws_al.max_row == 66  # 1 header + 65 data rows
        assert ws_al.max_column == 25
        assert ws_al.cell(1, 1).value == "Design Version"
        assert ws_al.cell(2, 1).value == "20.0"
        assert ws_al.cell(2, 2).value == "NOTIFICATION_INV EXIT SURVEY DATA AVAILABLE"
        # QC columns must be empty
        for col_idx in range(21, 26):
            assert ws_al.cell(2, col_idx).value is None
    finally:
        wb_val.close()

    # 3. Inspect Design and Specification QC Form
    wb_des = openpyxl.load_workbook(str(design_path), data_only=True)
    try:
        expected_sheets = [
            ("Viedoc - Items and Groups", 504, 68),  # 1 header + 503 data rows
            ("Viedoc - Events", 15, 29),             # 1 header + 14 data rows
            ("Viedoc - Activities", 31, 16),         # 1 header + 30 data rows
            ("Viedoc - Forms", 49, 14),              # 1 header + 48 data rows
            ("Viedoc - Functions", 145, 15),         # 1 header + 144 data rows
            ("Viedoc - Code lists", 741, 11),        # 1 header + 740 data rows
        ]

        for s_name, exp_rows, exp_cols in expected_sheets:
            assert s_name in wb_des.sheetnames
            ws = wb_des[s_name]
            assert ws.max_row == exp_rows, f"{s_name} row count mismatch: {ws.max_row} != {exp_rows}"
            assert ws.max_column == exp_cols, f"{s_name} col count mismatch: {ws.max_column} != {exp_cols}"
            # Col 1 must be Design Version
            assert ws.cell(1, 1).value == "Design Version"
            assert ws.cell(2, 1).value == "20.0"
            # Page header verification
            hf = ws.HeaderFooter.oddHeader
            assert "SurfSubQ ML45551" in hf.left.text
            assert "20.0" in hf.center.text
    finally:
        wb_des.close()
