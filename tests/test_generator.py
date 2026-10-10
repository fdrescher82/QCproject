"""Unit and integration tests for Phase 2: QC Form Generator."""

from datetime import datetime
from pathlib import Path
import openpyxl
import pytest

from src.generator import find_template_path, generate_initial_qc_forms
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
        assert ws_al.max_row == 65  # 1 header + 64 data rows (help text excluded)
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
            ("Viedoc - Items and Groups", 501, 68),  # 1 header + 500 data rows (help texts excluded)
            ("Viedoc - Events", 12, 29),             # 1 header + 11 data rows (help texts excluded)
            ("Viedoc - Activities", 28, 16),         # 1 header + 27 data rows (help texts excluded)
            ("Viedoc - Forms", 47, 14),              # 1 header + 46 data rows (help texts excluded)
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


def test_set_cell_value_safe_prevents_formula_corruption(tmp_path):
    """Verifies that values starting with '=' are stored as plain strings

    and never produce OpenXML <f> formula tags in the generated workbook.
    """
    import zipfile
    import xml.etree.ElementTree as ET
    from src.generator import set_cell_value_safe

    wb = openpyxl.Workbook()
    ws = wb.active

    test_values = ["= / > / <", "=", "== null", "== 1", "= TRUE", "Normal Text", "- days", "+ hours"]
    for i, val in enumerate(test_values, 1):
        cell = ws.cell(i, 1)
        set_cell_value_safe(cell, val)

    out_file = tmp_path / "formula_test.xlsx"
    wb.save(str(out_file))

    # Inspect OpenXML directly
    with zipfile.ZipFile(str(out_file)) as z:
        tree = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        formulas = [
            c.attrib.get("r")
            for c in tree.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c")
            if c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f") is not None
        ]
        assert len(formulas) == 0, f"Found unexpected formula tags: {formulas}"

    # Read back values
    wb2 = openpyxl.load_workbook(str(out_file))
    ws2 = wb2.active
    for i, val in enumerate(test_values, 1):
        assert ws2.cell(i, 1).value == val
        assert ws2.cell(i, 1).data_type == "s"
    wb2.close()


def test_find_template_path_standard():
    """Verifies that find_template_path discovers templates in the standard directory."""
    val_template = find_template_path(TEMPLATES_DIR, "validation")
    assert val_template.exists()
    assert "validation check qc form" in val_template.name.lower()

    des_template = find_template_path(TEMPLATES_DIR, "design")
    assert des_template.exists()
    assert "design and specification qc form" in des_template.name.lower()


def test_find_template_path_dynamic_pattern_and_lock_files(tmp_path):
    """Verifies matching renamed template files and ignoring Excel lock files (~$)."""
    # 1. Excel temporary lock file should be ignored
    lock_file = tmp_path / "~$Validation Check QC Form.xlsx"
    lock_file.write_text("lock")

    # 2. Renamed template file
    real_val = tmp_path / "SOP_DMF_012.9_Validation Check QC Form_FINAL.xlsx"
    real_val.write_text("data")

    real_des = tmp_path / "SOP_DMF_026.6_Design and Specification QC Form_APPROVED.xlsm"
    real_des.write_text("data")

    found_val = find_template_path(tmp_path, "validation")
    assert found_val == real_val

    found_des = find_template_path(tmp_path, "design")
    assert found_des == real_des


def test_find_template_path_missing_raises(tmp_path):
    """Verifies FileNotFoundError when no matching template exists."""
    with pytest.raises(FileNotFoundError, match="No template found matching keyword"):
        find_template_path(tmp_path, "validation")

    with pytest.raises(ValueError, match="Unknown qc_form_type"):
        find_template_path(tmp_path, "invalid_type")


def test_populate_qc_form_ignores_non_viedoc_sheets(tmp_path):
    """Verifies that non-Viedoc sheets in a master template are pruned from output files (Issue #2)."""
    from src.generator import populate_qc_form
    from src.parser import parse_viedoc_report

    # Create dummy template with a non-Viedoc sheet
    template_wb = openpyxl.load_workbook(str(TEMPLATES_DIR / "GCPS_DMF_012.8_A-CValidation Check QC Form_DRAFT.xlsx"))
    template_wb.create_sheet("Rave - Data Checks")
    template_wb.create_sheet("Medidata - Forms")
    test_template_path = tmp_path / "mock_validation_template.xlsx"
    template_wb.save(str(test_template_path))
    template_wb.close()

    parsed = parse_viedoc_report(REPORT_20_PATH)
    out_file = tmp_path / "output_pruned.xlsx"

    populate_qc_form(
        parsed_config=parsed,
        template_path=test_template_path,
        output_path=out_file,
        qc_form_type="validation",
    )

    out_wb = openpyxl.load_workbook(str(out_file))
    try:
        # Non-Viedoc sheets must be removed
        assert "Rave - Data Checks" not in out_wb.sheetnames
        assert "Medidata - Forms" not in out_wb.sheetnames
        # Valid sheets must be preserved
        assert "Viedoc - Data Checks" in out_wb.sheetnames
        assert "Viedoc - Alerts" in out_wb.sheetnames
    finally:
        out_wb.close()


