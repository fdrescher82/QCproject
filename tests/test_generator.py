"""Unit and integration tests for Phase 2: QC Form Generator."""

from datetime import datetime
from pathlib import Path
import re
import zipfile
import openpyxl
import pytest

from src.generator import find_template_path, generate_initial_qc_forms
from src.parser import parse_viedoc_report

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
REPORT_20_PATH = PROJECT_ROOT / "SurfSubQ ML45551_Configuration Report_20.0_20260302.xlsx"
SAMPLE_V11_PATH = PROJECT_ROOT / "output_test" / "ConfigurationReport_ViedocDefault_V11_R0_20260709113338.xlsx"
TEST_OUTPUT_DIR = PROJECT_ROOT / "output_test"


@pytest.mark.skipif(not REPORT_20_PATH.exists(), reason="SurfSubQ 20.0 report not found in workspace")
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
    from src.generator import populate_qc_form, find_template_path
    from src.parser import parse_viedoc_report

    val_template = find_template_path(TEMPLATES_DIR, "validation")
    template_wb = openpyxl.load_workbook(str(val_template))
    template_wb.create_sheet("Rave - Data Checks")
    template_wb.create_sheet("Medidata - Forms")
    test_template_path = tmp_path / "mock_validation_template.xlsx"
    template_wb.save(str(test_template_path))
    template_wb.close()

    report_path = SAMPLE_V11_PATH if SAMPLE_V11_PATH.exists() else REPORT_20_PATH
    parsed = parse_viedoc_report(report_path)
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
        assert "QCTMS LC" not in out_wb.sheetnames
        assert "Castor" not in out_wb.sheetnames
        assert "New" not in out_wb.sheetnames
        # Valid sheets must be preserved
        assert "Viedoc - Data Checks" in out_wb.sheetnames
        assert "Viedoc - Alerts" in out_wb.sheetnames
    finally:
        out_wb.close()


def test_updated_templates_pruning_mode1(tmp_path):
    """Verifies that the new master SOP templates containing other EDC sheets (QCTMS LC, Castor, New, etc.)

    are ingested correctly and that output files omit all non-Viedoc sheets in Initial Setup (Mode 1).
    """
    assert SAMPLE_V11_PATH.exists(), f"Sample report missing: {SAMPLE_V11_PATH}"
    parsed = parse_viedoc_report(SAMPLE_V11_PATH)

    val_path, des_path = generate_initial_qc_forms(
        parsed_config=parsed,
        templates_dir=TEMPLATES_DIR,
        output_dir=tmp_path,
        gen_date=datetime(2026, 10, 13),
    )

    # 1. Validation QC Form
    val_wb = openpyxl.load_workbook(str(val_path))
    try:
        expected_val = ["Viedoc - Data Checks", "Viedoc - Alerts"]
        assert val_wb.sheetnames == expected_val
        for forbidden in ["QCTMS LC", "Castor", "New", "QCTMS ePRO"]:
            assert forbidden not in val_wb.sheetnames
    finally:
        val_wb.close()

    # 2. Design QC Form
    des_wb = openpyxl.load_workbook(str(des_path))
    try:
        expected_des = [
            "Viedoc - Items and Groups",
            "Viedoc - Events",
            "Viedoc - Activities",
            "Viedoc - Forms",
            "Viedoc - Functions",
            "Viedoc - Code lists",
        ]
        assert des_wb.sheetnames == expected_des
        for forbidden in ["QCTMS LC", "Castor", "New", "QCTMS ePRO"]:
            assert forbidden not in des_wb.sheetnames
    finally:
        des_wb.close()


def test_header_format_alignment_issue5():
    """Verifies that update_sheet_header_footer produces headers adhering to specifications:
    - Left: Study Code & System
    - Center: 4 separate lines with Arial sizes 16, 14, 11, 11
    - Right: &G (logo)
    """
    from src.generator import update_sheet_header_footer

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Viedoc - Data Checks"

    update_sheet_header_footer(
        ws=ws,
        study_code="ML45551",
        design_version="20.0",
        header_date="16/OCT/2026",
        qc_form_name="Validation Check QC Form",
    )

    hf = ws.oddHeader
    expected_left = (
        '&"Arial"&10Study Code: ML45551\n'
        'System: Viedoc'
    )
    assert hf.left.text == expected_left
    expected_center = (
        '&"Arial"&16Validation Check QC Form\n'
        '&"Arial"&14Data Checks\n'
        '&"Arial"&11Version No.: 20.0\n'
        '&"Arial"&11Version Date: 16/OCT/2026'
    )
    assert hf.center.text == expected_center
    assert hf.right.text == "&G"


def test_generated_qc_form_logo_and_header_injection_issue5(tmp_path):
    """Verifies end-to-end that generating a QC form properly injects:
    - GCP-Service logo image (xl/media/image1.jpeg)
    - Right Header VML drawing with shape id="RH"
    - legacyDrawingHF element in each worksheet
    - Clean newlines without _x000a_ artifacts
    """
    assert SAMPLE_V11_PATH.exists(), f"Sample report missing: {SAMPLE_V11_PATH}"

    parsed = parse_viedoc_report(SAMPLE_V11_PATH)
    val_path, des_path = generate_initial_qc_forms(
        parsed_config=parsed,
        templates_dir=TEMPLATES_DIR,
        output_dir=tmp_path,
        gen_date=datetime(2026, 10, 16),
    )

    for out_file, expected_form_name in [(val_path, "Validation Check QC Form"), (des_path, "Design and Specification QC Form")]:
        assert out_file.exists()
        with zipfile.ZipFile(str(out_file)) as z:
            names = z.namelist()
            # 1. Logo image in media
            media_files = [n for n in names if n.startswith("xl/media/") and n.endswith((".jpeg", ".jpg", ".png"))]
            assert len(media_files) > 0, f"No logo image found in {out_file.name}"
            assert z.getinfo(media_files[0]).file_size > 1000

            # 2. VML drawings with RH shape
            vml_files = [n for n in names if "vmlDrawingHF" in n and n.endswith(".vml")]
            assert len(vml_files) > 0, f"No header VML drawings found in {out_file.name}"
            for vml in vml_files:
                vml_content = z.read(vml).decode("utf-8")
                assert 'id="RH"' in vml_content, f"VML shape id='RH' not found in {vml}"

            # 3. Worksheets have legacyDrawingHF and clean newlines
            sheet_files = [n for n in names if re.match(r"^xl/worksheets/sheet\d+\.xml$", n)]
            for sf in sheet_files:
                sheet_xml = z.read(sf).decode("utf-8")
                assert "<legacyDrawingHF" in sheet_xml, f"legacyDrawingHF missing in {sf}"
                assert "_x000a_" not in sheet_xml, f"Unescaped _x000a_ found in {sf}"
                assert "&amp;R&amp;G" in sheet_xml, f"Right logo &amp;R&amp;G missing in {sf}"
                assert f'&amp;16{expected_form_name}' in sheet_xml, f"Arial 16 form name missing in {sf}"

        # 4. Openpyxl can reload without error
        wb = openpyxl.load_workbook(str(out_file))
        assert wb.sheetnames
        ws = wb[wb.sheetnames[0]]
        assert ws.oddHeader.right.text == "&G"
        assert "Study Code:" in ws.oddHeader.left.text
        wb.close()


def test_templates_aligned_issue5():
    """Verifies that templates in templates/ directory have aligned headers and RH VML shapes."""
    for tpath in TEMPLATES_DIR.glob("*.xlsx"):
        with zipfile.ZipFile(str(tpath)) as z:
            # Check VML shapes have id="RH"
            vml_files = [n for n in z.namelist() if "vmlDrawing" in n and n.endswith(".vml")]
            for vml in vml_files:
                content = z.read(vml).decode("utf-8", errors="replace")
                assert 'id="RH"' in content, f"Template {tpath.name} {vml} shape should be id='RH'"
                assert 'id="LH"' not in content, f"Template {tpath.name} {vml} still has id='LH'"


def test_get_default_output_dir():
    """Verifies that get_default_output_dir returns a dynamic path ending with 'output' and exists."""
    from src.gui import get_default_output_dir
    out_dir = get_default_output_dir()
    assert out_dir.name == "output"
    assert out_dir.exists()
    assert out_dir.is_dir()


def test_validation_form_data_checks_columns_and_grey_headers(tmp_path):
    """Verifies Issue #14: Validation QC form Data Checks sheet uses correct data check columns

    and that the QC columns have grey header styling.
    """
    assert SAMPLE_V11_PATH.exists()
    parsed_config = parse_viedoc_report(SAMPLE_V11_PATH)
    val_path, _ = generate_initial_qc_forms(
        parsed_config=parsed_config,
        templates_dir=TEMPLATES_DIR,
        output_dir=tmp_path,
    )

    wb = openpyxl.load_workbook(val_path)
    assert "Viedoc - Data Checks" in wb.sheetnames
    ws = wb["Viedoc - Data Checks"]

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    assert len(headers) == 15
    assert headers[0] == "Design Version"
    assert headers[1] == "OID"
    assert headers[2] == "Form ID"
    assert headers[3] in ("Form Name", "Form name")
    assert headers[4] == "Field ID"
    assert headers[5] == "Field Output Label"
    assert headers[6] == "True Expression"
    assert headers[7] == "Query Message"
    assert headers[8] in ("Allows form save", "Allow form save")
    assert headers[9:15] == [
        "Critical Attributes",
        "QC Method",
        "Validation Date",
        "Validator",
        "Validation Status",
        "Comment",
    ]

    # Verify no alert headers are in Data Checks
    alert_headers = {"Internal description of alert", "Trigger type", "Condition", "Context form", "Action type"}
    assert not any(h in alert_headers for h in headers)

    # Verify QC headers (cols 10-15) have grey fill
    for c in range(10, 16):
        cell = ws.cell(1, c)
        assert cell.fill and cell.fill.fgColor is not None
        assert cell.fill.fgColor.theme == 0
        assert round(cell.fill.fgColor.tint, 2) == -0.15






