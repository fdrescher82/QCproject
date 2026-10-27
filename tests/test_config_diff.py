"""Unit and integration tests for Configuration Report Diff and Team Meeting Updates."""

from datetime import datetime
from pathlib import Path
import openpyxl
import pytest

from src.config_diff import (
    diff_global_reports_configuration,
    generate_config_diff_filename,
    generate_configuration_report_diff,
    parse_report_config_sections,
)
from src.delta_engine import clear_sheet_header_footer, format_validation_date
from src.generator import (
    DEFAULT_QC_CONFIG,
    load_qc_config,
    set_sheet_view_normal,
    update_sheet_header_footer,
)


def test_format_validation_date():
    """Test date parsing and standardization to DD/MMM/YYYY uppercase."""
    assert format_validation_date(None) is None
    assert format_validation_date("") is None
    assert format_validation_date("2026-10-26") == "26/OCT/2026"
    assert format_validation_date("2026/10/26") == "26/OCT/2026"
    assert format_validation_date("26/10/2026") == "26/OCT/2026"
    assert format_validation_date("26/oct/2026") == "26/OCT/2026"
    assert format_validation_date("26/OCT/2026") == "26/OCT/2026"
    assert format_validation_date(datetime(2026, 10, 26)) == "26/OCT/2026"


def test_load_qc_config():
    """Test loading of qc_config.json with fallback defaults."""
    cfg = load_qc_config()
    assert "system_name" in cfg
    assert "validation_qc_form" in cfg
    assert "design_qc_form" in cfg
    assert cfg["system_name"] == "Viedoc"
    assert "form_identifier_prefix" in cfg["validation_qc_form"]
    assert cfg["validation_qc_form"]["form_identifier_prefix"] == "GCPS_DMF_012.8_C"
    assert cfg["design_qc_form"]["form_identifier_prefix"] == "GCPS_DMF_026.4_C"


def test_update_sheet_header_footer_with_config():
    """Test that headers and footers use loaded qc_config and running sheet numbers."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Viedoc - Data Checks"

    update_sheet_header_footer(
        ws,
        study_code="DEMO_STUDY",
        design_version="1.0",
        header_date="26/OCT/2026",
        qc_form_type="validation",
    )

    hf = ws.HeaderFooter
    assert "DEMO_STUDY" in hf.oddHeader.left.text
    assert "Viedoc" in hf.oddHeader.left.text
    assert "Validation Check QC Form" in hf.oddHeader.center.text
    # Sheet 1 of validation form -> C1
    assert "GCPS_DMF_012.8_C1" in hf.oddFooter.left.text
    assert "Effective Date: 16/DEC/2024" in hf.oddFooter.center.text
    assert "Page &P of &N" in hf.oddFooter.right.text

    # Test Design form sheet numbering
    ws2 = wb.create_sheet(title="Viedoc - Items and Groups")
    update_sheet_header_footer(
        ws2,
        study_code="DEMO_STUDY",
        design_version="1.0",
        header_date="26/OCT/2026",
        qc_form_type="design",
    )
    assert "GCPS_DMF_026.4_C1" in ws2.HeaderFooter.oddFooter.left.text

    ws3 = wb.create_sheet(title="Viedoc - Events")
    update_sheet_header_footer(
        ws3,
        study_code="DEMO_STUDY",
        design_version="1.0",
        header_date="26/OCT/2026",
        qc_form_type="design",
    )
    assert "GCPS_DMF_026.4_C2" in ws3.HeaderFooter.oddFooter.left.text


def test_clear_sheet_header_footer():
    """Test clearing all headers and footers for Change Log sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.HeaderFooter.oddHeader.center.text = "Some Header"
    ws.HeaderFooter.oddFooter.center.text = "Some Footer"

    clear_sheet_header_footer(ws)
    assert ws.HeaderFooter.oddHeader.center.text is None
    assert ws.HeaderFooter.oddFooter.center.text is None


def test_set_sheet_view_normal():
    """Test setting sheet view to normal."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.sheet_view.view = "pageLayout"
    if ws.views.sheetView:
        ws.views.sheetView[0].view = "pageLayout"

    set_sheet_view_normal(ws)
    assert ws.sheet_view.view == "normal"
    if ws.views.sheetView:
        assert ws.views.sheetView[0].view == "normal"


def test_generate_config_diff_filename():
    """Test output filename generation for configuration report diff."""
    dt = datetime(2026, 10, 25)
    name = generate_config_diff_filename("TERESA", "11.0", "12.1", dt)
    assert name == "TERESA_20261025_Configuration Report Diff_11.0_to_12.1.xlsx"


def test_configuration_report_diff_integration(tmp_path):
    """Integration test comparing two sample configuration reports."""
    p16 = Path("output_test/ConfigurationReport_ViedocDefault_V16_R0_20260831110250.xlsx")
    p17 = Path("output_test/ConfigurationReport_ViedocDefault_V17_R0_20261006081911.xlsx")

    if not p16.exists() or not p17.exists():
        pytest.skip("Test configuration report files not present in output_test")

    out_path, diffs = generate_configuration_report_diff(p16, p17, tmp_path)
    assert out_path.exists()
    assert len(diffs) > 0

    # Ensure Printed date/time in Design info was excluded
    design_info_diffs = [
        d for d in diffs if d["sheet_name"] == "Design info" and d["record_key"] == "Printed date/time"
    ]
    assert len(design_info_diffs) == 0

    wb = openpyxl.load_workbook(str(out_path))
    try:
        # Check Change Log sheet is sheet 0
        assert wb.sheetnames[0] == "Change Log"
        ws_log = wb["Change Log"]
        assert ws_log.cell(1, 1).value == "Sheet Name"
        assert ws_log.cell(1, 6).value == "Current Value"

        # Check all sheets are normal view
        for ws in wb.worksheets:
            assert ws.sheet_view.view == "normal"

        # Check no headers or footers
        for ws in wb.worksheets:
            hf = ws.HeaderFooter
            for s in [hf.oddHeader, hf.evenHeader, hf.firstHeader, hf.oddFooter, hf.evenFooter, hf.firstFooter]:
                if s:
                    assert not (s.left and s.left.text)
                    assert not (s.center and s.center.text)
                    assert not (s.right and s.right.text)

        # Issue #11: When an attribute is changed, mark the whole line in yellow
        ws_items = wb["Items and Groups"]
        for r in range(4, ws_items.max_row + 1):
            if ws_items.cell(r, 9).value == "AEOUT":
                for col in range(1, 16):
                    assert ws_items.cell(r, col).fill.start_color.rgb == "00FFF2CC"
                break
    finally:
        wb.close()


def test_modified_row_entire_line_yellow_fill(tmp_path):
    """Test that when an attribute is changed, the entire row is marked yellow (Issue #11)."""
    old_file = tmp_path / "old_report.xlsx"
    new_file = tmp_path / "new_report.xlsx"

    wb_old = openpyxl.Workbook()
    ws_old = wb_old.active
    ws_old.title = "Forms"
    ws_old.cell(1, 1, "Forms")
    ws_old.cell(2, 1, "General")
    headers = ["Id", "Form name", "Description", "Type"]
    for col, h in enumerate(headers, 1):
        ws_old.cell(3, col, h)
    # Row 4: DM form
    ws_old.cell(4, 1, "DM")
    ws_old.cell(4, 2, "Demographics")
    ws_old.cell(4, 3, "Baseline demographic info")
    ws_old.cell(4, 4, "Standard")
    # Row 5: AE form
    ws_old.cell(5, 1, "AE")
    ws_old.cell(5, 2, "Adverse Events")
    ws_old.cell(5, 3, "Adverse events log")
    ws_old.cell(5, 4, "Repeating")
    wb_old.save(str(old_file))
    wb_old.close()

    wb_new = openpyxl.Workbook()
    ws_new = wb_new.active
    ws_new.title = "Forms"
    ws_new.cell(1, 1, "Forms")
    ws_new.cell(2, 1, "General")
    for col, h in enumerate(headers, 1):
        ws_new.cell(3, col, h)
    # Row 4: DM form - unchanged
    ws_new.cell(4, 1, "DM")
    ws_new.cell(4, 2, "Demographics")
    ws_new.cell(4, 3, "Baseline demographic info")
    ws_new.cell(4, 4, "Standard")
    # Row 5: AE form - modified Description only
    ws_new.cell(5, 1, "AE")
    ws_new.cell(5, 2, "Adverse Events")
    ws_new.cell(5, 3, "Updated AE description")
    ws_new.cell(5, 4, "Repeating")
    wb_new.save(str(new_file))
    wb_new.close()

    out_path, diffs = generate_configuration_report_diff(old_file, new_file, tmp_path)
    assert out_path.exists()
    assert len(diffs) == 1
    assert diffs[0]["change_type"] == "MODIFIED"
    assert diffs[0]["record_key"] == "AE"
    assert diffs[0]["changed_field"] == "Description"

    wb_out = openpyxl.load_workbook(str(out_path))
    try:
        ws_out = wb_out["Forms"]
        # Row 4 (DM) is unchanged -> no yellow fill
        for col in range(1, 5):
            cell_fill = ws_out.cell(4, col).fill
            assert cell_fill is None or cell_fill.fill_type is None or cell_fill.start_color.rgb in ("00000000", None)

        # Row 5 (AE) is modified -> ENTIRE row (all cols 1 to 4) must have yellow fill (Issue #11)
        for col in range(1, 5):
            cell_fill = ws_out.cell(5, col).fill
            assert cell_fill is not None
            assert cell_fill.start_color.rgb == "00FFF2CC"
    finally:
        wb_out.close()


def _build_sample_reports_config_sheet(ws):
    """Populates a worksheet with standard Global-Reports configuration layout."""
    ws.title = "Global-Reports configuration"
    ws.cell(1, 1, "Reports configuration")
    # Section: Visibility
    ws.cell(2, 1, "Visibility")
    vis_headers = ["Dashboard", "Demographics", "Adverse Events", "Data browser", "Reports"]
    for c, h in enumerate(vis_headers, 1):
        ws.cell(3, c, h)
    for c in range(1, 6):
        ws.cell(4, c, "")

    # Section: Dashboard
    ws.cell(5, 1, "Dashboard")
    dash_headers = ["#", "Form ID", "Form Name", "Item ID"]
    for c, h in enumerate(dash_headers, 1):
        ws.cell(6, c, h)
    ws.cell(7, 1, "1")
    for c in range(2, 5):
        ws.cell(7, c, "")

    # Section: Demographics
    ws.cell(8, 1, "Demographics")
    demo_headers = ["#", "Form ID", "Form Name", "Item ID", "Title"]
    for c, h in enumerate(demo_headers, 1):
        ws.cell(9, c, h)
    for r_idx, r_num in enumerate(range(10, 15), 1):
        ws.cell(r_num, 1, str(r_idx))
        for c in range(2, 6):
            ws.cell(r_num, c, "")

    # Section: Adverse Events
    ws.cell(15, 1, "Adverse Events")
    ae_headers = ["#", "Form ID", "Form Name", "Item ID", "Title"]
    for c, h in enumerate(ae_headers, 1):
        ws.cell(16, c, h)
    for r_idx, r_num in enumerate(range(17, 22), 1):
        ws.cell(r_num, 1, str(r_idx))
        for c in range(2, 6):
            ws.cell(r_num, c, "")

    # Section: Custom Reports
    ws.cell(22, 1, "Custom Reports")
    cr_headers = ["#", "Name", "Roles", "Filename"]
    for c, h in enumerate(cr_headers, 1):
        ws.cell(23, c, h)


def test_global_reports_configuration_identical():
    """Test that identical Global-Reports configuration sheets report no diffs (Issue #12)."""
    wb_old = openpyxl.Workbook()
    ws_old = wb_old.active
    _build_sample_reports_config_sheet(ws_old)

    wb_new = openpyxl.Workbook()
    ws_new = wb_new.active
    _build_sample_reports_config_sheet(ws_new)

    diffs, has_changes = diff_global_reports_configuration(
        ws_old, ws_new, "1.0", "2.0", "Global-Reports configuration"
    )

    assert len(diffs) == 0
    assert not has_changes
    assert ws_new.sheet_properties.tabColor is None or ws_new.sheet_properties.tabColor.rgb is None


def test_global_reports_configuration_modifications(tmp_path):
    """Test detection of changes, additions, and removals in Global-Reports configuration."""
    old_file = tmp_path / "old_report.xlsx"
    new_file = tmp_path / "new_report.xlsx"

    wb_old = openpyxl.Workbook()
    ws_old = wb_old.active
    _build_sample_reports_config_sheet(ws_old)
    # Add an existing custom report in old
    ws_old.cell(24, 1, "1")
    ws_old.cell(24, 2, "Old Custom Report")
    ws_old.cell(24, 3, "Investigator")
    ws_old.cell(24, 4, "old_report.pdf")
    wb_old.save(str(old_file))
    wb_old.close()

    wb_new = openpyxl.Workbook()
    ws_new = wb_new.active
    _build_sample_reports_config_sheet(ws_new)

    # 1. Modify Visibility: Adverse Events -> Medical Monitor
    ws_new.cell(4, 3, "Medical Monitor")

    # 2. Modify Demographics row 10 (# = 1): Form ID -> DM, Item ID -> BRTHDAT
    ws_new.cell(10, 2, "DM")
    ws_new.cell(10, 4, "BRTHDAT")

    # 3. Old Custom Report #1 is removed (no row 24 in new)

    # 4. Add a new Custom Report #2
    ws_new.cell(24, 1, "2")
    ws_new.cell(24, 2, "New Safety Report")
    ws_new.cell(24, 3, "CRA")
    ws_new.cell(24, 4, "safety.pdf")

    wb_new.save(str(new_file))
    wb_new.close()

    out_path, diffs = generate_configuration_report_diff(old_file, new_file, tmp_path)
    assert out_path.exists()

    gr_diffs = [d for d in diffs if d["sheet_name"] == "Global-Reports configuration"]
    # Expected: 4 diffs:
    # - Visibility (MODIFIED) Adverse Events
    # - Demographics / 1 (MODIFIED) Form ID
    # - Demographics / 1 (MODIFIED) Item ID
    # - Custom Reports / 2 (NEW) (Entire Record)
    # - Custom Reports / 1 (REMOVED) (Entire Record)
    assert len(gr_diffs) == 5

    change_types = {d["change_type"] for d in gr_diffs}
    assert change_types == {"MODIFIED", "NEW", "REMOVED"}

    vis_diff = next(d for d in gr_diffs if d["record_key"] == "Visibility")
    assert vis_diff["changed_field"] == "Adverse Events"
    assert vis_diff["current_value"] == "Medical Monitor"

    demo_diffs = [d for d in gr_diffs if d["record_key"] == "Demographics / 1"]
    assert len(demo_diffs) == 2
    demo_fields = {d["changed_field"]: d["current_value"] for d in demo_diffs}
    assert demo_fields["Form ID"] == "DM"
    assert demo_fields["Item ID"] == "BRTHDAT"

    new_cr = next(d for d in gr_diffs if d["change_type"] == "NEW")
    assert new_cr["record_key"] == "Custom Reports / 2"

    rem_cr = next(d for d in gr_diffs if d["change_type"] == "REMOVED")
    assert rem_cr["record_key"] == "Custom Reports / 1"

    wb_out = openpyxl.load_workbook(str(out_path))
    try:
        ws_out = wb_out["Global-Reports configuration"]
        assert ws_out.sheet_properties.tabColor is not None
        assert ws_out.sheet_properties.tabColor.rgb == "00FFF2CC"

        # Check Demographics row 10 has yellow fill across all columns (Issue #11)
        for c in range(1, 6):
            assert ws_out.cell(10, c).fill.start_color.rgb == "00FFF2CC"

        # Check Change Log sheet has the entries
        ws_log = wb_out["Change Log"]
        log_rows = [ws_log.cell(r, 3).value for r in range(2, ws_log.max_row + 1)]
        assert "Visibility" in log_rows
        assert "Demographics / 1" in log_rows
        assert "Custom Reports / 2" in log_rows
        assert "Custom Reports / 1" in log_rows
    finally:
        wb_out.close()

