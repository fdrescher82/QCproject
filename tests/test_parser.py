"""Unit tests for Phase 1: Viedoc Configuration Report Parser."""

from datetime import datetime
from pathlib import Path
import pytest

from src.models import (
    SHEET_CONFIGS,
    StudyMetadata,
)
from src.parser import (
    extract_metadata,
    format_header_date,
    generate_output_filename,
    parse_viedoc_report,
    update_header_footer_placeholders,
)

REPORT_20_PATH = Path(r"d:\QCproject\SurfSubQ ML45551_Configuration Report_20.0_20260302.xlsx")
REPORT_23_PATH = Path(r"d:\QCproject\SurfSubQ ML45551_Configuration Report_23.0_20260629.xlsx")


def test_format_header_date():
    """Test date formatting produces DD/MMM/YYYY uppercase standard."""
    test_dt = datetime(2026, 10, 5)
    formatted = format_header_date(test_dt)
    assert formatted == "05/OCT/2026"

    test_dt_jan = datetime(2026, 1, 15)
    assert format_header_date(test_dt_jan) == "15/JAN/2026"


def test_generate_output_filename():
    """Test filename naming schema: STUDYNAME_YYYYMMDD_name of qc form_design version of ecrf.xlsx"""
    test_dt = datetime(2026, 10, 5)
    fname = generate_output_filename(
        study_code="SurfSubQ ML45551",
        qc_form_name="Validation Check QC Form",
        design_version="20.0",
        gen_date=test_dt,
    )
    assert fname == "SurfSubQ ML45551_20261005_Validation Check QC Form_20.0.xlsx"

    fname_design = generate_output_filename(
        study_code="SurfSubQ ML45551",
        qc_form_name="Design and Specification QC Form",
        design_version="23.0",
        gen_date=test_dt,
    )
    assert fname_design == "SurfSubQ ML45551_20261005_Design and Specification QC Form_23.0.xlsx"


def test_update_header_footer_placeholders():
    """Test replacing <Code>, <X.X>, and <DD/MMM/YYYY> in Excel headers."""
    raw_header = "Study Code: <Code>&K01+000\nSystem: Viedoc | Version No.: <X.X>\nDate: <DD/MMM/YYYY>"
    updated = update_header_footer_placeholders(
        raw_text=raw_header,
        study_code="SurfSubQ ML45551",
        design_version="20.0",
        header_date="05/OCT/2026",
    )
    assert "<Code>" not in updated
    assert "SurfSubQ ML45551" in updated
    assert "<X.X>" not in updated
    assert "20.0" in updated
    assert "<DD/MMM/YYYY>" not in updated
    assert "05/OCT/2026" in updated


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SAMPLE_V11_PATH = PROJECT_ROOT / "output_test" / "ConfigurationReport_ViedocDefault_V11_R0_20260709113338.xlsx"
SAMPLE_V17_PATH = PROJECT_ROOT / "output_test" / "ConfigurationReport_ViedocDefault_V17_R0_20261006081911.xlsx"


def test_parse_viedoc_report_available_samples():
    """Test parsing of available Viedoc Configuration Reports (V11 and V17)."""
    assert SAMPLE_V11_PATH.exists(), f"Missing test file: {SAMPLE_V11_PATH}"
    parsed_11 = parse_viedoc_report(SAMPLE_V11_PATH)
    assert parsed_11.metadata.version == "11"
    assert parsed_11.metadata.study_name == "ViedocDefault"
    assert len(parsed_11.sheets) == 8

    if SAMPLE_V17_PATH.exists():
        parsed_17 = parse_viedoc_report(SAMPLE_V17_PATH)
        assert parsed_17.metadata.version == "17"
        assert len(parsed_17.sheets) == 8


@pytest.mark.skipif(not REPORT_20_PATH.exists(), reason="SurfSubQ 20.0 report not found in workspace")
def test_parse_viedoc_report_version_20():
    """Test full parsing of the 20.0 Viedoc Configuration Report."""
    assert REPORT_20_PATH.exists(), f"Missing test file: {REPORT_20_PATH}"
    parsed = parse_viedoc_report(REPORT_20_PATH)

    # 1. Verify Metadata
    assert parsed.metadata.internal_description == "SurfSubQ ML45551"
    assert parsed.metadata.study_name == "SurfSubQ"
    assert parsed.metadata.version == "20"
    assert parsed.metadata.revised_version == "0"
    assert parsed.metadata.full_version == "20.0"
    assert parsed.metadata.study_code == "SurfSubQ ML45551"
    assert "2026-03-02" in parsed.metadata.printed_date_time

    # 2. Verify all 8 configured sheets are parsed
    assert len(parsed.sheets) == 8

    # Check Data Checks
    data_checks = parsed.sheets["Viedoc - Data Checks"]
    assert data_checks.row_count > 0
    assert data_checks.lowest_headers[0] == "OID"
    assert data_checks.lowest_headers[1] == "Form ID"
    assert "True Expression" in data_checks.lowest_headers
    assert "Query Message" in data_checks.lowest_headers

    # Check Alerts
    alerts = parsed.sheets["Viedoc - Alerts"]
    assert alerts.row_count > 0
    assert alerts.lowest_headers[0] == "Internal description of alert"
    assert alerts.lowest_headers[1] == "Trigger type"

    # Check Items and Groups
    items = parsed.sheets["Viedoc - Items and Groups"]
    assert items.row_count > 0
    assert items.lowest_headers[0] == "Form ID"
    assert items.lowest_headers[1] == "Form Name"
    assert items.lowest_headers[2] == "Field type"

    # Check Events
    events = parsed.sheets["Viedoc - Events"]
    assert events.row_count > 0
    assert events.lowest_headers[0] == "Study event ID"

    # Check Activities
    activities = parsed.sheets["Viedoc - Activities"]
    assert activities.row_count > 0
    assert activities.lowest_headers[0] == "Study event ID"
    assert activities.lowest_headers[2] == "Activity ID"

    # Check Forms
    forms = parsed.sheets["Viedoc - Forms"]
    assert forms.row_count > 0
    assert forms.lowest_headers[0] == "Event ID"
    assert forms.lowest_headers[4] == "Form ID"

    # Check Functions
    funcs = parsed.sheets["Viedoc - Functions"]
    assert funcs.row_count > 0
    assert funcs.lowest_headers[0] == "Event ID"
    assert "Java script expression" in funcs.lowest_headers

    # Check Code lists
    codelists = parsed.sheets["Viedoc - Code lists"]
    assert codelists.row_count > 0
    assert codelists.lowest_headers[0] == "OID"
    assert codelists.lowest_headers[4] == "Code text"


@pytest.mark.skipif(not REPORT_23_PATH.exists(), reason="SurfSubQ 23.0 report not found in workspace")
def test_parse_viedoc_report_version_23():
    """Test full parsing of the 23.0 Viedoc Configuration Report."""
    assert REPORT_23_PATH.exists(), f"Missing test file: {REPORT_23_PATH}"
    parsed = parse_viedoc_report(REPORT_23_PATH)

    # Verify Version 23.0
    assert parsed.metadata.version == "23"
    assert parsed.metadata.revised_version == "0"
    assert parsed.metadata.full_version == "23.0"
    assert parsed.metadata.study_code == "SurfSubQ ML45551"

    # Verify sheet row counts exist
    for sheet_name, sheet_data in parsed.sheets.items():
        assert sheet_data.row_count > 0, f"Sheet {sheet_name} should have data rows"


def test_is_help_or_info_row():
    """Verifies identification and filtering of Viedoc informational help texts (Issue #3)."""
    from src.parser import is_help_or_info_row

    # Test cases that MUST be filtered out as help text
    help_samples = [
        ["See the Functions and Conditions sheet for details about advanced conditions.", None, None, None],
        ["See the Data checks sheet for details about data checks.", None, None],
        ["Min/Max repeat values include the first instance.", None, None, None, None],
        ['If recurrence is enabled, the "No of times setting" will not include the original event.', None],
        ["See the Event reminders sheet for detailed reminder settings.", None, None],
        ['The "At time" setting defines the time at the site where the subject is located.', None],
        ["See also the Items and Groups sheet for detailed item visibility settings.", None],
        ["Repeating number includes the first instance.", None],
        ["See the Items and Groups sheet for detailed item tracking settings.", None],
        ["Note! An external validation is needed for the configuration", None],
    ]
    for row in help_samples:
        assert is_help_or_info_row(row) is True, f"Failed to detect help text: {row[0]}"

    # Test cases that MUST be preserved as valid study specification records
    study_samples = [
        ["DM", "Demographics", "Item", "DM_GRP", "Demographics Group", "No", "1", "1", "AGE"],
        ["SCR", "Screening Visit", "Scheduled", "0", "None"],
        ["RC_BMI_8_0_1", "VS", "BMI >= 18.5 && BMI <= 30", "BMI out of range"],
        ["ALERT_01", "Form data", "AGE < 18", "SCR", "TRACKER"],
        ["EXIT", "Study Termination", "ACT_EXIT", "Exit Activity"],
    ]
    for row in study_samples:
        assert is_help_or_info_row(row) is False, f"Erroneously flagged study record: {row[0]}"

