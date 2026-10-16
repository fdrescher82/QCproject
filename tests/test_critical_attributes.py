"""Unit tests for Critical Attributes evaluation engine (SOI-EDC-QC-001 v2.0)."""

import pytest
from src.critical_attributes import (
    _clean_js_strings_and_comments,
    evaluate_row_critical_attributes,
    has_complex_js_logic,
    has_cross_form_reference,
)


def test_clean_js_strings_and_comments():
    expr = 'var x = "hello // world"; // comment\n/* multi\nline */ var y = \'test\';'
    cleaned = _clean_js_strings_and_comments(expr)
    assert "// comment" not in cleaned
    assert "multi" not in cleaned
    assert '""' in cleaned


def test_has_cross_form_reference():
    # Built-in JS object and method calls should NOT trigger cross-form reference
    assert not has_cross_form_reference("Math.round(x)")
    assert not has_cross_form_reference("Date.parse(val)")
    assert not has_cross_form_reference("val.toString()")
    assert not has_cross_form_reference("val.toLowerCase()")
    assert not has_cross_form_reference("arr.length")
    assert not has_cross_form_reference("x > 10 && y < 20")

    # Cross-form references SHOULD trigger
    assert has_cross_form_reference("$THIS.VS.VSDAT")
    assert has_cross_form_reference("SCR.DM.SEX == 'F'")
    assert has_cross_form_reference("ITEM1.SUBFIELD.VAL")


def test_has_complex_js_logic():
    # Simple logic - not complex
    assert not has_complex_js_logic("x > 10")
    assert not has_complex_js_logic("if (x > 10) { return true; } else { return false; }")
    assert not has_complex_js_logic("x > 10 ? 1 : 0")

    # Regular expressions
    assert has_complex_js_logic("new RegExp('^[0-9]+$')")
    assert has_complex_js_logic("/^[A-Z]{3}$/.test(val)")
    assert has_complex_js_logic("val.match(/abc/)")

    # Nested if / else if
    assert has_complex_js_logic("if (x > 10) {} else if (x > 5) {}")
    assert has_complex_js_logic("if (a) { if (b) { return 1; } }")

    # Nested ternary
    assert has_complex_js_logic("a ? 1 : b ? 2 : 3")


def test_evaluate_items_and_groups():
    headers = [
        "Item OID",
        "Type",
        "Show/Hide on simple condition",
        "Show/Hide on advanced condition",
        "Enable edit for roles",
        "Show to roles",
    ]

    # Standard row with default roles
    row_standard = ["ITEM_1", "Text", "", "", "All roles", "All roles"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Items and Groups", headers, row_standard)
    assert crit is None
    assert method == "Code review"

    # Function type item
    row_func = ["ITEM_2", "Function", "", "", "All roles", "All roles"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Items and Groups", headers, row_func)
    assert crit == "Type"
    assert method is None

    # Simple show/hide condition
    row_sh = ["ITEM_3", "Text", "AGE > 18", "", "All roles", "All roles"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Items and Groups", headers, row_sh)
    assert crit == "Show/Hide on simple condition"
    assert method is None

    # Restricted role
    row_role = ["ITEM_4", "Text", "", "", "Investigator", "All roles"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Items and Groups", headers, row_role)
    assert crit == "Enable edit for roles"
    assert method is None


def test_evaluate_events_and_activities():
    headers = ["Event OID", "Event Name", "Visibility condition"]

    # No visibility condition
    crit, method = evaluate_row_critical_attributes("Viedoc - Events", headers, ["EV1", "Visit 1", ""])
    assert crit is None
    assert method == "Code review"

    # With visibility condition
    crit, method = evaluate_row_critical_attributes("Viedoc - Events", headers, ["EV2", "Visit 2", "DM.SEX == 1"])
    assert crit == "Visibility condition"
    assert method is None

    # Activities sheet
    crit, method = evaluate_row_critical_attributes("Viedoc - Activities", headers, ["ACT1", "Act 1", "DM.ARM == 'A'"])
    assert crit == "Visibility condition"
    assert method is None


def test_evaluate_forms_and_codelists():
    # Forms and Code lists have no critical attributes triggers
    headers = ["Form OID", "Form Name"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Forms", headers, ["FORM1", "Demographics"])
    assert crit is None
    assert method == "Code review"

    cl_headers = ["Code List ID", "Data Type", "Code", "Text"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Code lists", cl_headers, ["CL1", "Integer", "1", "Male"])
    assert crit is None
    assert method == "Code review"


def test_evaluate_functions():
    headers = ["Function ID", "JavaScript Expression"]

    # Simple math function
    row_simple = ["FN1", "Math.round(WEIGHT / (HEIGHT * HEIGHT))"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Functions", headers, row_simple)
    assert crit is None
    assert method == "Code review"

    # Function with cross-form reference
    row_cross = ["FN2", "$THIS.VS.HEIGHT * 100"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Functions", headers, row_cross)
    assert crit == "JavaScript Expression"
    assert method is None


def test_evaluate_data_checks():
    headers = ["OID", "Check Name", "True Expression", "Allows form save"]

    # Standard check, allows save Yes
    row_std = ["RC_01", "Age check", "AGE < 18", "Yes"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Data Checks", headers, row_std)
    assert crit is None
    assert method == "Code review"

    # Cross-form reference in expression
    row_cross = ["RC_02", "Visit date", "$THIS.VS.VSDAT > SCR.DM.BRTHDAT", "Yes"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Data Checks", headers, row_cross)
    assert crit == "True Expression"
    assert method is None

    # Blocking query (Allows form save == No)
    row_block = ["RC_03", "Required check", "VAL == null", "No"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Data Checks", headers, row_block)
    assert crit == "Allows form save"
    assert method is None

    # Special rule: _NOFUTURE checks never require dynamic subject testing
    row_nofuture = ["RC_DATE_NOFUTURE", "Future date check", "$THIS.VS.VSDAT > new Date()", "No"]
    crit, method = evaluate_row_critical_attributes("Viedoc - Data Checks", headers, row_nofuture)
    assert crit is None
    assert method == "Code review"


def test_evaluate_alerts():
    headers = ["Alert ID", "Alert Name", "Send message when TRUE/FALSE"]

    # Active alert
    crit, method = evaluate_row_critical_attributes("Viedoc - Alerts", headers, ["ALT1", "SAE Alert", "TRUE"])
    assert crit == "Send message when TRUE/FALSE"
    assert method is None
