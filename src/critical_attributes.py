"""Critical Attributes evaluation engine for Viedoc EDC QC Automation.

Implements Section 3 & Section 4 of SOI-EDC-QC-001 (v2.0):
- Evaluates each specification row to detect dynamic testing triggers.
- Identifies JavaScript-related critical attributes with precise reasons:
  'cross-form reference', 'complex logic', 'long code' (Issue #6).
- For Alerts, excludes 'Send message when true/false' and retains dynamic testing (Issue #7).
- Determines default QC Method ('Configuration Review' vs None for '[Subject ID]').
"""

import re
from typing import Any, Dict, List, Optional, Tuple

JS_SAFE_OBJECTS = {
    "math", "date", "number", "string", "array", "object", "json", "console", "window",
}
JS_SAFE_PROPS = {
    # String / Array properties & methods
    "length", "tofixed", "tostring", "valueof", "indexof", "includes", "contains",
    "substring", "substr", "slice", "trim", "tolowercase", "touppercase",
    "split", "join", "push", "pop", "shift", "unshift", "concat",
    "startswith", "endswith", "replace", "replaceall", "padstart", "padend",
    "charat", "charcodeat", "lastindexof", "match", "search",
    "filter", "map", "foreach", "some", "every", "reduce", "reduceright",
    "find", "findindex", "reverse", "sort", "entries", "keys", "values",
    # Math methods
    "abs", "round", "floor", "ceil", "min", "max", "pow", "sqrt", "random",
    "parse", "isnan", "isfinite",
    # Date methods
    "gettime", "getfullyear", "getmonth", "getdate", "getday", "gethours",
    "getminutes", "getseconds", "getmilliseconds", "gettimezoneoffset",
    "getutcdate", "getutcday", "getutcfullyear", "getutchours", "getutcminutes",
    "getutcmonth", "getutcseconds", "setfullyear", "setmonth", "setdate",
    "sethours", "setminutes", "setseconds", "settime",
}


def _clean_js_strings_and_comments(expr: str) -> str:
    """Strips string literals and comments from JS expressions before token analysis."""
    # Remove single-line and multi-line comments
    no_comments = re.sub(r"//.*?$|/\*.*?\*/", "", expr, flags=re.MULTILINE | re.DOTALL)
    # Replace quoted strings with empty quotes
    return re.sub(r'(".*?"|\'.*?\')', '""', no_comments)


def has_cross_form_reference(expr: str) -> bool:
    """Detects cross-form dot notation (e.g., $THIS.VS.VSDAT, SCR.DM.SEX, EXIT.DS.DSSTDAT) in JS expressions."""
    if not expr or not str(expr).strip():
        return False
    cleaned = _clean_js_strings_and_comments(str(expr))
    tokens = re.findall(
        r"[\$A-Za-z_][\$A-Za-z0-9_]*(?:\.[\$A-Za-z_][\$A-Za-z0-9_]*)+", cleaned
    )
    for token in tokens:
        parts = [p.lower() for p in token.split(".")]
        if parts[0] in JS_SAFE_OBJECTS:
            continue
        if len(parts) == 2 and parts[1] in JS_SAFE_PROPS:
            continue
        return True
    return False


def has_complex_js_logic(expr: str) -> bool:
    """Detects Regular Expressions, nested if/else / ternary, loops, switch, or nested functions."""
    if not expr or not str(expr).strip():
        return False
    raw = str(expr)
    cleaned = _clean_js_strings_and_comments(raw)

    # 1. Regex usage: new RegExp, .match(, .test(, .search(, or regex literal /pattern/flags
    if re.search(r"\bnew\s+RegExp\b|\.(match|test|search)\s*\(", cleaned):
        return True
    if re.search(r"(?:^|[=!(,:;\s])\s*/[^/\n*][^/\n]*/[gimsuy]*", cleaned):
        return True

    # 2. Nested if/else: 'else if' or 2+ 'if (...)' statements or nested ternary (? ... : ... ?)
    if re.search(r"\belse\s+if\b", cleaned):
        return True
    if len(re.findall(r"\bif\s*\(", cleaned)) >= 2:
        return True
    if cleaned.count("?") >= 2 and cleaned.count(":") >= 2:
        return True

    # 3. Loops, switch, or nested functions
    if re.search(r"\b(for|while|switch)\s*\(|\bdo\s*\{|\bfunction\b", cleaned):
        return True

    return False


def is_long_js_code(expr: str, max_chars: int = 150, max_lines: int = 4) -> bool:
    """Detects whether a JS expression contains long code (>150 characters or >=4 lines)."""
    if not expr or not str(expr).strip():
        return False
    cleaned = _clean_js_strings_and_comments(str(expr)).strip()
    if len(cleaned) >= max_chars:
        return True
    non_empty_lines = [l for l in str(expr).strip().splitlines() if l.strip()]
    if len(non_empty_lines) >= max_lines:
        return True
    return False


def get_js_critical_reasons(expr: str) -> List[str]:
    """Evaluates a JavaScript expression and returns short, precise reasons why it is critical:
    - 'cross-form reference'
    - 'complex logic'
    - 'long code'
    """
    if not expr or not str(expr).strip():
        return []
    reasons: List[str] = []
    if has_cross_form_reference(expr):
        reasons.append("cross-form reference")
    if has_complex_js_logic(expr):
        reasons.append("complex logic")
    if is_long_js_code(expr):
        reasons.append("long code")
    return reasons


def _get_col_val(
    row_map: Dict[str, Tuple[str, Any]], *candidate_names: str
) -> Tuple[Optional[str], str]:
    """Case-insensitive lookup returning (actual_header_name, trimmed_string_value)."""
    for cand in candidate_names:
        cand_lower = cand.lower()
        if cand_lower in row_map:
            orig_header, raw_val = row_map[cand_lower]
            val_str = str(raw_val).strip() if raw_val is not None else ""
            return orig_header, val_str
    # Fallback substring search
    for key_lower, (orig_header, raw_val) in row_map.items():
        for cand in candidate_names:
            if cand.lower() in key_lower:
                val_str = str(raw_val).strip() if raw_val is not None else ""
                return orig_header, val_str
    return None, ""


def evaluate_row_critical_attributes(
    qc_sheet_name: str,
    headers: List[str],
    row_vals: List[Any],
) -> Tuple[Optional[str], Optional[str]]:
    """Evaluates a row against Section 4 Critical Attributes Matrix.

    Returns:
        (critical_attributes_str, default_qc_method)
        - critical_attributes_str: Comma-separated triggering column names with JS reasons, or None.
        - default_qc_method: None if dynamic subject testing is required, or 'Configuration Review'.
    """
    row_map: Dict[str, Tuple[str, Any]] = {}
    for i, h in enumerate(headers):
        if h:
            val = row_vals[i] if i < len(row_vals) else None
            row_map[h.strip().lower()] = (h.strip(), val)

    triggered_cols: List[str] = []

    if qc_sheet_name == "Viedoc - Items and Groups":
        # 1. Show/Hide on simple condition populated
        h_name, val = _get_col_val(
            row_map,
            "Show on simple condition",
            "Show/Hide on simple condition",
            "Hide on simple condition",
        )
        if h_name and val:
            reasons = get_js_critical_reasons(val)
            if reasons:
                triggered_cols.append(f"{h_name} ({', '.join(reasons)})")
            else:
                triggered_cols.append(h_name)

        # 2. Show/Hide on advanced condition populated
        h_name, val = _get_col_val(
            row_map,
            "Show on advanced condition",
            "Show/Hide on advanced condition",
            "Hide on advanced condition",
        )
        if h_name and val:
            reasons = get_js_critical_reasons(val)
            if reasons:
                triggered_cols.append(f"{h_name} ({', '.join(reasons)})")
            else:
                triggered_cols.append(h_name)

        # 3. Enable edit for roles != All roles
        h_name, val = _get_col_val(row_map, "Enable edit for roles", "Enable edit for")
        if h_name and val and val.lower() != "all roles":
            triggered_cols.append(h_name)

        # 4. Show to roles != All roles
        h_name, val = _get_col_val(row_map, "Show to roles")
        if h_name and val and val.lower() != "all roles":
            triggered_cols.append(h_name)

        # 5. Type == Function
        h_name, val = _get_col_val(row_map, "Type", "Data type", "Item type")
        if h_name and val.lower() == "function":
            triggered_cols.append(h_name)

    elif qc_sheet_name in ("Viedoc - Events", "Viedoc - Activities"):
        h_name, val = _get_col_val(row_map, "Visibility condition")
        if h_name and val:
            reasons = get_js_critical_reasons(val)
            if reasons:
                triggered_cols.append(f"{h_name} ({', '.join(reasons)})")
            else:
                triggered_cols.append(h_name)

    elif qc_sheet_name == "Viedoc - Forms":
        # Issue #9: Item Visibility != "All" triggers critical attribute
        h_name, val = _get_col_val(row_map, "Item visibility", "Item Visibility")
        if h_name and val and val.lower() != "all":
            triggered_cols.append(h_name)

    elif qc_sheet_name == "Viedoc - Functions":
        # Issue #10: All entries in Viedoc - Functions should be reviewed.
        # Keep critical factors if present, but do not mark any entry with "Configuration Review".
        h_name, val = _get_col_val(
            row_map, "Java script Expression", "JavaScript Expression", "Expression"
        )
        if h_name and val:
            reasons = get_js_critical_reasons(val)
            if reasons:
                triggered_cols.append(f"{h_name} ({', '.join(reasons)})")

        crit_str = ", ".join(triggered_cols) if triggered_cols else None
        return crit_str, None

    elif qc_sheet_name == "Viedoc - Code lists":
        # No dynamic testing triggers for Code lists
        pass

    elif qc_sheet_name == "Viedoc - Data Checks":
        _, oid_val = _get_col_val(row_map, "OID")
        # Special rule: _NOFUTURE checks never require dynamic subject entry
        if not oid_val.upper().endswith("_NOFUTURE"):
            h_expr, expr_val = _get_col_val(row_map, "True Expression")
            if h_expr and expr_val:
                reasons = get_js_critical_reasons(expr_val)
                if reasons:
                    triggered_cols.append(f"{h_expr} ({', '.join(reasons)})")

            h_save, save_val = _get_col_val(row_map, "Allows form save")
            if h_save and save_val.lower() == "no":
                triggered_cols.append(h_save)

    elif qc_sheet_name == "Viedoc - Alerts":
        # Issue #7: Do not show "Send message when TRUE/FALSE", only other attributes (e.g. Condition JS reasons)
        h_cond, cond_val = _get_col_val(row_map, "Condition")
        if h_cond and cond_val:
            reasons = get_js_critical_reasons(cond_val)
            if reasons:
                triggered_cols.append(f"{h_cond} ({', '.join(reasons)})")

        # Active alerts require dynamic testing, but "Send message when true/false" is not shown
        crit_str = ", ".join(triggered_cols) if triggered_cols else None
        return crit_str, None

    if triggered_cols:
        return ", ".join(triggered_cols), None
    return None, "Configuration Review"