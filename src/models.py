"""Data models for Viedoc EDC QC Automation."""

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass
class StudyMetadata:
    """Metadata extracted from Viedoc Configuration Report."""
    internal_description: str = ""
    study_name: str = ""
    study_description: str = ""
    protocol_name: str = ""
    protocol_version: str = ""
    version: str = ""
    revised_version: str = ""
    printed_date_time: str = ""
    system_version: str = ""

    @property
    def full_version(self) -> str:
        """Returns version formatted as X.X (e.g., '20.0')."""
        if self.version:
            rev = self.revised_version if self.revised_version else "0"
            return f"{self.version}.{rev}"
        return ""

    @property
    def study_code(self) -> str:
        """Best identifier for study code (used in headers and filenames)."""
        return self.internal_description or self.protocol_name or self.study_name or "UnknownStudy"


@dataclass
class SheetConfig:
    """Mapping configuration between Viedoc Configuration Report and QC form sheet."""
    viedoc_sheet_name: str
    qc_sheet_name: str
    qc_form_type: str  # 'validation' or 'design'
    lowest_header_row: int  # 1-indexed row number in Viedoc report
    data_start_row: int  # 1-indexed row number where data rows begin
    has_design_version_col: bool  # True if Col 1 in QC form is 'Design Version'


# Standard sheet configurations matching Viedoc exports to SOP QC forms
SHEET_CONFIGS: List[SheetConfig] = [
    # Validation Check QC Form
    SheetConfig(
        viedoc_sheet_name="Data checks",
        qc_sheet_name="Viedoc - Data Checks",
        qc_form_type="validation",
        lowest_header_row=2,
        data_start_row=3,
        has_design_version_col=True,
    ),
    SheetConfig(
        viedoc_sheet_name="Settings-Alerts",
        qc_sheet_name="Viedoc - Alerts",
        qc_form_type="validation",
        lowest_header_row=3,
        data_start_row=4,
        has_design_version_col=True,
    ),
    # Design and Specification QC Form
    SheetConfig(
        viedoc_sheet_name="Items and Groups",
        qc_sheet_name="Viedoc - Items and Groups",
        qc_form_type="design",
        lowest_header_row=3,
        data_start_row=4,
        has_design_version_col=True,
    ),
    SheetConfig(
        viedoc_sheet_name="Study workflow-Events",
        qc_sheet_name="Viedoc - Events",
        qc_form_type="design",
        lowest_header_row=3,
        data_start_row=4,
        has_design_version_col=True,
    ),
    SheetConfig(
        viedoc_sheet_name="Study workflow-Activities",
        qc_sheet_name="Viedoc - Activities",
        qc_form_type="design",
        lowest_header_row=3,
        data_start_row=4,
        has_design_version_col=True,
    ),
    SheetConfig(
        viedoc_sheet_name="Study workflow-Forms",
        qc_sheet_name="Viedoc - Forms",
        qc_form_type="design",
        lowest_header_row=2,
        data_start_row=3,
        has_design_version_col=True,
    ),
    SheetConfig(
        viedoc_sheet_name="Functions and Conditions",
        qc_sheet_name="Viedoc - Functions",
        qc_form_type="design",
        lowest_header_row=2,
        data_start_row=3,
        has_design_version_col=True,
    ),
    SheetConfig(
        viedoc_sheet_name="Code lists",
        qc_sheet_name="Viedoc - Code lists",
        qc_form_type="design",
        lowest_header_row=2,
        data_start_row=3,
        has_design_version_col=True,
    ),
]

SHEET_CONFIGS_BY_QC_NAME: Dict[str, SheetConfig] = {
    cfg.qc_sheet_name: cfg for cfg in SHEET_CONFIGS
}

SHEET_CONFIGS_BY_VIEDOC_NAME: Dict[str, SheetConfig] = {
    cfg.viedoc_sheet_name: cfg for cfg in SHEET_CONFIGS
}


@dataclass
class ParsedSheet:
    """Represents data parsed from a single sheet in the Viedoc report."""
    qc_sheet_name: str
    viedoc_sheet_name: str
    lowest_headers: List[str]
    rows: List[List[Any]] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.rows)


@dataclass
class ParsedViedocConfig:
    """Complete parsed Viedoc Configuration Report with metadata and all sheets."""
    file_path: Path
    metadata: StudyMetadata
    sheets: Dict[str, ParsedSheet] = field(default_factory=dict)
