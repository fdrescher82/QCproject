# Viedoc EDC Quality Control (QC) Automation

A standalone, GCP/SOP-compliant clinical data management tool that automates the generation and amendment diffing of electronic Case Report Form (eCRF) Quality Control documentation directly from Viedoc EDC Configuration Reports.

---

## 1. Key Capabilities

### Mode 1: Initial Study Setup
* Ingests a Viedoc Configuration Report (`.xlsx`).
* Automatically extracts study metadata (`Study Code`, `Design Version`, `Report Date`).
* Generates **both** required QC forms according to corporate SOP templates:
  1. **Validation Check QC Form** (`GCPS_DMF_012.8`): Covers Data Checks and Alerts.
  2. **Design and Specification QC Form** (`GCPS_DMF_026.5`): Covers Items & Groups, Events, Activities, Forms, Functions, and Codelists.
* Sets up the 5 trailing verification columns (`Subject ID`, `Validation Date`, `Validator`, `Validation Status`, `Comment`) formatted with clean borders and typography.
* Injects `Design Version` into Column 1 for all required sheets.
* Dynamically updates Excel Page Setup headers (`HeaderFooter`):
  * `<Code>` $\rightarrow$ Study Name / Study Code (formatted in standard **black text**, clearing template red placeholder styling)
  * `<X.X>` $\rightarrow$ eCRF Design Version (e.g. `20.0`) in black text
  * `<DD/MMM/YYYY>` $\rightarrow$ Document Generation Date (e.g. `05/OCT/2026`) in black text
* Applies Excel `AutoFilter` across all data columns on Row 1 and auto-fits column widths with text wrapping.
* Names output files using the standard naming schema:
  ```
  STUDYNAME_YYYYMMDD_name of qc form_design version of ecrf.xlsx
  ```

---

### Mode 2: Study Amendment / Modification
* Takes an existing (signed-off) QC workbook and compares it against a newer Viedoc Configuration Report (e.g. updating from `20.0` to `23.0`).
* **Intelligent Record Matching:** Uses primary compound keys (e.g. `OID` for checks, `Form ID + Item Group ID + Item ID` for items, `OID + Code value` for codelists).
* **Schema Evolution:** Uses the higher version configuration report columns as the canonical structure.
* **Natural Sequence Preservation:** Maintains the exact logical order of the incoming configuration report.
* **Visual Change Highlighting:**
  * **New rows:** Highlighted with **Yellow background (`#FFF2CC`)**. QC fields are blank and ready for testing.
  * **Modified rows:** Highlighted with **Yellow background (`#FFF2CC`)**. Changed cells with eCRF data display the **prior text as strikethrough text in dark crimson red (`#C00000`)** (for optimal readability against the yellow background) followed by the **new text in standard black**, accompanied by an Excel Cell Note detailing prior values. QC Status indicates `Requires Re-QC`, while historical tester info is retained in cell notes.
  * **Unchanged rows:** Retain standard white background and preserve **100% of previous QC sign-offs** (`Subject ID`, `Validation Date`, `Validator`, `Validation Status`, `Comment`).
  * **Removed rows:** Kept in place with **Strikethrough typography** and light gray fill (`#F2F2F2`) marked `REMOVED in v[X]` so historical audit trails are never lost.
* **Audit Trail / Change Log Tab:** Automatically generates an `Amendment Change Log` sheet listing every altered field with previous vs. current values and tester details (no headers or footers).

---

### Mode 3: Configuration Report Diff (System Changes Review)
* Compares two Viedoc Configuration Reports directly across all sheets (~26 sheets, including system settings, roles, workflows, reminders).
* Retains the incoming Configuration Report layout without adding QC columns or page headers/footers.
* Visual Change Highlighting:
  * **Modified cells:** Highlighted with **Yellow background (`#FFF2CC`)** showing red strikethrough text for old values and black text for new values, plus Excel cell comments.
  * **Added rows:** Highlighted with yellow background.
  * **Removed rows:** Inserted with light gray background (`#F2F2F2`) and strikethrough font (`#7F7F7F`).
  * **Modified sheets:** Highlighted with yellow tab color (`#FFF2CC`).
* **Change Log Sheet:** Added as the first sheet (`Change Log`) with 6 core columns detailing Sheet Name, Change Type, Record Key, Changed Field, Previous Value, and Current Value.
* Excludes export timestamp noise (`Printed date/time` in `Design info`).
* Names output file using: `STUDYNAME_YYYYMMDD_Configuration Report Diff_OLDVERSION_to_NEWVERSION.xlsx`.

---

### Configurable Headers & Footers (`config/qc_config.json`)
The application loads QC header titles, form identifier prefixes, and effective dates from `config/qc_config.json`:
* **Validation Check QC Form**: Form Identifier Prefix (Left Footer, default `GCPS_DMF_012.8_C`), Effective Date (Center Footer, default `16/DEC/2024`).
* **Design & Specification QC Form**: Form Identifier Prefix (Left Footer, default `GCPS_DMF_026.4_C`), Effective Date (Center Footer, default `16/DEC/2024`).
* Each sheet automatically receives its running sheet number appended to the prefix (e.g. `GCPS_DMF_026.4_C1`, `GCPS_DMF_026.4_C2`, etc.).
* Easily adaptable in Notepad without recompiling code.

---

## 2. How to Run the Application

### Option A: Standalone Windows Executable (Recommended)
You do not need Python installed on your machine.
* Double-click **`Start_ViedocQCAutomation.bat`** in the project root, or launch:
  ```
  d:\QCproject\dist\ViedocQCAutomation\ViedocQCAutomation.exe
  ```

### Option B: Python Launcher
* Double-click **`launch.bat`**: Dynamically discovers Python via `py` or `python` in system `PATH` (with automatic fallbacks for standard installation directories) and automatically installs any missing dependencies via pip (`requirements.txt`).
* Or run manually from PowerShell/Terminal:
  ```powershell
  pip install -r requirements.txt
  python app.py
  ```

### Option C: Minimal Python Deployment Package (For copying to other PCs)
* Double-click **`Package_Python_App.bat`** (or run `python package_source.py`).
* Generates a clean, self-contained distribution folder and ZIP in `dist/`:
  - `dist\ViedocQCAutomation_Python\`
  - `dist\ViedocQCAutomation_Python.zip`
* Contains only the runtime essentials (`launch.bat`, `app.py`, `requirements.txt`, `src/`, `templates/`) with zero repository clutter, tests, or caches. You can copy either the folder or the ZIP directly to another PC.

---

## 3. Project Structure

```
d:\QCproject\
├── dist\
│   └── ViedocQCAutomation\
│       ├── ViedocQCAutomation.exe   # Standalone Windows executable
│       └── _internal\templates\     # Bundled SOP templates
├── src\
│   ├── models.py                    # Data models and sheet mapping configurations
│   ├── parser.py                    # Ingests Viedoc report and lowest headers
│   ├── generator.py                 # Populates SOP templates for Initial Study Setup
│   ├── delta_engine.py              # Amendment comparison, yellow highlights, and change logs
│   └── gui.py                       # Modern CustomTkinter desktop interface
├── templates\                       # Master corporate SOP Excel templates
│   ├── GCPS_DMF_012.8_A-CValidation Check QC Form_DRAFT.xlsx
│   └── GCPS_DMF_026.5_A-EDesign and Specification QC Form_DRAFT.xlsm
├── tests\                           # Automated GCP validation test suite
│   ├── test_parser.py
│   ├── test_generator.py
│   └── test_delta_engine.py
├── app.py                           # Python root entry point
├── build_exe.py                     # PyInstaller build automation script
├── launch.bat                       # Python one-click launcher
├── Start_ViedocQCAutomation.bat     # Executable one-click launcher
└── README.md                        # Documentation
```

---

## 4. GCP Qualification & Automated Testing

The entire system is backed by automated regression tests verifying 100% data fidelity:
```powershell
python -m pytest -v tests
```
* `test_parser.py`: Verifies lowest-header detection, metadata extraction, and filename schema.
* `test_generator.py`: Verifies initial QC workbook generation, row counts, and page headers.
* `test_delta_engine.py`: Verifies signature preservation, yellow fills, cell comments, and change log creation across both QC deliverables.
