"""Packaging script to create a minimal, clean distribution folder and ZIP for running via launch.bat.

This packages only the strictly necessary runtime files:
- launch.bat
- app.py
- requirements.txt
- src/ (source code, excluding __pycache__)
- templates/ (master SOP templates, excluding temporary lock files)
- README.txt (quick start guide)

Outputs to:
  dist/ViedocQCAutomation_Python/
  dist/ViedocQCAutomation_Python.zip
"""

from pathlib import Path
import shutil
import zipfile

PROJECT_ROOT = Path(__file__).resolve().parent
DIST_DIR = PROJECT_ROOT / "dist"
TARGET_DIR = DIST_DIR / "ViedocQCAutomation_Python"
ZIP_PATH = DIST_DIR / "ViedocQCAutomation_Python.zip"

README_CONTENT = """================================================================================
Viedoc EDC Quality Control (QC) Automation - Python Runtime Package
================================================================================

HOW TO RUN:
1. Double-click 'launch.bat'.
2. The launcher will automatically find Python on your system (Python 3.10+ recommended).
3. If dependencies (customtkinter, openpyxl, packaging) are not yet installed,
   launch.bat will automatically install them via pip.
4. The application will launch.

REQUIREMENTS:
- Windows 10 or 11
- Python 3.10 or higher (https://www.python.org/downloads/)
  * Note: Ensure "Add Python to PATH" is checked during Python installation.

INCLUDED FILES:
- launch.bat       : One-click smart Python launcher
- app.py           : Application entry point
- requirements.txt : Python module dependencies
- src/             : Core automation logic and modern GUI interface
- templates/       : Corporate SOP master Excel templates
================================================================================
"""


def package_clean_distribution():
    print(f"Creating clean Python distribution package in:\n  {TARGET_DIR}")

    # Ensure dist folder exists
    DIST_DIR.mkdir(parents=True, exist_ok=True)

    # Clean existing target directory if present
    if TARGET_DIR.exists():
        shutil.rmtree(TARGET_DIR)
    TARGET_DIR.mkdir(parents=True)

    # 1. Copy root files
    files_to_copy = [
        "launch.bat",
        "app.py",
        "requirements.txt",
    ]
    for filename in files_to_copy:
        src_file = PROJECT_ROOT / filename
        if src_file.exists():
            shutil.copy2(src_file, TARGET_DIR / filename)
            print(f"  [+] Copied: {filename}")
        else:
            print(f"  [!] Missing file: {filename}")

    # 2. Write README.txt
    (TARGET_DIR / "README.txt").write_text(README_CONTENT, encoding="utf-8")
    print("  [+] Created: README.txt")

    # 3. Copy src/ excluding __pycache__
    src_dir = PROJECT_ROOT / "src"
    target_src = TARGET_DIR / "src"
    shutil.copytree(
        src_dir,
        target_src,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo"),
    )
    print("  [+] Copied: src/ (clean, no bytecode)")

    # 4. Copy templates/ excluding lock files
    templates_dir = PROJECT_ROOT / "templates"
    target_templates = TARGET_DIR / "templates"
    shutil.copytree(
        templates_dir,
        target_templates,
        ignore=shutil.ignore_patterns("~$*"),
    )
    print("  [+] Copied: templates/ (SOP templates)")

    # 5. Create ZIP archive
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()

    print(f"\nCompressing into ZIP archive:\n  {ZIP_PATH}")
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in TARGET_DIR.rglob("*"):
            arcname = item.relative_to(TARGET_DIR)
            zf.write(item, arcname)

    print("\nPackage generated successfully!")
    print(f"  Folder: {TARGET_DIR}")
    print(f"  ZIP:    {ZIP_PATH}")


if __name__ == "__main__":
    package_clean_distribution()
