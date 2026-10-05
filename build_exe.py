"""Build script to compile Viedoc QC Automation into a standalone Windows .exe using PyInstaller."""

from pathlib import Path
import subprocess
import sys

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
APP_PY = PROJECT_ROOT / "app.py"
ICON_PATH = PROJECT_ROOT / "assets" / "icon.ico"

cmd = [
    sys.executable,
    "-m",
    "PyInstaller",
    "--noconfirm",
    "--onedir",
    "--windowed",
    "--name",
    "ViedocQCAutomation",
    f"--add-data={TEMPLATES_DIR};templates",
    "--collect-all=customtkinter",
    "--collect-all=openpyxl",
    str(APP_PY),
]

print("Executing PyInstaller command:")
print(" ".join(cmd))
subprocess.run(cmd, check=True, cwd=str(PROJECT_ROOT))
print("\nBuild completed successfully! Executable is located in d:\\QCproject\\dist\\ViedocQCAutomation\\ViedocQCAutomation.exe")
