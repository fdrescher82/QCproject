"""Build script to compile Viedoc QC Automation into a standalone Windows .exe using PyInstaller."""

from pathlib import Path
import subprocess
import sys

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
APP_PY = PROJECT_ROOT / "app.py"

# Anaconda stores libffi in Library\bin which PyInstaller does not scan automatically.
# We must explicitly bundle these DLLs so _ctypes works on any PC.
ANACONDA_LIB_BIN = Path(sys.executable).parent / "Library" / "bin"
ANACONDA_DLLS = Path(sys.executable).parent / "DLLs"

extra_binaries = []
for dll_name in ["ffi-8.dll", "ffi-7.dll", "ffi.dll"]:
    dll_path = ANACONDA_LIB_BIN / dll_name
    if dll_path.exists():
        extra_binaries.append(f"--add-binary={dll_path};.")

ctypes_pyd = ANACONDA_DLLS / "_ctypes.pyd"
if ctypes_pyd.exists():
    extra_binaries.append(f"--add-binary={ctypes_pyd};.")

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
    *extra_binaries,
    str(APP_PY),
]

print("Executing PyInstaller command:")
print(" ".join(cmd))
print()
subprocess.run(cmd, check=True, cwd=str(PROJECT_ROOT))
print("\nBuild completed! Executable is at: d:\\QCproject\\dist\\ViedocQCAutomation\\ViedocQCAutomation.exe")
