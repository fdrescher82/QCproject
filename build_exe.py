"""Build script to compile Viedoc QC Automation into a standalone Windows .exe using PyInstaller."""

from pathlib import Path
import shutil
import subprocess
import sys

PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
APP_PY = PROJECT_ROOT / "app.py"

# Anaconda stores libffi, Tcl, and Tk in Library\bin which PyInstaller does not scan automatically.
# We explicitly bundle these DLLs so _ctypes and _tkinter work out of the box on any PC.
ANACONDA_LIB_BIN = Path(sys.executable).parent / "Library" / "bin"
ANACONDA_DLLS = Path(sys.executable).parent / "DLLs"

extra_binaries = []
target_dlls = [
    "ffi-8.dll",
    "ffi-7.dll",
    "ffi.dll",
    "tcl86t.dll",
    "tk86t.dll",
    "sqlite3.dll",
    "zlib.dll",
]

for dll_name in target_dlls:
    dll_path = ANACONDA_LIB_BIN / dll_name
    if dll_path.exists():
        extra_binaries.append(f"--add-binary={dll_path};.")

for pyd_name in ["_ctypes.pyd", "_tkinter.pyd"]:
    pyd_path = ANACONDA_DLLS / pyd_name
    if pyd_path.exists():
        extra_binaries.append(f"--add-binary={pyd_path};.")

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

# Post-build safeguard: ensure critical DLLs are present inside _internal
dist_internal = PROJECT_ROOT / "dist" / "ViedocQCAutomation" / "_internal"
if dist_internal.exists():
    for dll_name in target_dlls:
        src_dll = ANACONDA_LIB_BIN / dll_name
        dest_dll = dist_internal / dll_name
        if src_dll.exists() and not dest_dll.exists():
            print(f"Ensuring {dll_name} in _internal...")
            shutil.copy2(src_dll, dest_dll)

    for pyd_name in ["_ctypes.pyd", "_tkinter.pyd"]:
        src_pyd = ANACONDA_DLLS / pyd_name
        dest_pyd = dist_internal / pyd_name
        if src_pyd.exists() and not dest_pyd.exists():
            print(f"Ensuring {pyd_name} in _internal...")
            shutil.copy2(src_pyd, dest_pyd)

print("\nBuild completed successfully! Executable is at: d:\\QCproject\\dist\\ViedocQCAutomation\\ViedocQCAutomation.exe")
