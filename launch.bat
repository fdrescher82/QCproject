@echo off
setlocal
title Viedoc EDC QC Automation

:: 0. Check for project-local virtual environment (.venv)
if exist "%~dp0.venv\Scripts\python.exe" (
    echo [INFO] Found project virtual environment in .venv.
    set "PYTHON_CMD=%~dp0.venv\Scripts\python.exe"
    goto :run_app
)

:: 1. Try py launcher from system PATH
py -c "import sys" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Found Python via 'py' in PATH.
    set "PYTHON_CMD=py"
    goto :run_app
)

:: 2. Try python from system PATH
python -c "import sys" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Found Python via 'python' in PATH.
    set "PYTHON_CMD=python"
    goto :run_app
)

:: 3. Fallbacks for standard local installations
if exist "%USERPROFILE%\anaconda3\python.exe" (
    echo [INFO] Found Python in Anaconda user directory.
    set "PYTHON_CMD=%USERPROFILE%\anaconda3\python.exe"
    goto :run_app
)

if exist "%ProgramData%\anaconda3\python.exe" (
    echo [INFO] Found Python in Anaconda system directory.
    set "PYTHON_CMD=%ProgramData%\anaconda3\python.exe"
    goto :run_app
)

if exist "%USERPROFILE%\miniconda3\python.exe" (
    echo [INFO] Found Python in Miniconda user directory.
    set "PYTHON_CMD=%USERPROFILE%\miniconda3\python.exe"
    goto :run_app
)

if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    echo [INFO] Found Python 3.13 in LocalAppData.
    set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    goto :run_app
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    echo [INFO] Found Python 3.12 in LocalAppData.
    set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto :run_app
)

if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    echo [INFO] Found Python 3.11 in LocalAppData.
    set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :run_app
)

if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    echo [INFO] Found Python 3.10 in LocalAppData.
    set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    goto :run_app
)

if exist "%ProgramFiles%\Python313\python.exe" (
    echo [INFO] Found Python 3.13 in Program Files.
    set "PYTHON_CMD=%ProgramFiles%\Python313\python.exe"
    goto :run_app
)

if exist "%ProgramFiles%\Python312\python.exe" (
    echo [INFO] Found Python 3.12 in Program Files.
    set "PYTHON_CMD=%ProgramFiles%\Python312\python.exe"
    goto :run_app
)

if exist "%ProgramFiles%\Python311\python.exe" (
    echo [INFO] Found Python 3.11 in Program Files.
    set "PYTHON_CMD=%ProgramFiles%\Python311\python.exe"
    goto :run_app
)

if exist "%ProgramFiles%\Python310\python.exe" (
    echo [INFO] Found Python 3.10 in Program Files.
    set "PYTHON_CMD=%ProgramFiles%\Python310\python.exe"
    goto :run_app
)

if exist "C:\Python312\python.exe" (
    echo [INFO] Found Python 3.12 in C:\Python312.
    set "PYTHON_CMD=C:\Python312\python.exe"
    goto :run_app
)

if exist "C:\Python311\python.exe" (
    echo [INFO] Found Python 3.11 in C:\Python311.
    set "PYTHON_CMD=C:\Python311\python.exe"
    goto :run_app
)

if exist "C:\Python310\python.exe" (
    echo [INFO] Found Python 3.10 in C:\Python310.
    set "PYTHON_CMD=C:\Python310\python.exe"
    goto :run_app
)

echo ===============================================================================
echo [ERROR] Python was not found via 'py', 'python' in system PATH, or standard install locations.
echo Please install Python 3.10+ (https://www.python.org/downloads/)
echo and ensure "Add Python to PATH" is checked during installation.
echo ===============================================================================
echo.
pause
exit /b 1

:run_app
:: Check and install required packages if missing
if "%PYTHON_CMD%"=="py" (
    py -c "import customtkinter, openpyxl, packaging" >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo [INFO] Required packages not found. Installing dependencies via pip...
        if exist "%~dp0requirements.txt" (
            py -m pip install -r "%~dp0requirements.txt"
        ) else (
            py -m pip install customtkinter openpyxl packaging
        )
    )
    echo [INFO] Launching Viedoc EDC QC Automation...
    py "%~dp0app.py"
    goto :end
)

if "%PYTHON_CMD%"=="python" (
    python -c "import customtkinter, openpyxl, packaging" >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo [INFO] Required packages not found. Installing dependencies via pip...
        if exist "%~dp0requirements.txt" (
            python -m pip install -r "%~dp0requirements.txt"
        ) else (
            python -m pip install customtkinter openpyxl packaging
        )
    )
    echo [INFO] Launching Viedoc EDC QC Automation...
    python "%~dp0app.py"
    goto :end
)

"%PYTHON_CMD%" -c "import customtkinter, openpyxl, packaging" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Required packages not found. Installing dependencies via pip...
    if exist "%~dp0requirements.txt" (
        "%PYTHON_CMD%" -m pip install -r "%~dp0requirements.txt"
    ) else (
        "%PYTHON_CMD%" -m pip install customtkinter openpyxl packaging
    )
)
echo [INFO] Launching Viedoc EDC QC Automation...
"%PYTHON_CMD%" "%~dp0app.py"

:end
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with error code: %ERRORLEVEL%
    pause
)
