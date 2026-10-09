@echo off
title Viedoc EDC QC Automation

:: 1. Try py launcher from system PATH
py -c "import sys" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Found Python via 'py' in PATH.
    py "%~dp0app.py"
    goto :end
)

:: 2. Try python from system PATH
python -c "import sys" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Found Python via 'python' in PATH.
    python "%~dp0app.py"
    goto :end
)

:: 3. Fallbacks for standard local installations
if exist "%USERPROFILE%\anaconda3\python.exe" (
    echo [INFO] Found Python in Anaconda user directory.
    "%USERPROFILE%\anaconda3\python.exe" "%~dp0app.py"
    goto :end
)

if exist "%ProgramData%\anaconda3\python.exe" (
    echo [INFO] Found Python in Anaconda system directory.
    "%ProgramData%\anaconda3\python.exe" "%~dp0app.py"
    goto :end
)

if exist "%USERPROFILE%\miniconda3\python.exe" (
    echo [INFO] Found Python in Miniconda user directory.
    "%USERPROFILE%\miniconda3\python.exe" "%~dp0app.py"
    goto :end
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0app.py"
    goto :end
)

if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" "%~dp0app.py"
    goto :end
)

if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" "%~dp0app.py"
    goto :end
)

echo [ERROR] Python was not found via 'py', 'python' in system PATH, or standard install locations.
echo Please install Python 3.10+ and ensure "Add Python to PATH" is checked during installation.
echo.
pause
exit /b 1

:end
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with error code: %ERRORLEVEL%
    pause
)
