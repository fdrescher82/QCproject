@echo off
title Viedoc EDC QC Automation
echo Launching Viedoc EDC QC Automation...
"C:\Users\fdres\anaconda3\python.exe" "%~dp0app.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with an error.
    pause
)
