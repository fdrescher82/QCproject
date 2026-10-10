@echo off
title Package Viedoc QC Automation Python Bundle
python package_source.py
if %ERRORLEVEL% NEQ 0 (
    "%USERPROFILE%\anaconda3\python.exe" package_source.py
)
pause
