@echo off
cd /d "%~dp0"
python roi_calculator.py
if errorlevel 1 py roi_calculator.py
pause
