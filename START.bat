@echo off
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    .venv\Scripts\python.exe process_data_app.py
) else (
    python process_data_app.py
)

if errorlevel 1 pause
