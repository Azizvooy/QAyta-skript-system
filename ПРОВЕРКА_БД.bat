@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo СТАТИСТИКА БАЗЫ ДАННЫХ
echo ============================================================
echo.

cd /d "%~dp0"

".venv\Scripts\python.exe" scripts\database\check_db.py

echo.
pause
