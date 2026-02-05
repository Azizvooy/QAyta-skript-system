@echo off
chcp 65001 >nul
echo.
echo ================================================================================
echo ИМПОРТ ИЗ GOOGLE SHEETS API В POSTGRESQL
echo ================================================================================
echo.
echo Импорт данных напрямую из Google Sheets через API
echo.

set PYTHONIOENCODING=utf-8

python scripts\database\import_from_sheets_api.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ================================================================================
    echo ГОТОВО!
    echo ================================================================================
    echo.
    echo Данные доступны:
    echo   - pgAdmin:    http://localhost:5050
    echo   - PostgreSQL: localhost:5432
    echo.
) else (
    echo.
    echo Ошибка импорта
    echo.
)

pause
