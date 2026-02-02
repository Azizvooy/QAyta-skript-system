@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ
echo ============================================================
echo.

cd /d "%~dp0"

echo Удаление старой БД...
if exist "data\fiksa_database.db" del /f "data\fiksa_database.db"

echo.
echo Создание схемы БД...
".venv\Scripts\python.exe" scripts\database\db_schema.py

echo.
echo ============================================================
echo.
choice /C YN /M "Запустить импорт данных"
if errorlevel 2 goto end
if errorlevel 1 goto import

:import
echo.
echo Импорт данных...
".venv\Scripts\python.exe" scripts\database\import_all_data.py

:end
echo.
echo ============================================================
echo Готово!
echo ============================================================
pause
