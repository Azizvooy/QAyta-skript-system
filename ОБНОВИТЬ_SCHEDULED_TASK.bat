@echo off
chcp 65001 >nul

echo ================================================================================
echo 🔄 ОБНОВЛЕНИЕ SCHEDULED TASK (АВТОЗАПУСК ОБНОВЛЕНИЯ)
echo ================================================================================
echo.

set "VENV_PYTHON=%CD%\.venv\Scripts\python.exe"
set "COLLECTOR=%CD%\scripts\data_collection\sheets_to_db_collector.py"

echo 📍 Пути:
echo    Python: %VENV_PYTHON%
echo    Скрипт: %COLLECTOR%
echo.

echo 🗑️  Удаление старой задачи...
schtasks /Delete /TN "GoogleSheets_HourlyUpdate" /F 2>nul

echo.
echo ➕ Создание новой задачи...
schtasks /Create /TN "GoogleSheets_HourlyUpdate" /TR "\"%VENV_PYTHON%\" \"%COLLECTOR%\"" /SC HOURLY /ST 01:00 /RL HIGHEST /F

if errorlevel 1 (
    echo.
    echo ❌ ОШИБКА при создании задачи
    echo.
    echo 💡 Попробуйте запустить этот файл от имени Администратора:
    echo    Правый клик → Запуск от имени администратора
) else (
    echo.
    echo ✅ Задача успешно создана!
    echo.
    echo 📋 Проверка:
    schtasks /Query /TN "GoogleSheets_HourlyUpdate" /FO LIST | findstr /C:"Задача" /C:"Состояние" /C:"Запустить"
)

echo.
echo ================================================================================
pause
