@echo off
chcp 65001 >nul
echo.
echo ════════════════════════════════════════════════════════════════════════════════
echo МАССОВЫЙ ИМПОРТ ДАННЫХ ВСЕХ 35 ОПЕРАТОРОВ В POSTGRESQL
echo ════════════════════════════════════════════════════════════════════════════════
echo.

REM Проверка виртуального окружения
if not exist .venv\ (
    echo ❌ Виртуальное окружение не найдено
    echo Создайте его командой: python -m venv .venv
    pause
    exit /b 1
)

REM Активация виртуального окружения и запуск скрипта
call .venv\Scripts\activate.bat

echo 🔄 Запуск импорта всех операторов...
echo.

python scripts\database\import_all_operators_to_postgresql.py

echo.
pause
