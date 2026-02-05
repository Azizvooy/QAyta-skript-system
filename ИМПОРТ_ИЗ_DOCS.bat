@echo off
chcp 65001 >nul
echo.
echo ════════════════════════════════════════════════════════════════════════════════
echo ИМПОРТ ДАННЫХ ИЗ GOOGLE DOCS В POSTGRESQL
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

echo 🔄 Запуск импорта...
echo.

python scripts\database\import_from_docs_to_postgresql.py

echo.
pause
