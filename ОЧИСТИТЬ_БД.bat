@echo off
chcp 65001 >nul
title Очистка БД от лишних записей

echo ╔════════════════════════════════════════════════════════════════╗
echo ║           ОЧИСТКА БД ОТ ЛИШНИХ ЗАПИСЕЙ                         ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.
echo ⚠️  ВНИМАНИЕ: Будут удалены операторы типа:
echo    • Тренды
echo    • Сводка
echo    • Итого
echo    • И связанные с ними данные
echo.

cd /d "%~dp0"

python scripts\database\clean_db.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ошибка при очистке!
    pause
    exit /b %ERRORLEVEL%
)

pause
