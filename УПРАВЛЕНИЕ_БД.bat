@echo off
chcp 65001 >nul
title Управление База Данных

echo ╔════════════════════════════════════════════════════════════════╗
echo ║                 УПРАВЛЕНИЕ БАЗОЙ ДАННЫХ                       ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

python scripts\database\db_manager.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ошибка при выполнении!
    pause
    exit /b %ERRORLEVEL%
)

pause
