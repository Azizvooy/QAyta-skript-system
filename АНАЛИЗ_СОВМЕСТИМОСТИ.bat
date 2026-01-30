@echo off
chcp 65001 >nul
title Анализ Совместимости Проекта

echo ╔════════════════════════════════════════════════════════════════╗
echo ║            АНАЛИЗ СОВМЕСТИМОСТИ С НОВОЙ БД                     ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

python scripts\database\analyze_compatibility.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ошибка при выполнении!
    pause
    exit /b %ERRORLEVEL%
)

pause
