@echo off
chcp 65001 >nul
title Отправка примеров отчетов в Telegram

echo ╔════════════════════════════════════════════════════════════════╗
echo ║        ОТПРАВКА ПРИМЕРОВ ОТЧЕТОВ В TELEGRAM                    ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

python scripts\telegram\send_examples.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ошибка при отправке!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ✅ Проверьте Telegram!
pause
