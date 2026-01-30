@echo off
chcp 65001 >nul
title Обработка данных 112 и Google Sheets

echo ================================================================================
echo          ПРИЛОЖЕНИЕ ДЛЯ ОБРАБОТКИ ДАННЫХ 112 И GOOGLE SHEETS
echo ================================================================================
echo.
echo Запуск приложения...
echo.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo ❌ Ошибка: Виртуальное окружение не найдено!
    echo.
    echo Пожалуйста, сначала создайте виртуальное окружение:
    echo    python -m venv .venv
    echo    .venv\Scripts\activate
    echo    pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

.venv\Scripts\python.exe process_data_app.py

if errorlevel 1 (
    echo.
    echo ❌ Ошибка при запуске приложения
    pause
)
