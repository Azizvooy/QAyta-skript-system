@echo off
chcp 65001 >nul
title Обработка данных 112 и Google Sheets
color 0A

echo ============================================================
echo    ЗАПУСК ПРИЛОЖЕНИЯ ДЛЯ ОБРАБОТКИ ДАННЫХ
echo ============================================================
echo.
echo Активация виртуального окружения...
echo.

cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo Виртуальное окружение активировано
    echo.
) else (
    echo ВНИМАНИЕ: Виртуальное окружение не найдено
    echo Запуск без активации...
    echo.
)

echo Запуск приложения...
echo.
python process_data_app.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo    ОШИБКА ПРИ ЗАПУСКЕ!
    echo ============================================================
    echo.
    pause
) else (
    echo.
    echo Приложение закрыто
)

exit
