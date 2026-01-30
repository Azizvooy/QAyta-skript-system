@echo off
chcp 65001 >nul
color 0A
title 📥 ОБНОВЛЕНИЕ ДАННЫХ ИЗ GOOGLE SHEETS

echo.
echo ════════════════════════════════════════════════════════════════════
echo      📥 ОБНОВЛЕНИЕ ДАННЫХ ИЗ GOOGLE SHEETS
echo ════════════════════════════════════════════════════════════════════
echo.
echo [1] Подключение к Google Sheets...
echo [2] Загрузка данных от операторов...
echo [3] Обновление базы данных...
echo [4] Генерация отчетов...
echo.
echo ════════════════════════════════════════════════════════════════════
echo.

pause

echo.
echo [ЗАПУСК] Обновление данных...
python update_from_sheets.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ ОШИБКА при обновлении данных!
    echo.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════════════════════
echo.
echo [ЗАПУСК] Генерация отчетов...
python generate_full_report.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ ОШИБКА при генерации отчетов!
    echo.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════════════════════
echo      ✅ ВСЕ ОПЕРАЦИИ ЗАВЕРШЕНЫ УСПЕШНО!
echo ════════════════════════════════════════════════════════════════════
echo.
echo 📊 Отчеты сохранены в папке: output\reports\
echo.
echo [?] Открыть папку с отчетами?
echo     [1] Да
echo     [2] Нет
echo.
set /p choice="Ваш выбор: "

if "%choice%"=="1" (
    start "" "output\reports\"
)

echo.
pause
