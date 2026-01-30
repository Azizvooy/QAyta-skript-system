@echo off
chcp 65001 >nul
title 📊 Результаты обновления

cd /d "%~dp0"

echo.
echo ════════════════════════════════════════════════════════════════════
echo     ✅ РЕЗУЛЬТАТЫ ОБНОВЛЕНИЯ
echo ════════════════════════════════════════════════════════════════════
echo.

echo [СОЗДАННЫЕ ОТЧЕТЫ]
echo.
dir /b output\reports\*.xlsx 2>nul | findstr /i "ДАШБОРД ОПЕРАТОРЫ ФИДБЭКИ" && (
    echo ✓ Отчеты созданы успешно
) || (
    echo ⚠ Отчеты не найдены
)

echo.
echo [ДАННЫЕ В БД]
echo.
python show_stats.py

echo.
echo [РАЗМЕР ФАЙЛОВ]
echo.
for %%f in (output\reports\*.xlsx) do (
    echo   %%~nf - %%~zf байт
)

echo.
echo ════════════════════════════════════════════════════════════════════
echo.
echo Выберите действие:
echo.
echo   [1] Открыть папку с отчетами
echo   [2] Просмотреть документацию
echo   [3] Сгенерировать отчеты заново
echo   [4] Выход
echo.
echo ════════════════════════════════════════════════════════════════════
echo.

set /p choice="Ваш выбор (1-4): "

if "%choice%"=="1" explorer output\reports
if "%choice%"=="2" notepad РЕЗУЛЬТАТ_ОБНОВЛЕНИЯ.md
if "%choice%"=="3" python scripts\analysis\optimized_reports.py
if "%choice%"=="4" exit /b

echo.
pause
