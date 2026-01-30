@echo off
chcp 65001 >nul
title 📊 Генерация оптимизированных отчетов

cd /d "%~dp0"

echo.
echo ════════════════════════════════════════════════════════════════════
echo     📊 ГЕНЕРАЦИЯ ОПТИМИЗИРОВАННЫХ ОТЧЕТОВ
echo ════════════════════════════════════════════════════════════════════
echo.
echo Выберите тип отчета:
echo.
echo   [1] 📈 Все отчеты (Дашборд + Операторы + Фидбэки)
echo   [2] 📊 Только статистика операторов
echo   [3] 📞 Только фидбэки от служб
echo   [4] 📈 Только дашборд
echo   [5] 🔙 Назад
echo.
echo ════════════════════════════════════════════════════════════════════
echo.

set /p choice="Ваш выбор (1-5): "

if "%choice%"=="1" goto all_reports
if "%choice%"=="2" goto operators
if "%choice%"=="3" goto feedbacks
if "%choice%"=="4" goto dashboard
if "%choice%"=="5" exit /b
goto menu

:all_reports
echo.
echo [*] Генерация всех отчетов...
python scripts\analysis\optimized_reports.py
goto end

:operators
echo.
echo [*] Генерация статистики операторов...
python -c "from scripts.analysis.optimized_reports import operator_performance_report; operator_performance_report()"
goto end

:feedbacks
echo.
echo [*] Генерация отчета по фидбэкам...
python -c "from scripts.analysis.optimized_reports import service_feedback_report; service_feedback_report()"
goto end

:dashboard
echo.
echo [*] Генерация дашборда...
python -c "from scripts.analysis.optimized_reports import dashboard_report; dashboard_report()"
goto end

:end
echo.
echo ════════════════════════════════════════════════════════════════════
echo [OK] Готово! Отчеты в папке: output\reports\
echo ════════════════════════════════════════════════════════════════════
echo.
pause
