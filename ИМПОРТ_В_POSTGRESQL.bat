@echo off
chcp 65001 >nul
echo.
echo ================================================================================
echo ИМПОРТ ДАННЫХ В POSTGRESQL
echo ================================================================================
echo.

python scripts\database\import_to_postgresql.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ================================================================================
    echo ✅ ИМПОРТ ЗАВЕРШЕН УСПЕШНО!
    echo ================================================================================
    echo.
    echo Теперь вы можете:
    echo   • Использовать PostgreSQL для аналитики
    echo   • Подключаться через инструменты: pgAdmin, DBeaver, psql
    echo   • Запускать SQL запросы для анализа
    echo.
    echo Полезные команды:
    echo   psql -U qayta_user -d qayta_data
    echo   SELECT * FROM v_operator_statistics;
    echo   SELECT * FROM v_service_statistics;
    echo.
) else (
    echo.
    echo ❌ Ошибка при импорте данных
    echo.
)

pause
