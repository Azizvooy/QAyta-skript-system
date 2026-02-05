@echo off
chcp 65001 >nul
echo.
echo ================================================================================
echo УСТАНОВКА POSTGRESQL И НАСТРОЙКА
echo ================================================================================
echo.

REM Проверка наличия PostgreSQL
where psql >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo ✅ PostgreSQL уже установлен
    psql --version
) else (
    echo ⚠️ PostgreSQL не найден
    echo.
    echo Установите PostgreSQL одним из способов:
    echo   1. Скачать с https://www.postgresql.org/download/windows/
    echo   2. Использовать Chocolatey: choco install postgresql
    echo   3. Использовать EDB Installer
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================================================
echo УСТАНОВКА PYTHON ЗАВИСИМОСТЕЙ
echo ================================================================================
echo.

pip install -r requirements_postgresql.txt

echo.
echo ================================================================================
echo СОЗДАНИЕ БАЗЫ ДАННЫХ
echo ================================================================================
echo.

REM Считываем конфигурацию
for /f "tokens=1,2 delims==" %%a in (config\postgresql.env) do (
    if "%%a"=="DB_NAME" set DB_NAME=%%b
    if "%%a"=="DB_USER" set DB_USER=%%b
    if "%%a"=="DB_PASSWORD" set DB_PASSWORD=%%b
)

echo База данных: %DB_NAME%
echo Пользователь: %DB_USER%
echo.

REM Запускаем скрипт создания схемы
python scripts\database\create_postgresql_schema.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ================================================================================
    echo ✅ POSTGRESQL НАСТРОЕН УСПЕШНО!
    echo ================================================================================
    echo.
    echo Следующий шаг:
    echo   Запустите: ИМПОРТ_В_POSTGRESQL.bat
    echo.
) else (
    echo.
    echo ❌ Ошибка при настройке PostgreSQL
    echo.
)

pause
