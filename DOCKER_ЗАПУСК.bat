@echo off
chcp 65001 >nul

REM Проверка прав администратора
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ========================================
    echo ТРЕБУЮТСЯ ПРАВА АДМИНИСТРАТОРА
    echo ========================================
    echo.
    echo Запускаю с правами администратора...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ================================================================================
echo 🐳 ЗАПУСК POSTGRESQL ЧЕРЕЗ DOCKER
echo ================================================================================
echo.

REM Проверка Docker
where docker >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Docker не установлен!
    echo.
    echo Установите Docker Desktop:
    echo   https://www.docker.com/products/docker-desktop/
    echo.
    echo После установки перезапустите этот файл.
    echo.
    pause
    exit /b 1
)

echo ✅ Docker обнаружен
docker --version
echo.

REM Проверка Docker Compose
docker compose version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Docker Compose не найден!
    echo.
    pause
    exit /b 1
)

echo ✅ Docker Compose обнаружен
docker compose version
echo.

echo ================================================================================
echo 📦 ЗАПУСК КОНТЕЙНЕРОВ
echo ================================================================================
echo.

REM Останавливаем старые контейнеры если есть
echo Остановка старых контейнеров...
docker compose down 2>nul

echo.
echo Запуск PostgreSQL и pgAdmin...
docker compose up -d

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ошибка при запуске контейнеров
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================================================
echo ⏳ ОЖИДАНИЕ ГОТОВНОСТИ POSTGRESQL
echo ================================================================================
echo.

timeout /t 10 /nobreak >nul

REM Проверка состояния
docker compose ps

echo.
echo ================================================================================
echo 🔍 ПРОВЕРКА ПОДКЛЮЧЕНИЯ
echo ================================================================================
echo.

docker exec qayta-postgres pg_isready -U qayta_user -d qayta_data

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ================================================================================
    echo ✅ POSTGRESQL УСПЕШНО ЗАПУЩЕН!
    echo ================================================================================
    echo.
    echo 📊 Информация:
    echo    • PostgreSQL: http://localhost:5432
    echo    • pgAdmin:    http://localhost:5050
    echo.
    echo 🔐 Учетные данные:
    echo    PostgreSQL:
    echo      Host:     localhost
    echo      Port:     5432
    echo      User:     qayta_user
    echo      Password: qayta_password_2026
    echo      Database: qayta_data
    echo.
    echo    pgAdmin:
    echo      Email:    admin@qayta.uz
    echo      Password: admin
    echo.
    echo ================================================================================
    echo 📋 СЛЕДУЮЩИЙ ШАГ
    echo ================================================================================
    echo.
    echo Теперь запустите: НАСТРОЙКА_POSTGRESQL.bat
    echo.
) else (
    echo.
    echo ⚠️ PostgreSQL еще не готов, подождите 10 секунд...
    timeout /t 10 /nobreak >nul
    docker exec qayta-postgres pg_isready -U qayta_user -d qayta_data
    echo.
)

pause
