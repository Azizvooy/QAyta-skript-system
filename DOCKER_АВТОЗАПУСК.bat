@echo off
chcp 65001 >nul
cls
echo.
echo ================================================================================
echo 🐳 ПРОВЕРКА И ЗАПУСК DOCKER DESKTOP
echo ================================================================================
echo.

REM Проверка запущен ли Docker Desktop
docker info >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo ✅ Docker Desktop уже запущен!
    echo.
    goto :run_containers
)

echo ⚠️ Docker Desktop не запущен
echo.
echo Запускаю Docker Desktop...

REM Попытка найти Docker Desktop
set DOCKER_DESKTOP=""
if exist "C:\Program Files\Docker\Docker\Docker Desktop.exe" (
    set DOCKER_DESKTOP="C:\Program Files\Docker\Docker\Docker Desktop.exe"
)
if exist "%ProgramFiles%\Docker\Docker\Docker Desktop.exe" (
    set DOCKER_DESKTOP="%ProgramFiles%\Docker\Docker\Docker Desktop.exe"
)

if %DOCKER_DESKTOP%=="" (
    echo.
    echo ❌ Docker Desktop не найден!
    echo.
    echo Установите Docker Desktop:
    echo   https://www.docker.com/products/docker-desktop/
    echo.
    pause
    exit /b 1
)

REM Запуск Docker Desktop
start "" %DOCKER_DESKTOP%

echo.
echo ⏳ Ожидание запуска Docker Desktop...
echo    Это может занять 30-60 секунд...
echo.

:wait_loop
timeout /t 5 /nobreak >nul
docker info >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo    Еще ожидаем...
    goto :wait_loop
)

echo.
echo ✅ Docker Desktop запущен!
echo.

:run_containers
echo ================================================================================
echo 📦 ЗАПУСК КОНТЕЙНЕРОВ POSTGRESQL
echo ================================================================================
echo.

cd /d "%~dp0"

echo Остановка старых контейнеров...
docker compose down 2>nul

echo.
echo Запуск PostgreSQL и pgAdmin...
docker compose up -d

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ошибка при запуске контейнеров
    echo.
    echo Попробуйте:
    echo   1. Перезапустить Docker Desktop
    echo   2. Запустить этот BAT от администратора (ПКМ - Запустить от имени администратора)
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

docker compose ps

echo.
echo Проверка подключения к PostgreSQL...
docker exec qayta-postgres pg_isready -U qayta_user -d qayta_data 2>nul

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ================================================================================
    echo ✅ POSTGRESQL УСПЕШНО ЗАПУЩЕН!
    echo ================================================================================
    echo.
    echo 📊 Доступ к сервисам:
    echo    • PostgreSQL:  localhost:5432
    echo    • pgAdmin:     http://localhost:5050
    echo.
    echo 🔐 Учетные данные PostgreSQL:
    echo    Host:     localhost
    echo    Port:     5432
    echo    Database: qayta_data
    echo    User:     qayta_user
    echo    Password: qayta_password_2026
    echo.
    echo 🔐 Учетные данные pgAdmin:
    echo    URL:      http://localhost:5050
    echo    Email:    admin@qayta.uz
    echo    Password: admin
    echo.
    echo ================================================================================
    echo 📋 СЛЕДУЮЩИЙ ШАГ
    echo ================================================================================
    echo.
    echo Теперь запустите: НАСТРОЙКА_POSTGRESQL.bat
    echo.
) else (
    echo.
    echo ⚠️ PostgreSQL еще запускается...
    echo    Подождите 15 секунд и проверим снова...
    timeout /t 15 /nobreak >nul
    
    docker exec qayta-postgres pg_isready -U qayta_user -d qayta_data 2>nul
    if %ERRORLEVEL% EQU 0 (
        echo ✅ PostgreSQL готов!
    ) else (
        echo ⚠️ PostgreSQL все еще запускается
        echo    Проверьте статус: DOCKER_СТАТУС.bat
    )
    echo.
)

pause
