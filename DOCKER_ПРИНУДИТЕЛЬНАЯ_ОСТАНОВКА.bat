@echo off
chcp 65001 >nul
cls

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
echo 🛑 ПРИНУДИТЕЛЬНАЯ ОСТАНОВКА DOCKER DESKTOP
echo ================================================================================
echo.

echo [1/4] Останавливаем Docker Compose контейнеры...
cd /d "%~dp0"
docker compose down 2>nul
timeout /t 2 /nobreak >nul

echo.
echo [2/4] Останавливаем службу Docker Desktop...
net stop "Docker Desktop Service" 2>nul
net stop "com.docker.service" 2>nul
timeout /t 2 /nobreak >nul

echo.
echo [3/4] Принудительно завершаем все процессы Docker...
taskkill /F /IM "Docker Desktop.exe" 2>nul
taskkill /F /IM "com.docker.backend.exe" 2>nul
taskkill /F /IM "com.docker.build.exe" 2>nul
taskkill /F /IM "com.docker.cli.exe" 2>nul
taskkill /F /IM "com.docker.vpnkit.exe" 2>nul
taskkill /F /IM "dockerd.exe" 2>nul
timeout /t 2 /nobreak >nul

echo.
echo [4/4] Проверка остановки...
tasklist | findstr /I "docker" >nul
if %ERRORLEVEL% EQU 0 (
    echo    ⚠️ Некоторые процессы Docker все еще запущены
    echo    Показываю список:
    tasklist | findstr /I "docker"
    echo.
    echo    Попробуем еще раз...
    timeout /t 2 /nobreak >nul
    taskkill /F /IM "Docker Desktop.exe" 2>nul
    taskkill /F /IM "com.docker.backend.exe" 2>nul
    taskkill /F /IM "com.docker.build.exe" 2>nul
) else (
    echo    ✅ Все процессы Docker остановлены
)

echo.
echo ================================================================================
echo ✅ DOCKER DESKTOP ОСТАНОВЛЕН
echo ================================================================================
echo.
echo Что дальше?
echo.
echo ВАРИАНТ 1 (Рекомендуется):
echo   Перезагрузите компьютер сейчас!
echo   После перезагрузки все будет работать чисто.
echo.
echo ВАРИАНТ 2 (Быстро):
echo   1. Закройте это окно
echo   2. Запустите Docker Desktop вручную
echo   3. Дождитесь полного запуска (иконка кита станет зеленой)
echo   4. Запустите: DOCKER_АВТОЗАПУСК.bat
echo.
pause
