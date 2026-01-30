@echo off
chcp 65001 >nul
echo ================================================================================
echo ПРИНУДИТЕЛЬНАЯ ОСТАНОВКА ВСЕХ ПРОЦЕССОВ БОТА
echo ================================================================================
echo.

echo ⚠️  ВНИМАНИЕ: Этот скрипт остановит ВСЕ процессы Python!
echo.
set /p confirm="Продолжить? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo Отменено
    pause
    exit /b
)

echo.
echo [1/3] Остановка всех процессов Python...
taskkill /F /IM python.exe /T >nul 2>&1

echo [2/3] Ожидание...
timeout /t 2 /nobreak >nul

echo [3/3] Проверка...
tasklist | find /I "python.exe" >nul 2>&1
if "%ERRORLEVEL%"=="0" (
    echo [!] Некоторые процессы еще работают
    tasklist | find /I "python.exe"
) else (
    echo [OK] Все процессы Python остановлены!
)

echo.
pause
