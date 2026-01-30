@echo off
chcp 65001 >nul
echo ================================================================================
echo ЗАПУСК TELEGRAM БОТА В ФОНОВОМ РЕЖИМЕ
echo ================================================================================
echo.

cd /d "%~dp0"

REM Проверяем, запущен ли уже бот
tasklist /FI "WINDOWTITLE eq Telegram Bot FIKSA" 2>NUL | find /I /N "python.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [!] Бот уже запущен!
    echo [i] Для перезапуска сначала остановите бот: stop_telegram_bot.bat
    echo.
    pause
    exit /b
)

echo [1/3] Активация виртуального окружения...
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
    echo [OK] Виртуальное окружение активировано
) else (
    echo [ПРЕДУПРЕЖДЕНИЕ] Виртуальное окружение не найдено
)

echo.
echo [2/3] Запуск бота в фоновом режиме...

REM Создаем папку для логов
if not exist logs mkdir logs

REM Запускаем бот через PowerShell в фоновом режиме без окна
powershell -Command "Start-Process -FilePath '.\.venv\Scripts\python.exe' -ArgumentList 'scripts\telegram\working_bot.py' -WindowStyle Hidden -WorkingDirectory (Get-Location)"

timeout /t 2 /nobreak >nul

echo [OK] Бот запущен в фоновом режиме!
echo.
echo [3/3] Информация:
echo     - Бот работает в фоновом режиме (без окна)
echo     - Логи сохраняются в: logs\telegram_bot.log
echo     - Для остановки используйте: stop_telegram_bot.bat или kill_all_python.bat
echo     - Для просмотра логов: view_bot_logs.bat
echo.
echo [!] Бот будет работать пока вы не остановите его вручную
echo.
pause
