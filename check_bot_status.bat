@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ════════════════════════════════════════════════════════════════════════════
echo СТАТУС TELEGRAM БОТА
echo ════════════════════════════════════════════════════════════════════════════
echo.

set bot_running=0
set bot_pid=

REM Проверка 1: По заголовку окна
for /f "tokens=2" %%a in ('tasklist /FI "WINDOWTITLE eq Telegram Bot FIKSA" /FO LIST 2^>nul ^| find "PID:"') do (
    set bot_running=1
    set bot_pid=%%a
)

REM Проверка 2: По имени скрипта working_bot.py
if "!bot_running!"=="0" (
    for /f "tokens=2" %%a in ('wmic process where "commandline like '%%working_bot.py%%'" get processid 2^>nul ^| findstr /r "[0-9]"') do (
        set bot_running=1
        set bot_pid=%%a
    )
)

REM Проверка 3: По логам (последние 30 секунд активности)
if exist logs\telegram_bot.log (
    for /f "tokens=*" %%a in ('powershell -Command "(Get-Item logs\telegram_bot.log).LastWriteTime | Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"') do (
        set last_log_time=%%a
    )
)

if "!bot_running!"=="1" (
    echo [✅] БОТ РАБОТАЕТ
    echo.
    echo Информация:
    echo   PID: !bot_pid!
    if defined last_log_time echo   Последняя активность: !last_log_time!
) else (
    REM Дополнительная проверка через логи
    if exist logs\telegram_bot.log (
        powershell -Command "$lastWrite = (Get-Item logs\telegram_bot.log).LastWriteTime; $diff = (Get-Date) - $lastWrite; if ($diff.TotalSeconds -lt 30) { exit 0 } else { exit 1 }" >nul 2>&1
        if !ERRORLEVEL! EQU 0 (
            echo [✅] БОТ РАБОТАЕТ (фоновый режим)
            echo.
            echo Последняя активность менее 30 секунд назад
            if defined last_log_time echo Время: !last_log_time!
            set bot_running=1
        )
    )
)

if "!bot_running!"=="0" (
    echo [❌] БОТ НЕ РАБОТАЕТ
    echo.
    echo Для запуска используйте:
    echo   • МЕНЮ_БОТА.bat → выбор 1
    echo   • start_telegram_bot.bat
)

echo.
echo ════════════════════════════════════════════════════════════════════════════
echo.
echo Последние 10 строк лога:
echo ════════════════════════════════════════════════════════════════════════════
if exist logs\telegram_bot.log (
    powershell -Command "Get-Content logs\telegram_bot.log -Tail 10"
) else (
    echo [i] Лог-файл не найден
)

echo.
pause
