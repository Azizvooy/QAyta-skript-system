@echo off
chcp 65001 >nul
echo ================================================================================
echo ОСТАНОВКА TELEGRAM БОТА
echo ================================================================================
echo.

set stopped=0

REM Метод 1: Поиск по заголовку окна (если запущен с окном)
echo [1/4] Поиск процессов с окном...
for /f "tokens=2" %%a in ('tasklist /FI "WINDOWTITLE eq Telegram Bot FIKSA" /FO LIST 2^>nul ^| find "PID:"') do (
    echo Остановка процесса PID: %%a
    taskkill /PID %%a /F /T >nul 2>&1
    set stopped=1
)

REM Метод 2: Поиск по имени скрипта через PowerShell
echo [2/4] Поиск процессов working_bot.py...
for /f "tokens=1" %%a in ('powershell -Command "Get-WmiObject Win32_Process | Where-Object { $_.CommandLine -like '*working_bot.py*' } | Select-Object -ExpandProperty ProcessId"') do (
    echo Остановка процесса PID: %%a
    taskkill /PID %%a /F /T >nul 2>&1
    set stopped=1
)

REM Метод 3: Поиск старых процессов interactive_bot.py
echo [3/4] Поиск старых процессов interactive_bot.py...
for /f "tokens=1" %%a in ('powershell -Command "Get-WmiObject Win32_Process | Where-Object { $_.CommandLine -like '*interactive_bot.py*' } | Select-Object -ExpandProperty ProcessId"') do (
    echo Остановка процесса PID: %%a
    taskkill /PID %%a /F /T >nul 2>&1
    set stopped=1
)

REM Ждем завершения
timeout /t 1 /nobreak >nul

REM Проверка
echo [4/4] Проверка...
for /f "tokens=1" %%a in ('powershell -Command "Get-WmiObject Win32_Process | Where-Object { $_.CommandLine -like '*working_bot.py*' -or $_.CommandLine -like '*interactive_bot.py*' } | Measure-Object | Select-Object -ExpandProperty Count"') do (
    if %%a GTR 0 (
        echo [!] Найдено процессов: %%a
        echo [i] Попробуйте: kill_all_python.bat
    ) else (
        echo [OK] Все процессы бота остановлены!
    )
)

echo.
pause
