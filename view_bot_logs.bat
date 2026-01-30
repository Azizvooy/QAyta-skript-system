@echo off
chcp 65001 >nul
echo ================================================================================
echo ПРОСМОТР ЛОГОВ TELEGRAM БОТА
echo ================================================================================
echo.

if exist logs\telegram_bot.log (
    echo Последние 50 строк лога:
    echo ================================================================================
    echo.
    powershell -Command "Get-Content logs\telegram_bot.log -Tail 50"
    echo.
    echo ================================================================================
    echo.
    echo Полный лог находится в: logs\telegram_bot.log
    echo Размер файла:
    dir logs\telegram_bot.log | findstr telegram_bot.log
) else (
    echo [!] Лог-файл не найден
    echo.
    echo Лог создается автоматически при запуске бота
)

echo.
echo Нажмите любую клавишу для обновления...
pause >nul
cls
goto :EOF
