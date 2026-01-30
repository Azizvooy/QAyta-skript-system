@echo off
chcp 65001 >nul
cls

:menu
echo.
echo ╔════════════════════════════════════════════════════════════════════════════╗
echo ║                        🤖 TELEGRAM БОТ FIKSA                              ║
echo ╚════════════════════════════════════════════════════════════════════════════╝
echo.
echo   1. ▶️  Запустить бота (фоновый режим)
echo   2. 🧪 Тестовый запуск (с логами в консоли)
echo   3. ⏹️  Остановить бота
echo   4. 🔍 Проверить статус
echo   5. 📋 Просмотреть логи
echo   6. 🔄 Перезапустить бота
echo   7. 🛑 Принудительная остановка (если не работает обычная)
echo   8. ❌ Выход
echo.
echo ════════════════════════════════════════════════════════════════════════════
echo.

set /p choice="Выберите действие (1-8): "

if "%choice%"=="1" goto start_bot
if "%choice%"=="2" goto test_bot
if "%choice%"=="3" goto stop_bot
if "%choice%"=="4" goto check_status
if "%choice%"=="5" goto view_logs
if "%choice%"=="6" goto restart_bot
if "%choice%"=="7" goto kill_all
if "%choice%"=="8" goto exit
goto menu

:start_bot
cls
call start_telegram_bot.bat
pause
goto menu

:test_bot
cls
call test_bot.bat
goto menu

:stop_bot
cls
call stop_telegram_bot.bat
pause
goto menu

:check_status
cls
call check_bot_status.bat
pause
goto menu

:view_logs
cls
call view_bot_logs.bat
pause
goto menu

:restart_bot
cls
call restart_telegram_bot.bat
echo.
echo Нажмите любую клавишу для возврата в меню...
pause >nul
cls
goto menu

:kill_all
cls
call kill_all_python.bat
cls
goto menu

:exit
cls
echo.
echo ════════════════════════════════════════════════════════════════════════════
echo.
echo 👋 До свидания!
echo.
echo ════════════════════════════════════════════════════════════════════════════
echo.
timeout /t 2 /nobreak >nul
exit
