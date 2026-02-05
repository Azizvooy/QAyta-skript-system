@echo off
chcp 65001 > nul
echo ========================================
echo 🤖 ЗАПУСК TELEGRAM БОТА
echo ========================================
echo.

cd /d "%~dp0"

:: Проверка виртуального окружения
if not exist ".venv\Scripts\activate.bat" (
    echo ❌ Виртуальное окружение не найдено!
    echo 💡 Создайте его: python -m venv .venv
    echo.
    pause
    exit /b 1
)

echo ✅ Активация виртуального окружения...
call .venv\Scripts\activate.bat

echo.
echo 🚀 Запуск бота...
echo 📅 Время запуска: %date% %time%
echo.

:: Проверка наличия файла бота
if not exist "scripts\telegram\interactive_bot.py" (
    echo ❌ Файл бота не найден!
    echo.
    pause
    exit /b 1
)

echo 🔄 Запуск интерактивного бота с меню...
echo.
echo 💡 Бот будет работать в фоне с полным функционалом:
echo    - Меню с кнопками
echo    - Команды /start, /help, /stats
echo    - Отчеты по запросу
echo    - Обновление данных
echo.

:: Запуск бота в фоновом режиме
start "" /B pythonw.exe scripts\telegram\interactive_bot.py

timeout /t 3 /nobreak >nul

:: Проверка запуска
tasklist /FI "IMAGENAME eq pythonw.exe" 2>nul | find /I "pythonw.exe" >nul
if errorlevel 1 (
    echo ❌ Бот не запустился!
    echo.
    echo 💡 Попробуйте запустить вручную для просмотра ошибок:
    echo    python scripts\telegram\interactive_bot.py
    echo.
    pause
    exit /b 1
)

echo ✅ Интерактивный бот успешно запущен!
echo.
echo 📱 Откройте Telegram и напишите боту: /start
echo 💡 Для остановки: ОСТАНОВИТЬ_БОТ.bat
echo 📊 Для проверки: ПРОВЕРКА_СИСТЕМЫ.bat
echo.
timeout /t 2 /nobreak >nul
