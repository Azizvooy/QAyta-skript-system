@echo off
chcp 65001 > nul
echo ========================================
echo 🚀 ПОЛНАЯ НАСТРОЙКА СИСТЕМЫ
echo ========================================
echo.
echo Этот скрипт настроит:
echo ✅ 1. Автозапуск бота при старте Windows
echo ✅ 2. Автообновление данных каждый час
echo ✅ 3. Автогенерацию отчетов
echo ✅ 4. Планировщик задач Windows
echo.
pause
echo.

:: Проверка прав администратора
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo ❌ ОШИБКА: Требуются права администратора!
    echo.
    echo 👉 Запустите этот файл ПКМ → "Запуск от имени администратора"
    echo.
    pause
    exit /b 1
)

echo ✅ Права администратора подтверждены
echo.

set "PROJECT_PATH=%~dp0"
cd /d "%PROJECT_PATH%"

echo ========================================
echo 📝 СОЗДАНИЕ ЗАДАНИЙ В ПЛАНИРОВЩИКЕ
echo ========================================
echo.

:: 1. Автозапуск бота при входе в систему
echo [1/3] Создание задания: Автозапуск бота...
schtasks /Create /TN "TelegramBot_AutoStart" /TR "\"%PROJECT_PATH%ЗАПУСК_БОТ.bat\"" /SC ONLOGON /RL HIGHEST /F
if %errorLevel% EQU 0 (
    echo ✅ Задание создано: TelegramBot_AutoStart
) else (
    echo ❌ Ошибка создания задания бота
)
echo.

:: 2. Обновление данных каждый час
echo [2/3] Создание задания: Обновление данных каждый час...
schtasks /Create /TN "GoogleSheets_HourlyUpdate" /TR "\"%PROJECT_PATH%.venv\Scripts\python.exe\" \"%PROJECT_PATH%scripts\data_collection\sheets_to_db_collector.py\"" /SC HOURLY /RL HIGHEST /F
if %errorLevel% EQU 0 (
    echo ✅ Задание создано: GoogleSheets_HourlyUpdate
) else (
    echo ❌ Ошибка создания задания обновления
)
echo.

:: 3. Генерация отчетов каждый день в 9:00
echo [3/3] Создание задания: Генерация отчетов каждый день в 9:00...
schtasks /Create /TN "Reports_DailyGeneration" /TR "\"%PROJECT_PATH%.venv\Scripts\python.exe\" \"%PROJECT_PATH%generate_full_report.py\"" /SC DAILY /ST 09:00 /RL HIGHEST /F
if %errorLevel% EQU 0 (
    echo ✅ Задание создано: Reports_DailyGeneration
) else (
    echo ❌ Ошибка создания задания отчетов
)
echo.

echo ========================================
echo 📊 ПРОСМОТР СОЗДАННЫХ ЗАДАНИЙ
echo ========================================
echo.
schtasks /Query /TN "TelegramBot_AutoStart" /FO LIST
echo.
schtasks /Query /TN "GoogleSheets_HourlyUpdate" /FO LIST
echo.
schtasks /Query /TN "Reports_DailyGeneration" /FO LIST
echo.

echo ========================================
echo ✅ НАСТРОЙКА ЗАВЕРШЕНА!
echo ========================================
echo.
echo 🔹 Созданные задания:
echo    1. TelegramBot_AutoStart - бот запускается при входе
echo    2. GoogleSheets_HourlyUpdate - обновление каждый час
echo    3. Reports_DailyGeneration - отчеты каждый день в 9:00
echo.
echo 💡 Что дальше:
echo    1. Настройте Google Sheets API (config\credentials.json)
echo    2. Перезагрузите компьютер
echo    3. Проверьте работу: check_bot_status.bat
echo.
echo 📖 Подробная инструкция: ЛОКАЛЬНЫЙ_ЗАПУСК_НА_ПК.md
echo.
pause
