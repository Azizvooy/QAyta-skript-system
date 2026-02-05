@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo ========================================
echo 📊 ПРОВЕРКА СТАТУСА ВСЕЙ СИСТЕМЫ
echo ========================================
echo.

echo [1] ПРОВЕРКА ПЛАНИРОВЩИКА ЗАДАЧ
echo ----------------------------------------
echo.
echo 🤖 Telegram Bot:
schtasks /Query /TN "TelegramBot_AutoStart" 2>nul
if %errorLevel% EQU 0 (
    echo ✅ Автозапуск бота: ВКЛЮЧЕН
) else (
    echo ❌ Автозапуск бота: НЕ НАСТРОЕН
)
echo.

echo 🔄 Обновление данных:
schtasks /Query /TN "GoogleSheets_HourlyUpdate" 2>nul
if %errorLevel% EQU 0 (
    echo ✅ Автообновление каждый час: ВКЛЮЧЕНО
) else (
    echo ❌ Автообновление: НЕ НАСТРОЕНО
)
echo.

echo 📈 Генерация отчетов:
schtasks /Query /TN "Reports_DailyGeneration" 2>nul
if %errorLevel% EQU 0 (
    echo ✅ Ежедневные отчеты: ВКЛЮЧЕНЫ
) else (
    echo ❌ Ежедневные отчеты: НЕ НАСТРОЕНЫ
)
echo.

echo [2] ПРОВЕРКА ПРОЦЕССОВ PYTHON
echo ----------------------------------------
tasklist /FI "IMAGENAME eq python.exe" 2>nul | find /I "python.exe" >nul
if %errorLevel% EQU 0 (
    echo ✅ Python процессы запущены:
    tasklist /FI "IMAGENAME eq python.exe"
) else (
    echo ⚠️ Python процессы не найдены
)
echo.

echo [3] ПРОВЕРКА КОНФИГУРАЦИИ
echo ----------------------------------------
if exist "telegram_config.txt" (
    echo ✅ Telegram конфиг: НАЙДЕН
) else (
    echo ❌ Telegram конфиг: НЕ НАЙДЕН
)

if exist "config\credentials.json" (
    echo ✅ Google credentials: НАЙДЕНЫ
) else (
    echo ❌ Google credentials: НЕ НАЙДЕНЫ
    echo    👉 Требуется настройка Google API
)

if exist "config\token.json" (
    echo ✅ Google token: НАЙДЕН
) else (
    echo ⚠️ Google token: НЕ НАЙДЕН (создастся после первой авторизации)
)

if exist "data\fiksa_database.db" (
    echo ✅ База данных: НАЙДЕНА
) else (
    echo ❌ База данных: НЕ НАЙДЕНА
)
echo.

echo [4] ПРОВЕРКА ВИРТУАЛЬНОГО ОКРУЖЕНИЯ
echo ----------------------------------------
if exist ".venv\Scripts\python.exe" (
    echo ✅ Виртуальное окружение: НАСТРОЕНО
    ".venv\Scripts\python.exe" --version
) else (
    echo ❌ Виртуальное окружение: НЕ НАЙДЕНО
)
echo.

echo [5] ПОСЛЕДНИЕ ЛОГИ (если есть)
echo ----------------------------------------
if exist "data\hourly_update.log" (
    echo 📄 Последние 10 строк из hourly_update.log:
    powershell -Command "Get-Content 'data\hourly_update.log' -Tail 10 -Encoding UTF8"
) else (
    echo ⚠️ Лог файл не найден
)
echo.

echo ========================================
echo 💡 РЕКОМЕНДАЦИИ
echo ========================================
echo.
if not exist "config\credentials.json" (
    echo ⚠️ Настройте Google Sheets API:
    echo    1. Создайте проект: https://console.cloud.google.com/
    echo    2. Включите Google Sheets API
    echo    3. Скачайте credentials.json в папку config\
    echo.
)

schtasks /Query /TN "TelegramBot_AutoStart" 2>nul >nul
if %errorLevel% NEQ 0 (
    echo ⚠️ Запустите ПОЛНАЯ_НАСТРОЙКА_СИСТЕМЫ.bat для настройки
    echo.
)

echo ========================================
pause
