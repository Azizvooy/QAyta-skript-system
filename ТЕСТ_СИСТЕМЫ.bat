@echo off
chcp 65001 > nul
echo ========================================
echo 🧪 ТЕСТ ВСЕХ КОМПОНЕНТОВ СИСТЕМЫ
echo ========================================
echo.

cd /d "%~dp0"

:: Активация виртуального окружения
if not exist ".venv\Scripts\activate.bat" (
    echo ❌ Виртуальное окружение не найдено!
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
echo ✅ Виртуальное окружение активировано
echo.

echo [ТЕСТ 1/4] Проверка зависимостей Python
echo ----------------------------------------
python -c "import pandas; import openpyxl; import google.oauth2; print('✅ Все библиотеки установлены')" 2>nul
if %errorLevel% NEQ 0 (
    echo ❌ Не все зависимости установлены
    echo 💡 Запустите: pip install -r requirements.txt
    pause
    exit /b 1
)
echo.

echo [ТЕСТ 2/4] Проверка подключения к Google Sheets
echo ----------------------------------------
if not exist "config\credentials.json" (
    echo ❌ credentials.json не найден!
    echo 💡 Настройте Google API (см. config\README.md)
    pause
    exit /b 1
) else (
    echo ✅ credentials.json найден
    echo 💡 Для полного теста запустите: update_from_sheets.py
)
echo.

echo [ТЕСТ 3/4] Проверка Telegram Bot
echo ----------------------------------------
if not exist "telegram_config.txt" (
    echo ❌ telegram_config.txt не найден!
    pause
    exit /b 1
)

echo ✅ Конфигурация найдена
echo 📞 Тестовая отправка сообщения...
python -c "import requests; r=requests.post('https://api.telegram.org/bot8141079204:AAErNrjLqTu4Vj1_7VS2kjGFKcR3lU9L9N4/sendMessage', data={'chat_id':'2012682567','text':'🧪 Тест системы: все работает!'}); print('✅ Сообщение отправлено!' if r.status_code==200 else '❌ Ошибка отправки')"
echo.

echo [ТЕСТ 4/4] Проверка базы данных
echo ----------------------------------------
if exist "data\fiksa_database.db" (
    echo ✅ База данных найдена
    python -c "import sqlite3; conn=sqlite3.connect('data/fiksa_database.db'); cursor=conn.cursor(); cursor.execute('SELECT COUNT(*) FROM fiksa_calls'); count=cursor.fetchone()[0]; print(f'📊 Записей в БД: {count}'); conn.close()"
) else (
    echo ⚠️ База данных не найдена (будет создана при первом импорте)
)
echo.

echo ========================================
echo ✅ ТЕСТИРОВАНИЕ ЗАВЕРШЕНО
echo ========================================
echo.
echo 💡 Следующие шаги:
echo    1. Если все тесты прошли - запустите ПОЛНАЯ_НАСТРОЙКА_СИСТЕМЫ.bat
echo    2. Проверьте Telegram - должно прийти тестовое сообщение
echo    3. Перезагрузите компьютер
echo.
pause
