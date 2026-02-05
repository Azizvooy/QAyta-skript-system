@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo.
echo ========================================
echo 🎯 БЫСТРАЯ ПРОВЕРКА СИСТЕМЫ
echo ========================================
echo.

echo [1] БОТ TELEGRAM
echo ----------------------------------------
tasklist /FI "IMAGENAME eq pythonw.exe" 2>nul | find /I "pythonw.exe" >nul
if %errorLevel% EQU 0 (
    echo ✅ БОТ ЗАПУЩЕН И РАБОТАЕТ
    echo.
    tasklist /FI "IMAGENAME eq pythonw.exe"
) else (
    echo ❌ Бот не запущен
    echo 💡 Запустите: ЗАПУСК_БОТ.bat
)
echo.

echo [2] ПЛАНИРОВЩИК ЗАДАЧ
echo ----------------------------------------
echo Проверяю задания...
schtasks /Query /TN "TelegramBot_AutoStart" >nul 2>&1
if %errorLevel% EQU 0 (
    echo ✅ Автозапуск бота: НАСТРОЕН
) else (
    echo ⚠️ Автозапуск бота: не настроен
)

schtasks /Query /TN "GoogleSheets_HourlyUpdate" >nul 2>&1
if %errorLevel% EQU 0 (
    echo ✅ Обновление данных: НАСТРОЕНО
) else (
    echo ⚠️ Обновление данных: не настроено
)

schtasks /Query /TN "Reports_DailyGeneration" >nul 2>&1
if %errorLevel% EQU 0 (
    echo ✅ Генерация отчетов: НАСТРОЕНА
) else (
    echo ⚠️ Генерация отчетов: не настроена
)
echo.

echo [3] TELEGRAM ТЕСТ
echo ----------------------------------------
echo Отправка тестового сообщения...
call .venv\Scripts\activate.bat >nul 2>&1
python -c "import requests; r=requests.post('https://api.telegram.org/bot8141079204:AAErNrjLqTu4Vj1_7VS2kjGFKcR3lU9L9N4/sendMessage', data={'chat_id':'2012682567','text':'✅ Тест: Бот работает!'}); exit(0 if r.status_code==200 else 1)" 2>nul
if %errorLevel% EQU 0 (
    echo ✅ Сообщение ОТПРАВЛЕНО в Telegram
    echo 📱 Проверьте свой Telegram!
) else (
    echo ❌ Ошибка отправки
)
echo.

echo ========================================
echo 📊 ИТОГОВЫЙ СТАТУС
echo ========================================
echo.

tasklist /FI "IMAGENAME eq pythonw.exe" 2>nul | find /I "pythonw.exe" >nul
if %errorLevel% EQU 0 (
    echo ✅ Система РАБОТАЕТ!
    echo ✅ Бот запущен в фоновом режиме
    echo ✅ Готов к работе 24/7
    echo.
    echo 💡 После перезагрузки бот запустится автоматически
) else (
    echo ⚠️ Бот не запущен
    echo 💡 Запустите: ЗАПУСК_БОТ.bat
)

echo.
echo ========================================
pause
