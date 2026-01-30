@echo off
chcp 65001 >nul
cls

echo ╔════════════════════════════════════════════════════════════════════════════╗
echo ║              🧹 ОЧИСТКА СТАРЫХ ФАЙЛОВ ПРОЕКТА                             ║
echo ╚════════════════════════════════════════════════════════════════════════════╝
echo.
echo Этот скрипт удалит старые неиспользуемые файлы:
echo.
echo   • Старые CSV файлы (ALL_DATA*.csv, UNIQUE_CARDS*.csv)
echo   • Старые Excel файлы (QAYTA_ALOQA_REPORT*.xlsx)
echo   • Старые Word документы (*.docx)
echo   • Старые README файлы
echo   • Старые отчеты и логи
echo.
echo ⚠️  ВНИМАНИЕ: Файлы будут удалены безвозвратно!
echo.
echo Архивная папка archive\ НЕ будет удалена (можно удалить вручную позже).
echo.
echo ════════════════════════════════════════════════════════════════════════════
echo.

set /p confirm="Продолжить? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo.
    echo ❌ Отменено пользователем
    pause
    exit /b
)

echo.
echo 🧹 Начинаю очистку...
echo.

REM Старые CSV файлы
echo [1/6] Удаление старых CSV файлов...
if exist "ALL_DATA*.csv" (
    del /Q "ALL_DATA*.csv" 2>nul
    echo   ✓ ALL_DATA*.csv удалены
) else (
    echo   ℹ Файлы не найдены
)

if exist "UNIQUE_CARDS*.csv" (
    del /Q "UNIQUE_CARDS*.csv" 2>nul
    echo   ✓ UNIQUE_CARDS*.csv удалены
) else (
    echo   ℹ Файлы не найдены
)

REM Старые Excel файлы
echo.
echo [2/6] Удаление старых Excel файлов...
if exist "QAYTA_ALOQA_REPORT*.xlsx" (
    del /Q "QAYTA_ALOQA_REPORT*.xlsx" 2>nul
    echo   ✓ QAYTA_ALOQA_REPORT*.xlsx удалены
) else (
    echo   ℹ Файлы не найдены
)

REM Старые Word документы
echo.
echo [3/6] Удаление старых Word документов...
for %%f in (
    "ИТОГОВЫЙ_ОТЧЕТ_ПО_ОБЗВОНУ.docx"
    "Отчет_по_обзвону_2025.docx"
    "Полный_отчет_по_обзвону.docx"
    "ФИНАЛЬНЫЙ_ОТЧЕТ.docx"
) do (
    if exist %%f (
        del /Q %%f 2>nul
        echo   ✓ %%f удален
    )
)

REM Старые README файлы
echo.
echo [4/6] Удаление старых README файлов...
for %%f in (
    "README_DIRECT_PYTHON.md"
    "README_DOCS_INTEGRATION.md"
    "README_СИСТЕМА_ГОТОВА.md"
    "NEW_SYSTEM.md"
    "ALL_SCRIPTS.md"
    "PROJECT_STRUCTURE.md"
    "QUICK_START.md"
    "START_SERVICE.md"
) do (
    if exist %%f (
        del /Q %%f 2>nul
        echo   ✓ %%f удален
    )
)

REM Старые отчеты и логи
echo.
echo [5/6] Удаление старых отчетов...
for %%f in (
    "ИСПРАВЛЕНИЯ.md"
    "ОТЛАДКА.md"
    "ПРОБЛЕМНЫЕ_СТРОКИ.txt"
    "ОТЧЕТ.html"
) do (
    if exist %%f (
        del /Q %%f 2>nul
        echo   ✓ %%f удален
    )
)

REM Старые JS и Apps Script файлы
echo.
echo [6/6] Удаление старых скриптов...
for %%f in (
    "google_sheets_collector.js"
    "google_sheets_complete.js"
    "docs_collector.gs"
    "OPERATORS_LIST.txt"
    "token.pickle"
) do (
    if exist %%f (
        del /Q %%f 2>nul
        echo   ✓ %%f удален
    )
)

echo.
echo ════════════════════════════════════════════════════════════════════════════
echo.
echo ✅ Очистка завершена!
echo.
echo 📊 Результат:
echo   • Удалены старые CSV файлы
echo   • Удалены старые Excel файлы
echo   • Удалены старые Word документы
echo   • Удалены старые README файлы
echo   • Удалены старые отчеты
echo.
echo 💡 Следующие шаги:
echo   1. Проверьте папку archive\ - можно удалить если файлы не нужны
echo   2. Проверьте папку exported_sheets\ - можно очистить
echo   3. Проверьте папку 123\ - можно удалить
echo.
echo 🎉 Проект очищен от старых файлов!
echo.
pause
