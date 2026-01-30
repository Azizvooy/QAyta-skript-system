@echo off
chcp 65001 >nul
title 🗑️ Очистка и обновление данных

cd /d "%~dp0"

echo.
echo ════════════════════════════════════════════════════════════════════
echo     🗑️ ОЧИСТКА И ОБНОВЛЕНИЕ ДАННЫХ
echo ════════════════════════════════════════════════════════════════════
echo.

echo [1/4] Удаление старых данных из БД...
python -c "import sqlite3; conn = sqlite3.connect('data/fiksa_database.db'); conn.execute('DELETE FROM fiksa_records'); conn.execute('DELETE FROM call_history_112'); conn.commit(); print('[OK] БД очищена'); conn.close()"

echo.
echo [2/4] Поиск файлов для импорта...
python -c "from pathlib import Path; files = list(Path('data').glob('*.csv')) + list(Path('incoming_data').glob('**/*.csv')) + list(Path('incoming_data').glob('**/*.xlsx')); print(f'Найдено {len(files)} файлов'); [print(f'  - {f.name}') for f in files[:5]]; print(f'  ... и еще {len(files)-5}' if len(files) > 5 else '')"

echo.
echo [3/4] Импорт данных...
python scripts\database\db_import.py

echo.
echo [4/4] Генерация новых отчетов...
python scripts\analysis\optimized_reports.py

echo.
echo ════════════════════════════════════════════════════════════════════
echo [OK] Данные обновлены, отчеты созданы!
echo ════════════════════════════════════════════════════════════════════
echo.
pause
