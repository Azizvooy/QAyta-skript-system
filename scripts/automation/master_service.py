#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ГЛАВНЫЙ ФОНОВЫЙ СЕРВИС - ПОЛНАЯ АВТОМАТИЗАЦИЯ ПРОЕКТА
=============================================================================
Автоматически выполняет:
1. Сбор данных из Google Sheets (только со статусом в колонке E)
2. Импорт истории звонков 112
3. Расчет статистики операторов
4. Анализ фидбэков от служб
5. Генерация всех отчетов
6. Отправка в Telegram
=============================================================================
"""

import schedule
import time
import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent.parent

# =============================================================================
# ЗАДАЧИ
# =============================================================================

def run_script(script_path, description):
    """Запустить скрипт Python"""
    print(f'\n{"=" * 80}')
    print(f'[{datetime.now().strftime("%H:%M:%S")}] {description}')
    print(f'{"=" * 80}')
    
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            cwd=str(BASE_DIR)
        )
        
        if result.returncode == 0:
            print(f'[OK] {description} - успешно')
            if result.stdout:
                print(result.stdout[-500:])  # Последние 500 символов
            return True
        else:
            print(f'[ОШИБКА] {description}')
            if result.stderr:
                print(result.stderr[-500:])
            return False
            
    except Exception as e:
        print(f'[ОШИБКА] {description}: {e}')
        return False

# =============================================================================
# ЕЖЕДНЕВНЫЕ ЗАДАЧИ (09:00)
# =============================================================================

def daily_tasks():
    """Ежедневные задачи в 09:00"""
    print('\n' + '=' * 80)
    print(f'ЕЖЕДНЕВНЫЕ ЗАДАЧИ - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print('=' * 80)
    
    tasks = [
        # 1. Сбор данных FIKSA из Google Sheets
        {
            'script': BASE_DIR / 'scripts' / 'data_collection' / 'sheets_to_db_collector.py',
            'description': '1. Сбор данных FIKSA (только со статусом)'
        },
        
        # 2. Импорт истории звонков 112
        {
            'script': BASE_DIR / 'scripts' / 'automation' / 'import_call_history.py',
            'description': '2. Импорт истории звонков 112'
        },
        
        # 3. Обработка заявок
        {
            'script': BASE_DIR / 'scripts' / 'automation' / 'process_applications.py',
            'description': '3. Обработка заявок граждан'
        },
        
        # 4. Генерация отчетов по службам
        {
            'script': BASE_DIR / 'scripts' / 'analysis' / 'service_reports.py',
            'description': '4. Генерация отчетов по службам (61 файл)'
        },
        
        # 5. Генерация отчетов по регионам
        {
            'script': BASE_DIR / 'scripts' / 'analysis' / 'auto_report.py',
            'description': '5. Генерация отчетов по регионам (15 файлов)'
        },
    ]
    
    success_count = 0
    for task in tasks:
        if run_script(task['script'], task['description']):
            success_count += 1
    
    print(f'\n[ИТОГО] Выполнено успешно: {success_count} из {len(tasks)} задач')

# =============================================================================
# ПОЧАСОВЫЕ ЗАДАЧИ
# =============================================================================

def hourly_tasks():
    """Почасовые задачи - легкая синхронизация"""
    print(f'\n[{datetime.now().strftime("%H:%M:%S")}] Почасовая синхронизация...')
    
    # Быстрый сбор новых данных FIKSA
    run_script(
        BASE_DIR / 'scripts' / 'data_collection' / 'sheets_to_db_collector.py',
        'Синхронизация данных FIKSA'
    )

# =============================================================================
# СТАТИСТИКА В РЕАЛЬНОМ ВРЕМЕНИ
# =============================================================================

def show_current_stats():
    """Показать текущую статистику"""
    import sqlite3
    import pandas as pd
    
    db_path = BASE_DIR / 'data' / 'fiksa_database.db'
    if not db_path.exists():
        return
    
    conn = sqlite3.connect(db_path)
    
    # Статистика по таблицам
    tables = ['fiksa_records', 'call_history_112', 'applications', 'operator_stats']
    
    print('\n' + '=' * 80)
    print(f'ТЕКУЩАЯ СТАТИСТИКА - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print('=' * 80)
    
    for table in tables:
        try:
            count = pd.read_sql_query(f'SELECT COUNT(*) as count FROM {table}', conn).iloc[0]['count']
            print(f'{table:30} {count:>10,} записей')
        except:
            pass
    
    # Статистика операторов за сегодня
    try:
        today_stats = pd.read_sql_query('''
            SELECT 
                COUNT(DISTINCT operator_name) as operators,
                COUNT(*) as total,
                COUNT(CASE WHEN status LIKE '%Положительн%' THEN 1 END) as positive,
                COUNT(CASE WHEN status LIKE '%Отрицательн%' THEN 1 END) as negative
            FROM fiksa_records
            WHERE collection_date = date('now')
        ''', conn)
        
        if not today_stats.empty:
            row = today_stats.iloc[0]
            print(f'\n[СЕГОДНЯ]')
            print(f'  Операторов: {row["operators"]}')
            print(f'  Записей: {row["total"]:,}')
            print(f'  Положительных: {row["positive"]:,}')
            print(f'  Отрицательных: {row["negative"]:,}')
    except:
        pass
    
    conn.close()
    print('=' * 80)

# =============================================================================
# РАСПИСАНИЕ
# =============================================================================

def setup_schedule():
    """Настроить расписание задач"""
    print('\n' + '=' * 80)
    print('НАСТРОЙКА РАСПИСАНИЯ ФОНОВОГО СЕРВИСА')
    print('=' * 80)
    
    # Ежедневные задачи в 09:00
    schedule.every().day.at("09:00").do(daily_tasks)
    print('[OK] Ежедневные задачи: 09:00')
    
    # Почасовые задачи
    schedule.every().hour.at(":00").do(hourly_tasks)
    print('[OK] Почасовая синхронизация: каждый час')
    
    # Статистика каждые 30 минут
    schedule.every(30).minutes.do(show_current_stats)
    print('[OK] Показ статистики: каждые 30 минут')
    
    print('\n' + '=' * 80)
    print('СЕРВИС ЗАПУЩЕН И РАБОТАЕТ В ФОНОВОМ РЕЖИМЕ')
    print('=' * 80)
    print(f'Время запуска: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print('\nНажмите Ctrl+C для остановки')
    print('=' * 80)
    
    # Показать начальную статистику
    show_current_stats()

# =============================================================================
# ГЛАВНЫЙ ЦИКЛ
# =============================================================================

def main():
    """Главная функция"""
    setup_schedule()
    
    # Бесконечный цикл
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Проверка каждую минуту
            
    except KeyboardInterrupt:
        print('\n\n' + '=' * 80)
        print('СЕРВИС ОСТАНОВЛЕН ПОЛЬЗОВАТЕЛЕМ')
        print('=' * 80)
        print(f'Время остановки: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')

if __name__ == '__main__':
    main()
