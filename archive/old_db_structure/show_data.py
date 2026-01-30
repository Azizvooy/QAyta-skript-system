#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Показать все доступные данные в базе данных
"""

import sqlite3
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'

def show_available_data():
    """Показать все доступные данные"""
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('СПИСОК ВСЕХ ДОСТУПНЫХ ДАННЫХ В БАЗЕ')
    print('=' * 80)
    
    # 1. ТАБЛИЦА APPLICATIONS
    print('\n📋 ТАБЛИЦА: applications (Заявки)')
    print('-' * 80)
    
    apps_sample = pd.read_sql_query('SELECT * FROM applications LIMIT 3', conn)
    print(f'Всего записей: {pd.read_sql_query("SELECT COUNT(*) as cnt FROM applications", conn)["cnt"][0]}')
    print('\nПоля:')
    for col in apps_sample.columns:
        print(f'  - {col}')
    print('\nПример данных:')
    print(apps_sample.to_string(index=False))
    
    # 2. ТАБЛИЦА FIKSA_RECORDS
    print('\n\n📋 ТАБЛИЦА: fiksa_records (Данные фиксации)')
    print('-' * 80)
    
    fiksa_sample = pd.read_sql_query('SELECT * FROM fiksa_records LIMIT 3', conn)
    print(f'Всего записей: {pd.read_sql_query("SELECT COUNT(*) as cnt FROM fiksa_records", conn)["cnt"][0]}')
    print('\nПоля:')
    for col in fiksa_sample.columns:
        print(f'  - {col}')
    print('\nПример данных:')
    print(fiksa_sample.to_string(index=False))
    
    # 3. ТАБЛИЦА OPERATOR_STATS
    print('\n\n📋 ТАБЛИЦА: operator_stats (Статистика операторов)')
    print('-' * 80)
    
    stats_sample = pd.read_sql_query('SELECT * FROM operator_stats LIMIT 3', conn)
    print(f'Всего записей: {pd.read_sql_query("SELECT COUNT(*) as cnt FROM operator_stats", conn)["cnt"][0]}')
    print('\nПоля:')
    for col in stats_sample.columns:
        print(f'  - {col}')
    print('\nПример данных:')
    print(stats_sample.to_string(index=False))
    
    # 4. СТАТУСЫ
    print('\n\n📊 УНИКАЛЬНЫЕ СТАТУСЫ:')
    print('-' * 80)
    statuses = pd.read_sql_query('SELECT DISTINCT status FROM fiksa_records WHERE status IS NOT NULL', conn)
    for i, status in enumerate(statuses['status'], 1):
        count = pd.read_sql_query(f"SELECT COUNT(*) as cnt FROM fiksa_records WHERE status = '{status}'", conn)['cnt'][0]
        print(f'{i:2d}. {status:<50} ({count:>6} записей)')
    
    # 5. РЕГИОНЫ
    print('\n\n🌍 УНИКАЛЬНЫЕ РЕГИОНЫ (из адресов):')
    print('-' * 80)
    addresses = pd.read_sql_query('SELECT DISTINCT address FROM applications WHERE address IS NOT NULL LIMIT 100', conn)
    regions = set()
    for addr in addresses['address']:
        parts = str(addr).split(',')
        if parts:
            regions.add(parts[0].strip())
    
    for i, region in enumerate(sorted(regions), 1):
        print(f'{i:2d}. {region}')
    
    # 6. ОПЕРАТОРЫ
    print('\n\n👥 ОПЕРАТОРЫ:')
    print('-' * 80)
    operators = pd.read_sql_query('''
        SELECT operator_name, COUNT(*) as total_calls
        FROM fiksa_records 
        WHERE operator_name IS NOT NULL AND operator_name != ''
        GROUP BY operator_name
        ORDER BY total_calls DESC
    ''', conn)
    print(operators.to_string(index=False))
    
    # 7. СВОДНАЯ ИНФОРМАЦИЯ
    print('\n\n📈 СВОДНАЯ ИНФОРМАЦИЯ:')
    print('-' * 80)
    
    # Всего заявок
    total_apps = pd.read_sql_query('SELECT COUNT(*) as cnt FROM applications', conn)['cnt'][0]
    print(f'Всего заявок: {total_apps:,}')
    
    # Всего записей фиксации
    total_fiksa = pd.read_sql_query('SELECT COUNT(*) as cnt FROM fiksa_records', conn)['cnt'][0]
    print(f'Всего записей фиксации: {total_fiksa:,}')
    
    # Заявки с адресами
    apps_with_addr = pd.read_sql_query('SELECT COUNT(*) as cnt FROM applications WHERE address IS NOT NULL', conn)['cnt'][0]
    print(f'Заявок с адресами: {apps_with_addr:,}')
    
    # Заявки с телефонами
    apps_with_phone = pd.read_sql_query('SELECT COUNT(*) as cnt FROM applications WHERE phone IS NOT NULL', conn)['cnt'][0]
    print(f'Заявок с телефонами: {apps_with_phone:,}')
    
    # Период данных
    date_range = pd.read_sql_query('SELECT MIN(import_date) as min_date, MAX(import_date) as max_date FROM applications', conn)
    print(f'Период заявок: {date_range["min_date"][0]} - {date_range["max_date"][0]}')
    
    fiksa_date_range = pd.read_sql_query('SELECT MIN(call_date) as min_date, MAX(call_date) as max_date FROM fiksa_records', conn)
    print(f'Период фиксации: {fiksa_date_range["min_date"][0]} - {fiksa_date_range["max_date"][0]}')
    
    # 8. ОБЪЕДИНЕННЫЕ ДАННЫЕ (JOIN)
    print('\n\n🔗 ПРИМЕР ОБЪЕДИНЕННЫХ ДАННЫХ (applications + fiksa_records):')
    print('-' * 80)
    
    joined = pd.read_sql_query('''
        SELECT 
            a.application_number,
            a.address,
            a.phone,
            a.notes as complaint,
            f.operator_name,
            f.status,
            f.call_date
        FROM applications a
        LEFT JOIN fiksa_records f ON (
            f.full_name = a.application_number 
            OR f.phone LIKE '%' || REPLACE(REPLACE(a.phone, '+998', ''), '+', '') || '%'
        )
        LIMIT 5
    ''', conn)
    
    print('Поля объединенных данных:')
    for col in joined.columns:
        print(f'  - {col}')
    
    print('\nПример:')
    print(joined.to_string(index=False))
    
    conn.close()
    
    print('\n' + '=' * 80)
    print('ДАННЫЕ ДЛЯ ОТЧЕТНОСТИ:')
    print('=' * 80)
    print('''
📊 ДОСТУПНЫЕ ПОЛЯ ДЛЯ ОТЧЕТОВ:

ИЗ ЗАЯВОК (applications):
  • application_number - Номер заявки (например: 01.AAC4685/26)
  • phone - Телефон
  • address - Полный адрес (Область, Район, улица...)
  • notes - Описание жалобы/проблемы
  • import_date - Дата импорта заявки

ИЗ ФИКСАЦИИ (fiksa_records):
  • operator_name - Имя оператора
  • status - Статус звонка (Положительный, Отрицательный, и т.д.)
  • call_date - Дата звонка
  • card_number - Номер карты
  • full_name - Полное имя (тут номер заявки)
  • notes - Примечания оператора

ВЫЧИСЛЯЕМЫЕ ПОЛЯ:
  • Область - Извлекается из address (первая часть до запятой)
  • Район - Извлекается из address (вторая часть)
  • Улица - Извлекается из address (третья часть)

ГРУППИРОВКИ:
  • По регионам (Область)
  • По районам (Область + Район)
  • По статусам
  • По операторам
  • По датам (день/неделя/месяц)
  • По типам жалоб

СТАТИСТИКА:
  • Количество заявок
  • Количество звонков
  • Процент обработки
  • Распределение по статусам
  • Топ операторов
  • Топ регионов
    ''')

def get_db_connection():
    return sqlite3.connect(DB_PATH)

if __name__ == '__main__':
    show_available_data()
