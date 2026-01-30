#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Быстрый импорт данных из ALL_DATA_COLLECTED.csv
"""

import sqlite3
import pandas as pd
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
CSV_FILE = BASE_DIR / 'data' / 'ALL_DATA_COLLECTED.csv'

EXCLUDE_PATTERNS = ['Тренды', 'Текущий месяц - Сводка', 'Предыдущий месяц - Сводка',
                     'СВОДКА СОТРУДНИКИ', 'Ноябрь 2025', 'Декабрь 2025', 'сводка', 'итого']

def should_exclude(operator_name):
    if not operator_name or pd.isna(operator_name):
        return True
    op_str = str(operator_name).strip().lower()
    if op_str in ['', '-', 'nan', 'none', 'null']:
        return True
    for pattern in EXCLUDE_PATTERNS:
        if pattern.lower() in op_str:
            return True
    return False

print('\n' + '='*80)
print('ИМПОРТ ДАННЫХ')
print('='*80)

# Загрузка
print('\n[1/4] Загрузка CSV...')
df = pd.read_csv(CSV_FILE, low_memory=False)
print(f'  Загружено: {len(df):,} строк')

# Подключение к БД
conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

# Импорт в fiksa_records
print('\n[2/4] Импорт в fiksa_records...')
imported = 0
skipped = 0

for idx, row in df.iterrows():
    try:
        # Колонки: 3=Дата, 4=Статус, 7=Оператор, 1=Карта, 8=ФИО, 2=Телефон, 9=Адрес, 6=Примечания
        operator_name = row.iloc[7] if len(row) > 7 else None
        
        # Проверка исключений
        if should_exclude(operator_name):
            skipped += 1
            continue
        
        call_date = row.iloc[3] if len(row) > 3 else None
        status = row.iloc[4] if len(row) > 4 else None
        card_number = row.iloc[1] if len(row) > 1 else None
        full_name = row.iloc[8] if len(row) > 8 else None
        phone = row.iloc[2] if len(row) > 2 else None
        address = row.iloc[9] if len(row) > 9 else None
        notes = row.iloc[6] if len(row) > 6 else None
        
        # Парсинг даты
        try:
            if pd.notna(call_date):
                call_date_parsed = pd.to_datetime(str(call_date)).strftime('%Y-%m-%d %H:%M:%S')
            else:
                call_date_parsed = None
        except:
            call_date_parsed = None
        
        cursor.execute('''
            INSERT INTO fiksa_records 
            (collection_date, operator_name, card_number, full_name, phone, address, status, call_date, notes, created_at)
            VALUES (date('now'), ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
        ''', (
            str(operator_name) if pd.notna(operator_name) else None,
            str(card_number) if pd.notna(card_number) else None,
            str(full_name) if pd.notna(full_name) else None,
            str(phone) if pd.notna(phone) else None,
            str(address) if pd.notna(address) else None,
            str(status) if pd.notna(status) else None,
            call_date_parsed,
            str(notes) if pd.notna(notes) else None
        ))
        
        imported += 1
        
        if imported % 10000 == 0:
            print(f'    Импортировано: {imported:,} | Пропущено: {skipped:,}')
            conn.commit()
    
    except Exception as e:
        if idx < 10:
            print(f'    Ошибка в строке {idx}: {e}')

conn.commit()

print(f'\n  [OK] Импортировано: {imported:,}')
print(f'  [SKIP] Пропущено: {skipped:,}')

# Статистика
print('\n[3/4] Статистика операторов...')
stats = cursor.execute('''
    SELECT 
        operator_name,
        COUNT(*) as total,
        COUNT(CASE WHEN status LIKE '%Положительн%' THEN 1 END) as positive
    FROM fiksa_records
    WHERE operator_name IS NOT NULL
    GROUP BY operator_name
    ORDER BY total DESC
    LIMIT 10
''').fetchall()

print('  ТОП-10 операторов:')
for op, total, positive in stats:
    print(f'    {op[:40]:40} | {total:6,} звонков | {positive:5,} положит.')

# Финал
print('\n[4/4] Сохранение...')
conn.commit()
conn.close()

print('\n' + '='*80)
print('[OK] ИМПОРТ ЗАВЕРШЕН')
print('='*80)
