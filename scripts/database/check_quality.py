#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).parent.parent.parent / 'data' / 'fiksa_database.db'

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

print('\n' + '='*80)
print('ПРОВЕРКА КАЧЕСТВА ДАННЫХ')
print('='*80)

# 1. Общая статистика
print('\n1. Общая статистика:')
cursor.execute('SELECT COUNT(*) FROM applications')
print(f'   Заявок: {cursor.fetchone()[0]:,}')

cursor.execute('SELECT COUNT(*) FROM fixations')
print(f'   Фиксаций: {cursor.fetchone()[0]:,}')

# 2. Заявки с данными
print('\n2. Заявки с заполненными полями:')
cursor.execute('SELECT COUNT(*) FROM applications WHERE card_number IS NOT NULL')
print(f'   С номером карты: {cursor.fetchone()[0]:,}')

cursor.execute('SELECT COUNT(*) FROM applications WHERE caller_phone IS NOT NULL')
print(f'   С телефоном: {cursor.fetchone()[0]:,}')

cursor.execute('SELECT COUNT(*) FROM applications WHERE call_date IS NOT NULL')
print(f'   С датой звонка: {cursor.fetchone()[0]:,}')

cursor.execute('SELECT COUNT(*) FROM applications WHERE service_id IS NOT NULL')
print(f'   Со службой: {cursor.fetchone()[0]:,}')

# 3. Фиксации с данными
print('\n3. Фиксации с заполненными полями:')
cursor.execute('SELECT COUNT(*) FROM fixations WHERE application_id IS NOT NULL')
print(f'   Привязаны к заявке: {cursor.fetchone()[0]:,}')

cursor.execute('SELECT COUNT(*) FROM fixations WHERE fixation_date IS NOT NULL')
print(f'   С датой фиксации: {cursor.fetchone()[0]:,}')

cursor.execute('SELECT COUNT(*) FROM fixations WHERE phone_called IS NOT NULL')
print(f'   С телефоном: {cursor.fetchone()[0]:,}')

# 4. Примеры заявок с фиксациями
print('\n4. Примеры заявок с фиксациями:')
cursor.execute('''
    SELECT 
        a.card_number,
        a.caller_phone,
        a.call_date,
        COUNT(f.fixation_id) as fix_count
    FROM applications a
    LEFT JOIN fixations f ON a.application_id = f.application_id
    WHERE a.card_number IS NOT NULL
    GROUP BY a.application_id
    ORDER BY fix_count DESC
    LIMIT 5
''')

print(f'   {"Номер карты":<20} {"Телефон":<15} {"Дата":<12} Фиксаций')
print(f'   {"-"*20} {"-"*15} {"-"*12} {"-"*10}')
for row in cursor.fetchall():
    print(f'   {str(row[0]):<20} {str(row[1] or "-"):<15} {str(row[2] or "-"):<12} {row[3]}')

# 5. Примеры фиксаций
print('\n5. Примеры фиксаций с полными данными:')
cursor.execute('''
    SELECT 
        a.card_number,
        o.operator_name,
        f.fixation_date,
        f.status,
        f.phone_called
    FROM fixations f
    JOIN operators o ON f.operator_id = o.operator_id
    LEFT JOIN applications a ON f.application_id = a.application_id
    WHERE f.fixation_date IS NOT NULL
    LIMIT 5
''')

print(f'   {"Карта":<15} {"Оператор":<30} {"Дата фикс.":<20} {"Статус":<20}')
print(f'   {"-"*15} {"-"*30} {"-"*20} {"-"*20}')
for row in cursor.fetchall():
    print(f'   {str(row[0] or "-"):<15} {str(row[1]):<30} {str(row[2] or "-"):<20} {str(row[3]):<20}')

# 6. Распределение фиксаций по заявкам
print('\n6. Распределение фиксаций по заявкам:')
cursor.execute('''
    SELECT 
        CASE 
            WHEN cnt = 0 THEN '0 фиксаций'
            WHEN cnt = 1 THEN '1 фиксация'
            WHEN cnt BETWEEN 2 AND 5 THEN '2-5 фиксаций'
            WHEN cnt BETWEEN 6 AND 10 THEN '6-10 фиксаций'
            ELSE 'больше 10'
        END as range,
        COUNT(*) as apps
    FROM (
        SELECT a.application_id, COUNT(f.fixation_id) as cnt
        FROM applications a
        LEFT JOIN fixations f ON a.application_id = f.application_id
        GROUP BY a.application_id
    )
    GROUP BY range
    ORDER BY range
''')

for row in cursor.fetchall():
    print(f'   {row[0]:<20}: {row[1]:,} заявок')

conn.close()
print('\n' + '='*80)
