#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).parent.parent.parent / 'data' / 'fiksa_database.db'

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

print('\n' + '='*60)
print('СТАТИСТИКА БАЗЫ ДАННЫХ')
print('='*60)

tables = ['operators', 'services', 'regions', 'applications', 'fixations']

for table in tables:
    count = cursor.execute(f'SELECT COUNT(*) FROM {table}').fetchone()[0]
    print(f'{table:20s}: {count:,}')

print('='*60)

# Топ-10 операторов
print('\nТоп-10 операторов по количеству фиксаций:')
cursor.execute('''
    SELECT o.operator_name, COUNT(f.fixation_id) as cnt
    FROM operators o
    LEFT JOIN fixations f ON o.operator_id = f.operator_id
    GROUP BY o.operator_id
    ORDER BY cnt DESC
    LIMIT 10
''')
for row in cursor.fetchall():
    print(f'  {row[0]}: {row[1]:,}')

conn.close()
