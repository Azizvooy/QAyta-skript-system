#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Проверка статуса данных в PostgreSQL
"""
import psycopg2
from dotenv import load_dotenv
import os

# Загружаем конфигурацию
load_dotenv('config/postgresql.env')

try:
    # Подключение к БД
    conn = psycopg2.connect(
        host=os.getenv('DB_HOST'),
        port=os.getenv('DB_PORT'),
        dbname=os.getenv('DB_NAME'),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD')
    )
    cur = conn.cursor()
    
    # Общее количество записей
    cur.execute('SELECT COUNT(*) FROM fixations')
    total = cur.fetchone()[0]
    print(f'📊 Всего записей в БД: {total:,}')
    
    # По категориям
    cur.execute('SELECT status_category, COUNT(*) FROM fixations GROUP BY status_category ORDER BY COUNT(*) DESC')
    print('\n📈 Распределение по категориям:')
    for category, count in cur.fetchall():
        percentage = (count / total * 100) if total > 0 else 0
        print(f'  {category}: {count:,} ({percentage:.1f}%)')
    
    # Операторы
    cur.execute('SELECT COUNT(DISTINCT operator_id) FROM fixations')
    operators_count = cur.fetchone()[0]
    print(f'\n👥 Количество операторов: {operators_count}')
    
    # Топ-5 операторов
    cur.execute('''
        SELECT o.operator_name, COUNT(*) as cnt 
        FROM fixations f
        JOIN operators o ON f.operator_id = o.operator_id
        GROUP BY o.operator_name
        ORDER BY cnt DESC
        LIMIT 5
    ''')
    print('\n🏆 Топ-5 операторов:')
    for name, count in cur.fetchall():
        print(f'  {name}: {count:,}')
    
    # Диапазон дат
    cur.execute('SELECT MIN(call_date), MAX(call_date) FROM fixations WHERE call_date IS NOT NULL')
    min_date, max_date = cur.fetchone()
    print(f'\n📅 Диапазон дат: {min_date} - {max_date}')
    
    conn.close()
    print('\n✅ Данные успешно загружены и доступны!')
    
except Exception as e:
    print(f'❌ Ошибка: {e}')
