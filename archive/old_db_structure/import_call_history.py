#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт истории звонков 112 с данными по службам и инцидентам
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import re

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
IMPORT_DIR = BASE_DIR / '123'

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def create_call_history_table():
    """Создать таблицу для истории звонков 112"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS call_history_112 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            call_date DATE,
            call_time TEXT,
            duration TEXT,
            transfer_time TEXT,
            close_time TEXT,
            service_code TEXT,
            service_name TEXT,
            reason TEXT,
            card_number TEXT,
            incident_number TEXT,
            operator_name TEXT,
            status TEXT,
            caller_name TEXT,
            region TEXT,
            district TEXT,
            address TEXT,
            location_type TEXT,
            description TEXT,
            self_refusal TEXT,
            caller_phone TEXT,
            import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            source_file TEXT
        )
    ''')
    
    conn.commit()
    conn.close()
    print('✅ Таблица call_history_112 создана')

def import_call_history_files():
    """Импортировать все файлы истории звонков"""
    conn = get_db_connection()
    
    # Найти все Excel файлы с историей звонков
    excel_files = list(IMPORT_DIR.glob('История*.xlsx'))
    
    if not excel_files:
        print('❌ Нет файлов для импорта в папке 123/')
        return
    
    total_imported = 0
    
    for file_path in excel_files:
        print(f'\n📂 Импорт: {file_path.name}')
        
        try:
            # Прочитать Excel
            df = pd.read_excel(file_path)
            
            print(f'   Найдено записей: {len(df)}')
            
            # Переименовать колонки для удобства
            column_mapping = {
                'Дата': 'call_date',
                'Время приема вызова': 'call_time',
                'Продолжительность звонка': 'duration',
                'Время передачи на бригаду': 'transfer_time',
                'Время завершения вызова': 'close_time',
                'Служба': 'service_name',
                'Повод': 'reason',
                'Номер карточки': 'card_number',
                'Номер инцидента': 'incident_number',
                'Оператор': 'operator_name',
                'Статус': 'status',
                'ФИО вызывающего': 'caller_name',
                'Регион': 'region',
                'Район': 'district',
                'Адрес': 'address',
                'Место вызова': 'location_type',
                'Описание': 'description',
                'Само отказ': 'self_refusal',
                'Номер телефона заявителя': 'caller_phone'
            }
            
            df = df.rename(columns=column_mapping)
            
            # Извлечь код службы (102, 103, 104)
            df['service_code'] = df['service_name'].astype(str).str.extract(r'(\d{3})')
            
            # Добавить имя файла
            df['source_file'] = file_path.name
            
            # Выбрать нужные колонки
            columns_to_import = [
                'call_date', 'call_time', 'duration', 'transfer_time', 'close_time',
                'service_code', 'service_name', 'reason', 'card_number', 'incident_number',
                'operator_name', 'status', 'caller_name', 'region', 'district',
                'address', 'location_type', 'description', 'self_refusal', 
                'caller_phone', 'source_file'
            ]
            
            df_to_import = df[columns_to_import].copy()
            
            # Импортировать в базу
            df_to_import.to_sql('call_history_112', conn, if_exists='append', index=False)
            
            total_imported += len(df_to_import)
            print(f'   ✅ Импортировано: {len(df_to_import)} записей')
            
        except Exception as e:
            print(f'   ❌ Ошибка: {e}')
    
    conn.close()
    
    print(f'\n✅ ВСЕГО ИМПОРТИРОВАНО: {total_imported} записей')

def show_statistics():
    """Показать статистику по импортированным данным"""
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('СТАТИСТИКА ИСТОРИИ ЗВОНКОВ 112')
    print('=' * 80)
    
    # Всего записей
    total = pd.read_sql_query('SELECT COUNT(*) as cnt FROM call_history_112', conn)['cnt'][0]
    print(f'\n📊 Всего записей: {total:,}')
    
    # По службам
    print('\n🚑 ПО СЛУЖБАМ:')
    services = pd.read_sql_query('''
        SELECT service_code, COUNT(*) as count
        FROM call_history_112
        WHERE service_code IS NOT NULL
        GROUP BY service_code
        ORDER BY service_code
    ''', conn)
    print(services.to_string(index=False))
    
    # По регионам
    print('\n🌍 ПО РЕГИОНАМ:')
    regions = pd.read_sql_query('''
        SELECT region, COUNT(*) as count
        FROM call_history_112
        WHERE region IS NOT NULL
        GROUP BY region
        ORDER BY count DESC
        LIMIT 10
    ''', conn)
    print(regions.to_string(index=False))
    
    # Связь с заявками
    print('\n🔗 СВЯЗЬ С ЗАЯВКАМИ:')
    matched = pd.read_sql_query('''
        SELECT COUNT(DISTINCT ch.incident_number) as matched_incidents
        FROM call_history_112 ch
        INNER JOIN applications a ON a.application_number = ch.incident_number
    ''', conn)
    print(f'Найдено совпадений по номеру инцидента: {matched["matched_incidents"][0]:,}')
    
    conn.close()

if __name__ == '__main__':
    print('=' * 80)
    print('ИМПОРТ ИСТОРИИ ЗВОНКОВ 112')
    print('=' * 80)
    
    # Создать таблицу
    create_call_history_table()
    
    # Импортировать данные
    import_call_history_files()
    
    # Показать статистику
    show_statistics()
