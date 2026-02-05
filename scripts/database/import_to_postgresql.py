#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ИМПОРТ ДАННЫХ В POSTGRESQL
=============================================================================
Импорт данных из CSV файлов и SQLite в PostgreSQL
=============================================================================
"""

import psycopg2
from psycopg2.extras import execute_batch
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
import os
from dotenv import load_dotenv
from tqdm import tqdm

BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / 'config'
SQLITE_DB = BASE_DIR / 'data' / 'fiksa_database.db'
EXPORT_DIR = BASE_DIR / 'exported_sheets'

# Загрузка конфигурации
load_dotenv(CONFIG_DIR / 'postgresql.env')

DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', 5432)),
    'database': os.getenv('DB_NAME', 'qayta_data'),
    'user': os.getenv('DB_USER', 'qayta_user'),
    'password': os.getenv('DB_PASSWORD', 'qayta_password_2026')
}

print('\n' + '='*80)
print('📥 ИМПОРТ ДАННЫХ В POSTGRESQL')
print('='*80)

def get_or_create_operator(cursor, operator_name):
    """Получить или создать оператора"""
    if not operator_name or pd.isna(operator_name):
        return None
    
    operator_name = str(operator_name).strip()
    
    cursor.execute(
        "SELECT operator_id FROM operators WHERE operator_name = %s",
        (operator_name,)
    )
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    cursor.execute(
        "INSERT INTO operators (operator_name) VALUES (%s) RETURNING operator_id",
        (operator_name,)
    )
    return cursor.fetchone()[0]

def get_or_create_service(cursor, service_code):
    """Получить или создать службу"""
    if not service_code or pd.isna(service_code):
        return None
    
    service_code = str(service_code).strip()
    
    cursor.execute(
        "SELECT service_id FROM services WHERE service_code = %s",
        (service_code,)
    )
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    cursor.execute(
        """INSERT INTO services (service_code, service_name) 
           VALUES (%s, %s) RETURNING service_id""",
        (service_code, f'Служба {service_code}')
    )
    return cursor.fetchone()[0]

def get_or_create_region(cursor, region_name):
    """Получить или создать регион"""
    if not region_name or pd.isna(region_name):
        return None
    
    region_name = str(region_name).strip()
    
    cursor.execute(
        "SELECT region_id FROM regions WHERE region_name = %s",
        (region_name,)
    )
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    cursor.execute(
        "INSERT INTO regions (region_name) VALUES (%s) RETURNING region_id",
        (region_name,)
    )
    return cursor.fetchone()[0]

def import_from_sqlite():
    """Импорт данных из SQLite"""
    print('\n[1/2] Импорт из SQLite БД...')
    
    if not SQLITE_DB.exists():
        print('  ⚠️ SQLite база не найдена, пропускаем')
        return 0
    
    try:
        # Подключение к SQLite
        sqlite_conn = sqlite3.connect(SQLITE_DB)
        df = pd.read_sql_query("SELECT * FROM fiksa_records", sqlite_conn)
        sqlite_conn.close()
        
        print(f'  📊 Загружено из SQLite: {len(df):,} записей')
        
        # Подключение к PostgreSQL
        pg_conn = psycopg2.connect(**DB_CONFIG)
        cursor = pg_conn.cursor()
        
        imported = 0
        batch_size = 1000
        
        print('  🔄 Импорт данных...')
        for i in tqdm(range(0, len(df), batch_size), desc='  Прогресс'):
            batch = df.iloc[i:i+batch_size]
            
            for _, row in batch.iterrows():
                try:
                    # Получаем ID оператора
                    operator_id = get_or_create_operator(cursor, row.get('operator_name'))
                    
                    # Вставляем фиксацию
                    cursor.execute("""
                        INSERT INTO fixations 
                        (card_number, operator_id, call_date, status, phone, 
                         source_file, import_date, collection_date)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        ON CONFLICT (card_number, call_date) DO NOTHING
                    """, (
                        row.get('card_number'),
                        operator_id,
                        row.get('call_date'),
                        row.get('status'),
                        row.get('phone'),
                        row.get('source_file'),
                        row.get('import_date'),
                        row.get('collection_date')
                    ))
                    
                    imported += 1
                    
                except Exception as e:
                    continue
            
            pg_conn.commit()
        
        cursor.close()
        pg_conn.close()
        
        print(f'  ✅ Импортировано из SQLite: {imported:,} записей')
        return imported
        
    except Exception as e:
        print(f'  ❌ Ошибка импорта из SQLite: {e}')
        return 0

def import_from_csv():
    """Импорт данных из CSV файлов"""
    print('\n[2/2] Импорт из CSV файлов...')
    
    csv_files = list(EXPORT_DIR.rglob('*.csv'))
    
    if not csv_files:
        print('  ⚠️ CSV файлы не найдены')
        return 0
    
    print(f'  📁 Найдено файлов: {len(csv_files)}')
    
    try:
        pg_conn = psycopg2.connect(**DB_CONFIG)
        cursor = pg_conn.cursor()
        
        imported = 0
        errors = 0
        
        for csv_file in tqdm(csv_files, desc='  Обработка файлов'):
            try:
                df = pd.read_csv(csv_file, encoding='utf-8-sig', low_memory=False)
                
                if len(df) == 0:
                    continue
                
                # Определяем оператора из имени папки
                operator_name = csv_file.parent.name
                if operator_name in ['exported_sheets', '-']:
                    operator_name = csv_file.stem
                
                # Получаем ID оператора
                operator_id = get_or_create_operator(cursor, operator_name)
                
                # Определяем колонки
                card_col = next((col for col in ['Номер карты', 'Код карты', 'card_number'] if col in df.columns), None)
                status_col = next((col for col in ['Статус связи', 'Причина/Статус', 'Статус', 'status'] if col in df.columns), None)
                date_col = next((col for col in ['Дата фиксации', 'Дата открытия карты', 'call_date'] if col in df.columns), None)
                
                if not card_col:
                    continue
                
                for _, row in df.iterrows():
                    try:
                        cursor.execute("""
                            INSERT INTO fixations 
                            (card_number, operator_id, call_date, status, source_file, import_date)
                            VALUES (%s, %s, %s, %s, %s, %s)
                            ON CONFLICT (card_number, call_date) DO NOTHING
                        """, (
                            row.get(card_col),
                            operator_id,
                            row.get(date_col) if date_col else None,
                            row.get(status_col) if status_col else None,
                            csv_file.name,
                            datetime.now()
                        ))
                        
                        imported += 1
                        
                    except:
                        errors += 1
                        continue
                
                pg_conn.commit()
                
            except Exception as e:
                errors += 1
                continue
        
        cursor.close()
        pg_conn.close()
        
        print(f'  ✅ Импортировано из CSV: {imported:,} записей')
        if errors > 0:
            print(f'  ⚠️ Ошибок: {errors}')
        
        return imported
        
    except Exception as e:
        print(f'  ❌ Ошибка импорта из CSV: {e}')
        return 0

def show_statistics():
    """Показать статистику БД"""
    print('\n' + '='*80)
    print('📊 СТАТИСТИКА POSTGRESQL')
    print('='*80)
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        # Общая статистика
        cursor.execute("SELECT COUNT(*) FROM fixations")
        total = cursor.fetchone()[0]
        print(f'\n📌 Всего фиксаций: {total:,}')
        
        # Статистика по операторам
        cursor.execute("SELECT * FROM v_operator_statistics ORDER BY total_fixations DESC LIMIT 10")
        operators = cursor.fetchall()
        
        if operators:
            print(f'\n👥 ТОП-10 операторов:')
            for idx, (op_id, name, total, pos, neg, no_ans, pct) in enumerate(operators, 1):
                print(f'   {idx:2}. {name:<40} {total:>7,} ({pct or 0:>5.1f}% положит.)')
        
        # Статистика по службам
        cursor.execute("SELECT * FROM v_service_statistics ORDER BY total_fixations DESC")
        services = cursor.fetchall()
        
        if services:
            print(f'\n🚑 Статистика по службам:')
            for service_id, code, name, total, pos, neg, regions in services:
                if total > 0:
                    print(f'   {code}: {total:,} фиксаций ({pos:,} полож., {neg:,} отриц.)')
        
        # Статистика по категориям
        cursor.execute("""
            SELECT status_category, COUNT(*) 
            FROM fixations 
            WHERE status_category IS NOT NULL
            GROUP BY status_category 
            ORDER BY COUNT(*) DESC
        """)
        categories = cursor.fetchall()
        
        if categories:
            print(f'\n📋 По категориям:')
            for category, count in categories:
                pct = (count / total * 100) if total > 0 else 0
                print(f'   {category:<20} {count:>10,} ({pct:>5.1f}%)')
        
        cursor.close()
        conn.close()
        
        print('\n' + '='*80)
        
    except Exception as e:
        print(f'❌ Ошибка получения статистики: {e}')

def main():
    """Главная функция"""
    
    # Импорт из SQLite
    sqlite_count = import_from_sqlite()
    
    # Импорт из CSV
    csv_count = import_from_csv()
    
    # Статистика
    show_statistics()
    
    print('\n' + '='*80)
    print('✅ ИМПОРТ ЗАВЕРШЕН УСПЕШНО!')
    print('='*80)
    print(f'   Из SQLite: {sqlite_count:,}')
    print(f'   Из CSV:    {csv_count:,}')
    print(f'   ВСЕГО:     {sqlite_count + csv_count:,}')
    print('='*80 + '\n')

if __name__ == '__main__':
    main()
