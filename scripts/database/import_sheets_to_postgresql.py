#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ИМПОРТ ДАННЫХ ИЗ GOOGLE SHEETS В POSTGRESQL
=============================================================================
Импорт свежих данных напрямую из Google Sheets в PostgreSQL
=============================================================================
"""

import sys
import io

# Установка UTF-8 для вывода
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import psycopg2
from psycopg2.extras import execute_batch
import pandas as pd
from pathlib import Path
from datetime import datetime
import os
from dotenv import load_dotenv
from tqdm import tqdm

BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / 'config'
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
print('📥 ИМПОРТ ИЗ GOOGLE SHEETS В POSTGRESQL')
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

def import_from_sheets():
    """Импорт данных из Google Sheets CSV файлов"""
    print('\n[1/1] Импорт из Google Sheets...')
    
    # Находим все CSV файлы
    csv_files = list(EXPORT_DIR.rglob('*.csv'))
    
    if not csv_files:
        print('  ⚠️ CSV файлы не найдены в exported_sheets/')
        print('  Сначала запустите: update_from_sheets.py')
        return 0
    
    print(f'  📁 Найдено файлов: {len(csv_files)}')
    
    try:
        pg_conn = psycopg2.connect(**DB_CONFIG)
        cursor = pg_conn.cursor()
        
        # Очистить таблицу fixations перед импортом
        print('\n  🗑️ Очистка старых данных...')
        cursor.execute("TRUNCATE TABLE fixations RESTART IDENTITY CASCADE")
        pg_conn.commit()
        print('  ✅ Старые данные очищены')
        
        imported = 0
        errors = 0
        
        print('\n  🔄 Импорт данных из CSV файлов...')
        
        for csv_file in tqdm(csv_files, desc='  Обработка файлов'):
            try:
                # Читаем CSV
                df = pd.read_csv(csv_file, encoding='utf-8-sig', low_memory=False)
                
                if len(df) == 0:
                    continue
                
                # Определяем оператора из пути
                parts = csv_file.parts
                operator_name = None
                
                for i, part in enumerate(parts):
                    if part == 'exported_sheets' and i + 1 < len(parts):
                        operator_name = parts[i + 1]
                        break
                
                if not operator_name or operator_name == '-':
                    operator_name = csv_file.stem
                
                # Получаем ID оператора
                operator_id = get_or_create_operator(cursor, operator_name)
                
                # Определяем колонки
                card_col = next((col for col in ['Номер карты', 'Код карты', 'card_number'] if col in df.columns), None)
                status_col = next((col for col in ['Статус связи', 'Причина/Статус', 'Статус', 'status'] if col in df.columns), None)
                date_col = next((col for col in ['Дата фиксации', 'Дата открытия карты', 'call_date'] if col in df.columns), None)
                phone_col = next((col for col in ['Телефон', 'Номер телефона', 'phone'] if col in df.columns), None)
                
                if not card_col:
                    continue
                
                # Пакетная вставка
                batch = []
                for _, row in df.iterrows():
                    try:
                        card_number = row.get(card_col)
                        call_date_raw = row.get(date_col) if date_col else None
                        
                        if not card_number or pd.isna(card_number):
                            continue
                        
                        # Парс даты
                        call_date = None
                        if call_date_raw and not pd.isna(call_date_raw):
                            try:
                                call_date = pd.to_datetime(call_date_raw, dayfirst=True)
                            except:
                                call_date = None
                        
                        batch.append((
                            str(card_number),
                            operator_id,
                            call_date,
                            row.get(status_col) if status_col else None,
                            row.get(phone_col) if phone_col else None,
                            csv_file.name,
                            datetime.now()
                        ))
                        
                        # Вставляем пакетами по 5000
                        if len(batch) >= 5000:
                            try:
                                cursor.executemany("""
                                    INSERT INTO fixations 
                                    (card_number, operator_id, call_date, status, phone, source_file, import_date)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                                """, batch)
                                pg_conn.commit()
                                imported += len(batch)
                            except Exception as e:
                                print(f'\n  ❌ Ошибка вставки: {e}')
                                pg_conn.rollback()
                                errors += len(batch)
                            batch = []
                        
                    except Exception as e:
                        errors += 1
                        continue
                
                # Вставляем остаток
                if batch:
                    try:
                        cursor.executemany("""
                            INSERT INTO fixations 
                            (card_number, operator_id, call_date, status, phone, source_file, import_date)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """, batch)
                        pg_conn.commit()
                        imported += len(batch)
                    except Exception as e:
                        pg_conn.rollback()
                        errors += len(batch)
                
            except Exception as e:
                print(f'\n  ⚠️ Ошибка в файле {csv_file.name}: {e}')
                errors += 1
                continue
        
        cursor.close()
        pg_conn.close()
        
        print(f'\n  ✅ Импортировано записей: {imported:,}')
        if errors > 0:
            print(f'  ⚠️ Ошибок: {errors}')
        
        return imported
        
    except Exception as e:
        print(f'\n  ❌ Ошибка импорта: {e}')
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
        cursor.execute("""
            SELECT operator_name, total_fixations, positive_count, positive_percentage
            FROM v_operator_statistics 
            WHERE total_fixations > 0
            ORDER BY total_fixations DESC 
            LIMIT 10
        """)
        operators = cursor.fetchall()
        
        if operators:
            print(f'\n👥 ТОП-10 операторов:')
            for idx, (name, total, pos, pct) in enumerate(operators, 1):
                print(f'   {idx:2}. {name:<40} {total:>7,} ({pct or 0:>5.1f}% положит.)')
        
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
        
        # Дата импорта
        cursor.execute("SELECT MAX(import_date) FROM fixations")
        last_import = cursor.fetchone()[0]
        if last_import:
            print(f'\n📅 Последний импорт: {last_import}')
        
        cursor.close()
        conn.close()
        
        print('\n' + '='*80)
        
    except Exception as e:
        print(f'❌ Ошибка получения статистики: {e}')

def main():
    """Главная функция"""
    
    # Импорт из Google Sheets
    imported = import_from_sheets()
    
    if imported > 0:
        # Статистика
        show_statistics()
        
        print('\n' + '='*80)
        print('✅ ИМПОРТ ЗАВЕРШЕН УСПЕШНО!')
        print('='*80)
        print(f'   Импортировано: {imported:,} записей')
        print('\n📊 Доступ к данным:')
        print('   • pgAdmin:  http://localhost:5050')
        print('   • PostgreSQL: localhost:5432')
        print('='*80 + '\n')
    else:
        print('\n⚠️ Нет данных для импорта')
        print('Запустите сначала: update_from_sheets.py\n')

if __name__ == '__main__':
    main()
