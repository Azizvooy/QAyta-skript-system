#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ИМПОРТ ВСЕХ ДАННЫХ В БД
=============================================================================
Импортирует данные из exported_sheets и incoming_data в БД
=============================================================================
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import sys
import re

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
EXPORTED_SHEETS_DIR = BASE_DIR / 'exported_sheets'
INCOMING_DATA_DIR = BASE_DIR / 'incoming_data'
LOG_DIR = BASE_DIR / 'logs' / 'database'

# Паттерны для исключения
EXCLUDE_PATTERNS = [
    'тренды', 'сводка', 'итого', 'total', 'summary', 'текущий месяц', 'предыдущий месяц'
]

def log_import(message, level='INFO'):
    """Логирование импорта"""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f'import_all_{datetime.now().strftime("%Y%m%d")}.log'
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    try:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f'[{timestamp}] [{level}] {message}\n')
        
        # Безопасный вывод только для важных сообщений
        if level in ['ERROR', 'SUCCESS']:
            try:
                print(f'[{level}] {message}', flush=True)
            except:
                pass  # Игнорируем ошибки кодировки
    except Exception as e:
        pass  # Игнорируем ошибки логирования

def should_exclude(text):
    """Проверка, нужно ли исключить"""
    if not text or pd.isna(text):
        return True
    
    text_str = str(text).strip().lower()
    
    if text_str in ['', '-', 'nan', 'none', 'null']:
        return True
    
    for pattern in EXCLUDE_PATTERNS:
        if pattern in text_str:
            return True
    
    return False

def get_or_create_operator(cursor, operator_name):
    """Получить или создать оператора"""
    if should_exclude(operator_name):
        return None
    
    operator_name = str(operator_name).strip()
    
    cursor.execute('SELECT operator_id FROM operators WHERE operator_name = ?', (operator_name,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    cursor.execute('INSERT INTO operators (operator_name) VALUES (?)', (operator_name,))
    return cursor.lastrowid

def get_or_create_service(cursor, service_code):
    """Получить или создать службу"""
    if should_exclude(service_code):
        return None
    
    service_code = str(service_code).strip()
    
    cursor.execute('SELECT service_id FROM services WHERE service_code = ?', (service_code,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    cursor.execute(
        'INSERT INTO services (service_code, service_name) VALUES (?, ?)', 
        (service_code, f'Служба {service_code}')
    )
    return cursor.lastrowid

def get_or_create_region(cursor, region_name):
    """Получить или создать регион"""
    if should_exclude(region_name):
        return None
    
    region_name = str(region_name).strip()
    
    cursor.execute('SELECT region_id FROM regions WHERE region_name = ?', (region_name,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    cursor.execute('INSERT INTO regions (region_name) VALUES (?)', (region_name,))
    return cursor.lastrowid

def normalize_date(date_val):
    """Нормализация даты"""
    if pd.isna(date_val):
        return None
    
    # Если это уже дата
    if isinstance(date_val, pd.Timestamp):
        return date_val.strftime('%Y-%m-%d')
    
    # Попытка преобразовать
    try:
        date_str = str(date_val).strip()
        # Попытка разных форматов
        for fmt in ['%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y']:
            try:
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime('%Y-%m-%d')
            except:
                continue
        return date_str  # Возвращаем как есть
    except:
        return None

def import_csv_file(cursor, file_path, operator_name):
    """Импорт одного CSV файла"""
    try:
        df = pd.read_csv(file_path, encoding='utf-8')
        
        # Пропускаем файлы без данных
        if df.empty:
            return 0, 0
        
        imported = 0
        skipped = 0
        
        operator_id = get_or_create_operator(cursor, operator_name)
        if not operator_id:
            log_import(f'  Пропуск файла (исключенный оператор): {file_path.name}', 'WARNING')
            return 0, len(df)
        
        for idx, row in df.iterrows():
            try:
                # Пропускаем строки без ключевых данных (номер карты - основной идентификатор)
                card_number = row.get('Номер карты')
                if pd.isna(card_number) or str(card_number).strip() == '':
                    skipped += 1
                    continue
                
                card_num_str = str(card_number).strip()
                
                # Получаем или создаем связи
                service_id = None
                service_field = row.get('Выбор службы')
                if not pd.isna(service_field) and str(service_field).strip():
                    service_id = get_or_create_service(cursor, str(service_field).strip())
                
                # Формируем номер заявки из номера карты
                app_number = card_num_str
                
                # Проверяем, существует ли заявка
                cursor.execute('SELECT application_id FROM applications WHERE application_number = ?', (app_number,))
                app_result = cursor.fetchone()
                
                if not app_result:
                    # Создаем новую заявку только если её еще нет
                    cursor.execute('''
                        INSERT INTO applications (
                            application_number, card_number,
                            service_id,
                            call_date, caller_phone, status
                        ) VALUES (?, ?, ?, ?, ?, ?)
                    ''', (
                        app_number,
                        card_num_str,
                        service_id,
                        normalize_date(row.get('Дата открытия карты')),
                        str(row.get('Тел.номер', '')).strip() if not pd.isna(row.get('Тел.номер')) else None,
                        str(row.get('Статус связи', 'Новая')).strip() if not pd.isna(row.get('Статус связи')) else 'Новая'
                    ))
                    
                    cursor.execute('SELECT application_id FROM applications WHERE application_number = ?', (app_number,))
                    app_result = cursor.fetchone()
                
                if app_result:
                    application_id = app_result[0]
                    
                    # Добавляем фиксацию (каждая строка = новая фиксация)
                    fixation_date = normalize_date(row.get('Дата фиксации'))
                    
                    cursor.execute('''
                        INSERT INTO fixations (
                            application_id, operator_id,
                            fixation_date,
                            phone_called, status, notes
                        ) VALUES (?, ?, ?, ?, ?, ?)
                    ''', (
                        application_id,
                        operator_id,
                        fixation_date,
                        str(row.get('Тел.номер', '')).strip() if not pd.isna(row.get('Тел.номер')) else None,
                        str(row.get('Статус связи', 'Новая')).strip() if not pd.isna(row.get('Статус связи')) else 'Новая',
                        str(row.get('Комментарии', '')).strip() if not pd.isna(row.get('Комментарии')) else None
                    ))
                    
                    imported += 1
                else:
                    skipped += 1
                    
            except Exception as e:
                log_import(f'    Ошибка в строке {idx}: {str(e)}', 'ERROR')
                skipped += 1
                continue
        
        return imported, skipped
        
    except Exception as e:
        log_import(f'  Ошибка при чтении файла: {str(e)}', 'ERROR')
        return 0, 0

def import_all_csv_from_exported_sheets():
    """Импорт всех CSV из exported_sheets"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    log_import('='*80)
    log_import('ИМПОРТ ДАННЫХ ИЗ EXPORTED_SHEETS')
    log_import('='*80)
    
    total_imported = 0
    total_skipped = 0
    total_files = 0
    
    # Проходим по всем папкам операторов
    operator_dirs = [d for d in EXPORTED_SHEETS_DIR.iterdir() if d.is_dir()]
    total_operators = len(operator_dirs)
    
    print(f'Найдено операторов: {total_operators}')
    print()
    
    for idx, operator_dir in enumerate(operator_dirs, 1):
        operator_name = operator_dir.name
        
        # Пропускаем служебные папки
        if operator_name.startswith('-') or operator_name.startswith('.'):
            continue
        
        print(f'[{idx}/{total_operators}] {operator_name}', flush=True)
        log_import(f'\nОператор: {operator_name}')
        
        # Импортируем все CSV файлы из папки
        csv_files = list(operator_dir.glob('*.csv'))
        
        if not csv_files:
            log_import(f'  Нет CSV файлов')
            continue
        
        for csv_file in csv_files:
            log_import(f'  Файл: {csv_file.name}')
            imported, skipped = import_csv_file(cursor, csv_file, operator_name)
            total_imported += imported
            total_skipped += skipped
            total_files += 1
            
            if imported > 0:
                log_import(f'    OK: Импортировано: {imported}, Пропущено: {skipped}')
            
            # Коммитим после каждого файла
            conn.commit()
    
    conn.close()
    
    log_import('')
    log_import('='*80)
    log_import(f'ИТОГО:')
    log_import(f'  Файлов обработано: {total_files}')
    log_import(f'  Записей импортировано: {total_imported}')
    log_import(f'  Записей пропущено: {total_skipped}')
    log_import('='*80)
    
    return total_imported, total_skipped, total_files

def main():
    """Главная функция"""
    print('\n' + '='*80)
    print('ИМПОРТ ВСЕХ ДАННЫХ В БАЗУ ДАННЫХ')
    print('='*80)
    print(f'Дата: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    
    # Проверяем БД
    if not DB_PATH.exists():
        print('ОШИБКА: База данных не найдена!')
        print('   Сначала создайте БД: python scripts/database/db_schema.py')
        return
    
    print(f'OK: База данных: {DB_PATH}')
    print('Начинаю импорт...')
    print()
    
    # Запускаем импорт
    try:
        imported, skipped, files = import_all_csv_from_exported_sheets()
        
        print()
        print('='*80)
        print('Успешно завершен импорт!')
        print(f'   Файлов: {files}')
        print(f'   Импортировано: {imported:,}')
        print(f'   Пропущено: {skipped:,}')
        print('='*80)
        
    except Exception as e:
        print()
        print('='*80)
        print(f'ОШИБКА при импорте: {e}')
        print('='*80)
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
