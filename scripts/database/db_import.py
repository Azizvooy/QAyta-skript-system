#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ИМПОРТ ДАННЫХ В УЛУЧШЕННУЮ БД
=============================================================================
Импортирует данные из Excel/CSV в правильную структуру БД
Исключает пустые значения и неправильные связи
=============================================================================
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import sys

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
DATA_DIR = BASE_DIR / 'data'
IMPORT_DIR = BASE_DIR / '123'
LOG_DIR = BASE_DIR / 'logs' / 'database'

# Паттерны для исключения из импорта (имена операторов, которые нужно пропускать)
EXCLUDE_PATTERNS = [
    'Тренды',
    'Текущий месяц - Сводка',
    'Предыдущий месяц - Сводка',
    'СВОДКА СОТРУДНИКИ',
    'Ноябрь 2025',
    'Декабрь 2025',
    'сводка',  # любые сводки
    'итого',   # любые итоги
    'total',
    'summary',
]

def should_exclude_operator(operator_name):
    """Проверка, нужно ли исключить оператора"""
    if not operator_name:
        return True
    
    operator_str = str(operator_name).strip().lower()
    
    # Пустые значения
    if operator_str in ['', '-', 'nan', 'none', 'null']:
        return True
    
    # Проверка паттернов
    for pattern in EXCLUDE_PATTERNS:
        if pattern.lower() in operator_str:
            return True
    
    return False

def log_import(message, level='INFO'):
    """Логирование импорта"""
    log_file = LOG_DIR / f'import_log_{datetime.now().strftime("%Y%m%d")}.log'
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'[{timestamp}] [{level}] {message}\n')
    
    print(f'[{level}] {message}')

def get_or_create_operator(cursor, operator_name):
    """Получить или создать оператора"""
    # Проверяем, нужно ли исключить
    if should_exclude_operator(operator_name):
        return None
    
    operator_name = str(operator_name).strip()
    
    # Проверяем существование
    cursor.execute('SELECT operator_id FROM operators WHERE operator_name = ?', (operator_name,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    # Создаем нового
    cursor.execute('''
        INSERT INTO operators (operator_name, position)
        VALUES (?, ?)
    ''', (operator_name, 'Оператор'))
    
    return cursor.lastrowid

def get_or_create_service(cursor, service_code, service_name=None):
    """Получить или создать службу"""
    if not service_code or str(service_code).strip() == '':
        return None
    
    service_code = str(service_code).strip()
    
    # Проверяем существование
    cursor.execute('SELECT service_id FROM services WHERE service_code = ?', (service_code,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    # Создаем новую
    if not service_name:
        service_name = f'Служба {service_code}'
    
    cursor.execute('''
        INSERT INTO services (service_code, service_name)
        VALUES (?, ?)
    ''', (service_code, service_name))
    
    return cursor.lastrowid

def get_or_create_region(cursor, region_name):
    """Получить или создать регион"""
    if not region_name or str(region_name).strip() == '' or str(region_name).lower() == 'не указано':
        return None
    
    region_name = str(region_name).strip()
    
    # Проверяем существование
    cursor.execute('SELECT region_id FROM regions WHERE region_name = ?', (region_name,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    # Создаем новый
    cursor.execute('''
        INSERT INTO regions (region_name)
        VALUES (?)
    ''', (region_name,))
    
    return cursor.lastrowid

def import_applications_from_excel():
    """Импорт заявок из Excel файлов"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    log_import('Начало импорта заявок из Excel')
    
    # Ищем файлы с историей звонков
    excel_files = list(IMPORT_DIR.glob('История*.xlsx'))
    
    if not excel_files:
        log_import('Нет файлов для импорта в папке 123/', 'WARNING')
        conn.close()
        return
    
    total_imported = 0
    total_skipped = 0
    
    for file_path in excel_files:
        log_import(f'Импорт файла: {file_path.name}')
        
        try:
            df = pd.read_excel(file_path)
            log_import(f'  Найдено записей: {len(df)}')
            
            for idx, row in df.iterrows():
                try:
                    # Пропускаем записи без ключевых данных
                    if pd.isna(row.get('Дата звонка')) and pd.isna(row.get('Номер инцидента')):
                        total_skipped += 1
                        continue
                    
                    # Получаем или создаем связанные записи
                    service_id = None
                    if 'Код службы' in row and not pd.isna(row['Код службы']):
                        service_name = row.get('Название службы', None)
                        service_id = get_or_create_service(cursor, row['Код службы'], service_name)
                    
                    region_id = None
                    if 'Область' in row and not pd.isna(row['Область']):
                        region_id = get_or_create_region(cursor, row['Область'])
                    
                    # Формируем уникальный номер заявки
                    application_number = str(row.get('Номер инцидента', f'APP_{datetime.now().timestamp()}')).strip()
                    if not application_number or application_number == 'nan':
                        application_number = f'APP_{datetime.now().timestamp()}_{idx}'
                    
                    # Вставляем заявку
                    cursor.execute('''
                        INSERT OR REPLACE INTO applications (
                            application_number, card_number, incident_number,
                            service_id, region_id,
                            caller_name, caller_phone,
                            call_date, call_time, address, district, reason, description,
                            status, notes
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        application_number,
                        str(row.get('Номер карты', '')).strip() if not pd.isna(row.get('Номер карты')) else None,
                        str(row.get('Номер инцидента', '')).strip() if not pd.isna(row.get('Номер инцидента')) else None,
                        service_id,
                        region_id,
                        str(row.get('ФИО заявителя', '')).strip() if not pd.isna(row.get('ФИО заявителя')) else None,
                        str(row.get('Телефон', '')).strip() if not pd.isna(row.get('Телефон')) else None,
                        row.get('Дата звонка'),
                        str(row.get('Время звонка', '')).strip() if not pd.isna(row.get('Время звонка')) else None,
                        str(row.get('Адрес', '')).strip() if not pd.isna(row.get('Адрес')) else None,
                        str(row.get('Район', '')).strip() if not pd.isna(row.get('Район')) else None,
                        str(row.get('Повод', '')).strip() if not pd.isna(row.get('Повод')) else None,
                        str(row.get('Описание', '')).strip() if not pd.isna(row.get('Описание')) else None,
                        str(row.get('Статус', 'Новая')).strip() if not pd.isna(row.get('Статус')) else 'Новая',
                        str(row.get('Примечание', '')).strip() if not pd.isna(row.get('Примечание')) else None
                    ))
                    
                    total_imported += 1
                    
                except Exception as e:
                    log_import(f'  Ошибка в строке {idx}: {str(e)}', 'ERROR')
                    total_skipped += 1
                    continue
            
            conn.commit()
            log_import(f'  ✅ Импортировано: {total_imported}')
            
        except Exception as e:
            log_import(f'  Ошибка при обработке файла: {str(e)}', 'ERROR')
            conn.rollback()
    
    conn.close()
    
    log_import(f'Импорт завершен. Импортировано: {total_imported}, Пропущено: {total_skipped}', 'SUCCESS')
    return total_imported, total_skipped

def import_fixations_from_csv():
    """Импорт фиксаций из CSV файлов"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    log_import('Начало импорта фиксаций из CSV')
    
    # Ищем CSV файлы с данными
    csv_files = list(DATA_DIR.glob('ALL_DATA*.csv'))
    
    if not csv_files:
        log_import('Нет CSV файлов для импорта', 'WARNING')
        conn.close()
        return
    
    total_imported = 0
    total_skipped = 0
    
    for file_path in csv_files:
        log_import(f'Импорт файла: {file_path.name}')
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            log_import(f'  Найдено записей: {len(df)}')
            
            for idx, row in df.iterrows():
                try:
                    # Пропускаем записи без оператора
                    if pd.isna(row.get('Оператор фиксировавший')):
                        total_skipped += 1
                        continue
                    
                    # Получаем или создаем оператора
                    operator_id = get_or_create_operator(cursor, row['Оператор фиксировавший'])
                    if not operator_id:
                        total_skipped += 1
                        continue
                    
                    # Ищем заявку по номеру карты
                    application_id = None
                    if 'Номер карты' in row and not pd.isna(row['Номер карты']):
                        card_number = str(row['Номер карты']).strip()
                        cursor.execute('''
                            SELECT application_id FROM applications 
                            WHERE card_number = ? 
                            LIMIT 1
                        ''', (card_number,))
                        result = cursor.fetchone()
                        if result:
                            application_id = result[0]
                    
                    # Дата фиксации
                    collection_date = row.get('Дата фиксации', datetime.now().strftime('%Y-%m-%d'))
                    if pd.isna(collection_date):
                        collection_date = datetime.now().strftime('%Y-%m-%d')
                    
                    # Вставляем фиксацию
                    cursor.execute('''
                        INSERT INTO fixations (
                            application_id, operator_id, collection_date,
                            fixation_date, phone_called, status, notes
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        application_id,
                        operator_id,
                        collection_date,
                        row.get('Дата фиксации'),
                        str(row.get('Номер телефона', '')).strip() if not pd.isna(row.get('Номер телефона')) else None,
                        str(row.get('Статус', 'Неизвестно')).strip() if not pd.isna(row.get('Статус')) else 'Неизвестно',
                        str(row.get('Комментарий', '')).strip() if not pd.isna(row.get('Комментарий')) else None
                    ))
                    
                    total_imported += 1
                    
                except Exception as e:
                    log_import(f'  Ошибка в строке {idx}: {str(e)}', 'ERROR')
                    total_skipped += 1
                    continue
            
            conn.commit()
            log_import(f'  ✅ Импортировано: {total_imported}')
            
        except Exception as e:
            log_import(f'  Ошибка при обработке файла: {str(e)}', 'ERROR')
            conn.rollback()
    
    conn.close()
    
    log_import(f'Импорт фиксаций завершен. Импортировано: {total_imported}, Пропущено: {total_skipped}', 'SUCCESS')
    return total_imported, total_skipped

def main():
    """Основная функция импорта"""
    print('=' * 80)
    print('ИМПОРТ ДАННЫХ В УЛУЧШЕННУЮ БД')
    print('=' * 80)
    
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    
    # Импорт заявок
    print('\n1️⃣  Импорт заявок из Excel...')
    app_imported, app_skipped = import_applications_from_excel()
    print(f'   ✅ Импортировано заявок: {app_imported}')
    print(f'   ⚠️  Пропущено: {app_skipped}')
    
    # Импорт фиксаций
    print('\n2️⃣  Импорт фиксаций из CSV...')
    fix_imported, fix_skipped = import_fixations_from_csv()
    print(f'   ✅ Импортировано фиксаций: {fix_imported}')
    print(f'   ⚠️  Пропущено: {fix_skipped}')
    
    print('\n' + '=' * 80)
    print('ИМПОРТ ЗАВЕРШЕН!')
    print('=' * 80)

if __name__ == '__main__':
    main()
