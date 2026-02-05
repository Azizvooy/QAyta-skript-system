#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
АКТУАЛЬНЫЙ СБОРЩИК ДАННЫХ ИЗ GOOGLE SHEETS В БД
=============================================================================
Версия: 2.0 (для новой структуры БД)
Дата: 08.01.2026

📋 НАЗНАЧЕНИЕ:
Собирает данные из Google Sheets и сохраняет в НОВУЮ структуру БД
- Таблица fixations (фиксации с нормализованными связями)
- Таблица operators (операторы)
- Таблица services (службы 102/103/104)

ЗАМЕНЯЕТ УСТАРЕВШИЕ:
❌ improved_collector.py → archive/
❌ daily_db_collector.py → archive/

=============================================================================
"""

import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import socket

# Прокси для работы
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
TOKEN_FILE = BASE_DIR / 'config' / 'token.json'
CREDENTIALS_FILE = BASE_DIR / 'config' / 'credentials.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

# Паттерны для исключения служебных листов
EXCLUDE_PATTERNS = [
    'Настройки', 'Статистика', 'Сводка', 'Тренды', 
    'Текущий месяц', 'Предыдущий месяц', 'СВОДКА', 'Итого'
]

# =============================================================================
# GOOGLE API
# =============================================================================

def authenticate():
    """Аутентификация в Google API"""
    creds = None
    
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                print("❌ ОШИБКА: файл config/credentials.json не найден!")
                print("\n📝 Инструкция по настройке Google API:")
                print("   См. документ: НАСТРОЙКА_GOOGLE_API.md")
                sys.exit(1)
            
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Создаем папку config если не существует
        TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return creds

def get_operator_sheets(service):
    """Получить список всех листов операторов"""
    spreadsheet = service.spreadsheets().get(spreadsheetId=MASTER_SPREADSHEET_ID).execute()
    sheets = spreadsheet.get('sheets', [])
    
    operator_sheets = []
    for sheet in sheets:
        title = sheet['properties']['title']
        
        # Пропускаем служебные листы
        is_excluded = any(pattern.lower() in title.lower() for pattern in EXCLUDE_PATTERNS)
        
        if not is_excluded:
            operator_sheets.append(title)
    
    return operator_sheets

def collect_fiksa_data(operator_name, service):
    """Собрать данные с листа оператора
    
    Структура Google Sheets (колонки):
    A - Номер карты
    B - ФИО
    C - Телефон
    D - Адрес
    E - Статус связи (ОБЯЗАТЕЛЬНО должен быть заполнен!)
    F - Дата звонка
    G - Примечания
    """
    try:
        # Читаем данные со всего листа
        range_name = f"'{operator_name}'!A2:Z10000"
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return []
        
        records = []
        today = datetime.now().strftime('%Y-%m-%d')
        
        for row in values:
            # Проверяем что строка не пустая
            if not row or len(row) < 5:
                continue
            
            # ⭐ КРИТИЧНО: Проверяем колонку E (индекс 4) - статус должен быть заполнен
            # Это основное условие фильтрации - без статуса запись неполная
            status = row[4] if len(row) > 4 else ''
            if not status or status.strip() == '':
                continue  # Пропускаем строки без статуса
            
            # Проверяем номер карты (колонка A)
            card_number = row[0] if len(row) > 0 else ''
            if not card_number or card_number.strip() == '':
                continue  # Пропускаем строки без номера карты
            
            # Формируем запись
            record = {
                'collection_date': today,
                'operator_name': operator_name,
                'card_number': row[0].strip() if len(row) > 0 else None,
                'full_name': row[1].strip() if len(row) > 1 else None,
                'phone': row[2].strip() if len(row) > 2 else None,
                'address': row[3].strip() if len(row) > 3 else None,
                'status': status.strip(),
                'call_date': row[5].strip() if len(row) > 5 else None,
                'notes': row[6].strip() if len(row) > 6 else None,
            }
            
            records.append(record)
        
        return records
        
    except Exception as e:
        print(f'   ⚠️  [{operator_name}] Ошибка: {e}')
        return []

# =============================================================================
# БАЗА ДАННЫХ (НОВАЯ СТРУКТУРА)
# =============================================================================

def should_exclude_operator(operator_name):
    """Проверка, нужно ли исключить оператора из обработки"""
    if not operator_name:
        return True
    
    operator_str = str(operator_name).strip().lower()
    
    # Пустые значения
    if operator_str in ['', '-', 'nan', 'none', 'null']:
        return True
    
    # Служебные листы
    for pattern in EXCLUDE_PATTERNS:
        if pattern.lower() in operator_str:
            return True
    
    return False

def get_or_create_operator(cursor, operator_name):
    """Получить или создать оператора в БД"""
    if should_exclude_operator(operator_name):
        return None
    
    operator_name = str(operator_name).strip()
    
    # Проверяем существование
    cursor.execute('SELECT operator_id FROM operators WHERE operator_name = ?', (operator_name,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    # Создаем нового оператора
    cursor.execute('''
        INSERT INTO operators (operator_name, position, is_active)
        VALUES (?, ?, ?)
    ''', (operator_name, 'Оператор 112', 1))
    
    return cursor.lastrowid

def extract_service_code(status_text):
    """Извлечь код службы из статуса
    
    Примеры статусов:
    - "Служба 102"
    - "102 - отказ"
    - "Передано в 103"
    """
    if not status_text:
        return None
    
    status_str = str(status_text).lower()
    
    if '102' in status_str:
        return '102'
    elif '103' in status_str:
        return '103'
    elif '104' in status_str:
        return '104'
    
    return None

def get_or_create_service(cursor, service_code):
    """Получить или создать службу в БД"""
    if not service_code:
        return None
    
    service_code = str(service_code).strip()
    
    # Проверяем существование
    cursor.execute('SELECT service_id FROM services WHERE service_code = ?', (service_code,))
    result = cursor.fetchone()
    
    if result:
        return result[0]
    
    # Определяем название службы
    service_names = {
        '102': 'Милиция',
        '103': 'Скорая помощь',
        '104': 'Пожарная служба'
    }
    
    service_name = service_names.get(service_code, f'Служба {service_code}')
    
    # Создаем новую службу
    cursor.execute('''
        INSERT INTO services (service_code, service_name)
        VALUES (?, ?)
    ''', (service_code, service_name))
    
    return cursor.lastrowid

def save_to_database(records):
    """Сохранить записи в БД (новая структура fixations)"""
    if not records:
        return 0, 0
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Проверяем наличие таблиц новой структуры
    cursor.execute("""
        SELECT name FROM sqlite_master 
        WHERE type='table' AND name='fixations'
    """)
    
    if not cursor.fetchone():
        print("\n❌ ОШИБКА: Таблица 'fixations' не найдена!")
        print("   Необходимо создать новую структуру БД.")
        print("   Запустите: python scripts/database/db_schema.py")
        conn.close()
        return 0, 0
    
    inserted = 0
    updated = 0
    today = datetime.now().strftime('%Y-%m-%d')
    
    for record in records:
        # Получаем или создаем оператора
        operator_id = get_or_create_operator(cursor, record['operator_name'])
        if not operator_id:
            continue  # Пропускаем служебные записи
        
        # Извлекаем код службы из статуса
        service_code = extract_service_code(record['status'])
        service_id = get_or_create_service(cursor, service_code) if service_code else None
        
        # Проверяем существует ли запись (по номеру карты и оператору)
        cursor.execute('''
            SELECT fixation_id FROM fixations 
            WHERE card_number = ? AND operator_id = ?
        ''', (record['card_number'], operator_id))
        
        existing = cursor.fetchone()
        
        if existing:
            # Обновляем существующую запись
            cursor.execute('''
                UPDATE fixations SET
                    full_name = ?,
                    phone_called = ?,
                    address_declared = ?,
                    fixation_status = ?,
                    fixation_date = ?,
                    notes = ?,
                    service_id = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE fixation_id = ?
            ''', (
                record['full_name'],
                record['phone'],
                record['address'],
                record['status'],
                record['call_date'],
                record['notes'],
                service_id,
                existing[0]
            ))
            updated += 1
        else:
            # Вставляем новую запись
            cursor.execute('''
                INSERT INTO fixations (
                    operator_id, card_number, full_name, phone_called,
                    address_declared, fixation_status, fixation_date,
                    notes, service_id, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ''', (
                operator_id,
                record['card_number'],
                record['full_name'],
                record['phone'],
                record['address'],
                record['status'],
                record['call_date'],
                record['notes'],
                service_id
            ))
            inserted += 1
    
    conn.commit()
    conn.close()
    
    return inserted, updated

# =============================================================================
# ОСНОВНОЙ ПРОЦЕСС
# =============================================================================

def main():
    """Основная функция сбора данных"""
    print('\n' + '='*80)
    print('🚀 АКТУАЛЬНЫЙ СБОРЩИК ДАННЫХ ИЗ GOOGLE SHEETS В БД')
    print('='*80)
    print(f'⏰ Время запуска: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    
    try:
        # 1. Подключение к Google Sheets
        print('\n[1/4] 🔗 Подключение к Google Sheets API...')
        creds = authenticate()
        service = build('sheets', 'v4', credentials=creds)
        print('      ✅ Подключено')
        
        # 2. Получение списка операторов
        print('\n[2/4] 📋 Получение списка операторов...')
        operators = get_operator_sheets(service)
        print(f'      ✅ Найдено листов: {len(operators)}')
        
        # 3. Сбор данных
        print(f'\n[3/4] 📥 Сбор данных от операторов...')
        all_records = []
        
        for idx, operator in enumerate(operators, 1):
            print(f'      [{idx:2}/{len(operators)}] {operator[:60]:60}', end='', flush=True)
            
            records = collect_fiksa_data(operator, service)
            all_records.extend(records)
            
            print(f' → {len(records):,} записей')
        
        print(f'\n      📊 ИТОГО собрано: {len(all_records):,} записей')
        
        # 4. Сохранение в БД
        print(f'\n[4/4] 💾 Сохранение в базу данных...')
        inserted, updated = save_to_database(all_records)
        
        # Итоговая статистика
        print('\n' + '='*80)
        print('✅ ОБНОВЛЕНИЕ ЗАВЕРШЕНО УСПЕШНО')
        print('='*80)
        print(f'\n📊 РЕЗУЛЬТАТЫ:')
        print(f'   Операторов обработано: {len(operators)}')
        print(f'   Всего записей собрано: {len(all_records):,}')
        print(f'   ➕ Добавлено новых: {inserted:,}')
        print(f'   🔄 Обновлено существующих: {updated:,}')
        print(f'   💾 База данных: {DB_PATH}')
        print(f'   ⏰ Завершено: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        print('\n' + '='*80)
        
        return 0
        
    except FileNotFoundError as e:
        print(f'\n❌ ОШИБКА: Файл не найден - {e}')
        print('   Проверьте наличие config/credentials.json')
        return 1
        
    except Exception as e:
        print(f'\n❌ КРИТИЧЕСКАЯ ОШИБКА: {e}')
        import traceback
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
