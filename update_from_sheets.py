#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ОБНОВЛЕНИЕ ДАННЫХ ИЗ GOOGLE SHEETS
=============================================================================
Собирает все новые данные от операторов из Google Sheets и обновляет БД
=============================================================================
"""

import os
import sqlite3
from datetime import datetime
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import socket
import pandas as pd

# Прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
TOKEN_FILE = BASE_DIR / 'config' / 'token.json'
CREDENTIALS_FILE = BASE_DIR / 'config' / 'credentials.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

print('\n' + '='*80)
print('📥 ОБНОВЛЕНИЕ ДАННЫХ ИЗ GOOGLE SHEETS')
print('='*80)

# =============================================================================
# АУТЕНТИФИКАЦИЯ
# =============================================================================

def authenticate():
    """Аутентификация в Google API"""
    creds = None
    
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print('[AUTH] Обновление токена...')
            creds.refresh(Request())
        else:
            print('[AUTH] Первичная авторизация...')
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return creds

# =============================================================================
# ПОЛУЧЕНИЕ ДАННЫХ
# =============================================================================

def get_operator_sheets(service):
    """Получить список всех листов операторов"""
    spreadsheet = service.spreadsheets().get(spreadsheetId=MASTER_SPREADSHEET_ID).execute()
    sheets = spreadsheet.get('sheets', [])
    
    operator_sheets = []
    for sheet in sheets:
        title = sheet['properties']['title']
        # Пропускаем служебные листы
        if title not in ['Настройки', 'Статистика', 'Сводка', 'Тренды', 'Итого']:
            operator_sheets.append(title)
    
    return operator_sheets

def collect_operator_data(operator_name, service):
    """Собрать данные с листа оператора"""
    try:
        # Читаем весь лист FIKSA
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
        
        for idx, row in enumerate(values, start=2):
            # Пропускаем пустые строки
            if not row or len(row) < 5:
                continue
            
            # КРИТИЧНО: Колонка E (индекс 4) - статус
            status = row[4].strip() if len(row) > 4 and row[4] else ''
            if not status:
                continue  # Пропускаем строки без статуса
            
            # Колонка A - номер карты
            card_number = row[0].strip() if len(row) > 0 and row[0] else ''
            if not card_number:
                continue  # Пропускаем строки без номера карты
            
            # Парсим дату звонка (колонка F)
            call_date = None
            if len(row) > 5 and row[5]:
                try:
                    # Пробуем разные форматы
                    date_str = str(row[5]).strip()
                    for fmt in ['%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                        try:
                            call_date = datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
                            break
                        except:
                            continue
                except:
                    pass
            
            record = {
                'collection_date': today,
                'operator_name': operator_name,
                'card_number': card_number,
                'full_name': row[1].strip() if len(row) > 1 and row[1] else None,
                'phone': row[2].strip() if len(row) > 2 and row[2] else None,
                'address': row[3].strip() if len(row) > 3 and row[3] else None,
                'status': status,
                'call_date': call_date,
                'notes': row[6].strip() if len(row) > 6 and row[6] else None,
            }
            
            records.append(record)
        
        return records
        
    except Exception as e:
        print(f'   [ОШИБКА] {operator_name}: {e}')
        return []

# =============================================================================
# СОХРАНЕНИЕ В БД
# =============================================================================

def save_to_database(all_records):
    """Сохранить записи в БД (обновление существующих + добавление новых)"""
    if not all_records:
        print('[БД] Нет данных для сохранения')
        return 0, 0
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Создаем таблицу если не существует
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fiksa_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            collection_date DATE NOT NULL,
            operator_name TEXT NOT NULL,
            card_number TEXT,
            full_name TEXT,
            phone TEXT,
            address TEXT,
            status TEXT,
            call_date DATE,
            notes TEXT,
            region TEXT,
            district TEXT,
            incident_number TEXT,
            service_name TEXT,
            reason TEXT,
            description TEXT
        )
    ''')
    
    # Создаем индекс для быстрого поиска
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_card_operator 
        ON fiksa_records(card_number, operator_name)
    ''')
    
    updated = 0
    inserted = 0
    
    for record in all_records:
        # Проверяем существует ли запись
        cursor.execute('''
            SELECT id FROM fiksa_records 
            WHERE card_number = ? AND operator_name = ?
        ''', (record['card_number'], record['operator_name']))
        
        existing = cursor.fetchone()
        
        if existing:
            # Обновляем существующую запись
            cursor.execute('''
                UPDATE fiksa_records SET
                    collection_date = ?,
                    full_name = ?,
                    phone = ?,
                    address = ?,
                    status = ?,
                    call_date = ?,
                    notes = ?
                WHERE id = ?
            ''', (
                record['collection_date'],
                record['full_name'],
                record['phone'],
                record['address'],
                record['status'],
                record['call_date'],
                record['notes'],
                existing[0]
            ))
            updated += 1
        else:
            # Вставляем новую запись
            cursor.execute('''
                INSERT INTO fiksa_records (
                    collection_date, operator_name, card_number, full_name,
                    phone, address, status, call_date, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                record['collection_date'],
                record['operator_name'],
                record['card_number'],
                record['full_name'],
                record['phone'],
                record['address'],
                record['status'],
                record['call_date'],
                record['notes']
            ))
            inserted += 1
    
    conn.commit()
    conn.close()
    
    return inserted, updated

# =============================================================================
# ОСНОВНОЙ ПРОЦЕСС
# =============================================================================

try:
    # 1. Аутентификация
    print('\n[1/4] Подключение к Google Sheets...')
    creds = authenticate()
    service = build('sheets', 'v4', credentials=creds)
    print('  ✅ Подключено')
    
    # 2. Получение списка операторов
    print('\n[2/4] Получение списка операторов...')
    operators = get_operator_sheets(service)
    print(f'  Найдено операторов: {len(operators)}')
    
    # 3. Сбор данных
    print('\n[3/4] Сбор данных от операторов...')
    all_records = []
    
    for idx, operator in enumerate(operators, 1):
        print(f'  [{idx}/{len(operators)}] {operator[:50]:50}', end='', flush=True)
        
        records = collect_operator_data(operator, service)
        all_records.extend(records)
        
        print(f' → {len(records):,} записей')
    
    print(f'\n  Всего собрано: {len(all_records):,} записей')
    
    # 4. Сохранение в БД
    print('\n[4/4] Сохранение в базу данных...')
    inserted, updated = save_to_database(all_records)
    
    print('\n' + '='*80)
    print('✅ ОБНОВЛЕНИЕ ЗАВЕРШЕНО')
    print('='*80)
    print(f'\n📊 Результаты:')
    print(f'  Операторов обработано: {len(operators)}')
    print(f'  Всего записей: {len(all_records):,}')
    print(f'  ➕ Добавлено новых: {inserted:,}')
    print(f'  🔄 Обновлено существующих: {updated:,}')
    
    # Показываем топ операторов
    if all_records:
        df = pd.DataFrame(all_records)
        top_operators = df.groupby('operator_name').size().sort_values(ascending=False).head(5)
        
        print(f'\n🏆 ТОП-5 ОПЕРАТОРОВ ПО КОЛИЧЕСТВУ ЗАПИСЕЙ:')
        for operator, count in top_operators.items():
            print(f'  {operator[:50]:50} - {count:,} записей')
    
    print('\n' + '='*80)

except Exception as e:
    print(f'\n❌ ОШИБКА: {e}')
    import traceback
    traceback.print_exc()
