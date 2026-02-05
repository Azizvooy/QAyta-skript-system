#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт ТОЛЬКО релевантных листов из ВСЕХ 35 операторов в PostgreSQL
Фильтруются листы с реальными данными звонков:
- FIKSA (основной лист)
- ФИО + дата (напр. "Narziyeva Gavxar Atxamjanovna 09.2025")
- FIKSA (...) (напр. "FIKSA (Narziyeva Gavxar 20.06.2025)")
"""

import sys
import io
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import os
import socket
import re

# Прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

import psycopg2
from psycopg2.extras import execute_batch
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tqdm import tqdm
import time

BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / 'config'

load_dotenv(CONFIG_DIR / 'postgresql.env')

DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', 5432)),
    'database': os.getenv('DB_NAME', 'qayta_data'),
    'user': os.getenv('DB_USER', 'qayta_user'),
    'password': os.getenv('DB_PASSWORD', 'qayta_password_2026')
}

TOKEN_FILE = CONFIG_DIR / 'token.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

print('\n' + '='*80)
print('ИМПОРТ РЕЛЕВАНТНЫХ ЛИСТОВ ИЗ ВСЕХ 35 ОПЕРАТОРОВ')
print('='*80 + '\n')

def get_service():
    creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    return build('sheets', 'v4', credentials=creds)

def get_operators():
    """Получить список ВСЕх 35 операторов"""
    service = get_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SPREADSHEET_ID,
        range="Настройки!A2:C100"
    ).execute()
    
    operators = []
    for row in result.get('values', []):
        if len(row) >= 2:
            # Важно: берём ВСЕХ операторов, даже с неактивным статусом
            operators.append({'name': row[0].strip(), 'id': row[1].strip()})
    return operators

def is_relevant_sheet(sheet_name):
    """
    Проверить, является ли лист релевантным для импорта данных
    Релевантные: FIKSA, листы с ФИО+датой, FIKSA(...)
    """
    title = str(sheet_name).strip()
    
    # Основной лист FIKSA
    if title.upper() == 'FIKSA':
        return True
    
    # FIKSA (...) - листы с реализацией
    if title.startswith('FIKSA (') and title.endswith(')'):
        return True
    
    # Листы с ФИО + дата (напр. "Narziyeva Gavxar Atxamjanovna 09.2025")
    # Паттерн: слова (ФИО) + пробел + цифры/точки (дата)
    date_pattern = r'\d{2}\.\d{4}|\d{1,2}\.\d{1,2}\.\d{4}'
    if re.search(date_pattern, title):
        return True
    
    return False

def get_data_sheets(service, spreadsheet_id):
    """Получить ТОЛЬКО релевантные листы с данными"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = []
        
        for sheet in spreadsheet.get('sheets', []):
            title = sheet['properties']['title']
            rows = sheet['properties']['gridProperties'].get('rowCount', 0)
            
            # Фильтруем только релевантные листы с данными
            if is_relevant_sheet(title) and rows > 5:
                sheets.append({'title': title, 'rows': rows})
        
        return sheets
    except Exception as e:
        print(f'  ⚠️  Ошибка получения листов: {e}')
        return []

def read_sheet_data(service, spreadsheet_id, sheet_name, operator_name):
    """Читает данные с одного листа (все строки)"""
    try:
        # Получаем ВСЕ данные листа (без ограничений)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A:Z"
        ).execute()
        
        values = result.get('values', [])
        if not values or len(values) < 2:
            return []
        
        headers = [h.lower().strip() for h in values[0]]
        
        # Поиск колонок - ищем по ключевым словам
        col_card = None
        col_phone = None
        col_date = None
        col_status = None
        
        # Find card number column
        for i, h in enumerate(headers):
            if 'код' in h and 'карта' in h:
                col_card = i
            elif 'карта' in h and 'номер' in h:
                col_card = i
            elif h == 'код':
                col_card = i
        
        # Find phone column
        for i, h in enumerate(headers):
            if 'тел' in h or 'номер' in h or 'phone' in h:
                col_phone = i
                break
        
        # Find date column
        for i, h in enumerate(headers):
            if 'дата' in h or 'date' in h or 'time' in h:
                col_date = i
                break
        
        # Find status column
        for i, h in enumerate(headers):
            if 'статус' in h or 'status' in h:
                col_status = i
                break
        
        # Значения по умолчанию, если колонки не найдены
        if col_card is None:
            col_card = 0
        if col_phone is None:
            col_phone = 1 if len(headers) > 1 else 0
        if col_date is None:
            col_date = 2 if len(headers) > 2 else 1
        if col_status is None:
            col_status = 3 if len(headers) > 3 else 2
        
        records = []
        for row in values[1:]:
            if not row or len(row) <= col_card:
                continue
            
            card = str(row[col_card]).strip() if col_card < len(row) else ''
            if not card or card.upper() in ['КОД', 'CODE', '-', '']:
                continue
            
            records.append({
                'operator': operator_name,
                'card': card,
                'phone': str(row[col_phone]).strip() if col_phone < len(row) else '',
                'date': str(row[col_date]).strip() if col_date < len(row) else '',
                'status': str(row[col_status]).strip() if col_status < len(row) else '',
                'sheet': sheet_name
            })
        
        return records
    except HttpError as e:
        if e.resp.status == 429:
            time.sleep(60)
            return read_sheet_data(service, spreadsheet_id, sheet_name, operator_name)
        return []
    except Exception:
        return []

def parse_date(date_str):
    if not date_str or str(date_str).upper() in ['', '-', 'ДАТА']:
        return None
    
    for fmt in ['%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M', '%d.%m.%Y', '%Y-%m-%d', '%d.%m.%y']:
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except:
            continue
    return None

def import_records(records):
    if not records:
        return 0
    
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    
    # Кеш операторов
    ops_cache = {}
    cur.execute('SELECT operator_name, operator_id FROM operators')
    for name, oid in cur.fetchall():
        ops_cache[name] = oid
    
    imported = 0
    batch = []
    
    for rec in tqdm(records, desc='Импорт', leave=False):
        try:
            op_name = rec['operator']
            
            if op_name not in ops_cache:
                try:
                    cur.execute(
                        'INSERT INTO operators (operator_name) VALUES (%s) RETURNING operator_id',
                        (op_name,)
                    )
                    ops_cache[op_name] = cur.fetchone()[0]
                except:
                    continue
            
            batch.append((
                rec['card'],
                ops_cache[op_name],
                parse_date(rec.get('date')),
                rec.get('phone', ''),
                rec.get('status', ''),
                rec.get('sheet', '')
            ))
            
            if len(batch) >= 2000:
                try:
                    execute_batch(cur, '''
                        INSERT INTO fixations (card_number, operator_id, call_date, phone, status, description)
                        VALUES (%s, %s, %s, %s, %s, %s)
                        ON CONFLICT (card_number, call_date) DO UPDATE SET
                            phone = EXCLUDED.phone,
                            status = EXCLUDED.status,
                            operator_id = EXCLUDED.operator_id,
                            description = EXCLUDED.description
                    ''', batch)
                    conn.commit()
                    imported += len(batch)
                except Exception as e:
                    conn.rollback()
                batch = []
        except:
            continue
    
    if batch:
        try:
            execute_batch(cur, '''
                INSERT INTO fixations (card_number, operator_id, call_date, phone, status, description)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (card_number, call_date) DO UPDATE SET
                    phone = EXCLUDED.phone,
                    status = EXCLUDED.status,
                    operator_id = EXCLUDED.operator_id,
                    description = EXCLUDED.description
            ''', batch)
            conn.commit()
            imported += len(batch)
        except:
            pass
    
    conn.close()
    return imported

def main():
    print('[1/4] Получение списка ВСЕх 35 операторов...')
    operators = get_operators()
    print(f'✅ Найдено операторов: {len(operators)}\n')
    
    print('[2/4] Сбор информации о релевантных листах...')
    service = get_service()
    
    total_sheets = 0
    all_records = []
    
    for i, op in enumerate(operators, 1):
        op_name = op['name'] if op['name'].strip() else f'Operator {i}'
        print(f'\n[{i}/{len(operators)}] {op_name}')
        
        sheets = get_data_sheets(service, op['id'])
        if not sheets:
            print('  ⚠️  Нет релевантных листов')
            continue
        
        print(f'  ✓ Релевантных листов: {len(sheets)}')
        total_sheets += len(sheets)
        
        for sheet in sheets:
            print(f'    • {sheet["title"]}...', end='', flush=True)
            records = read_sheet_data(service, op['id'], sheet['title'], op_name)
            all_records.extend(records)
            print(f' {len(records)} записей')
            time.sleep(0.3)
    
    print(f'\n✅ Всего релевантных листов: {total_sheets}')
    print(f'✅ Собрано записей: {len(all_records):,}\n')
    
    if not all_records:
        print('❌ Нет данных для импорта!')
        return
    
    print('[3/4] Импорт в PostgreSQL...')
    imported = import_records(all_records)
    
    print(f'\n[4/4] Финальная статистика...')
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    
    cur.execute('SELECT COUNT(*) FROM fixations')
    total = cur.fetchone()[0]
    print(f'📊 Всего записей в БД: {total:,}')
    
    cur.execute('SELECT COUNT(DISTINCT operator_id) FROM fixations')
    ops_count = cur.fetchone()[0]
    print(f'👥 Операторов с данными: {ops_count}')
    
    cur.execute('SELECT MIN(call_date), MAX(call_date) FROM fixations WHERE call_date IS NOT NULL')
    result = cur.fetchone()
    if result and result[0]:
        min_d, max_d = result
        print(f'📅 Период: {min_d} - {max_d}')
    
    cur.execute('SELECT status, COUNT(*) FROM fixations WHERE status IS NOT NULL GROUP BY status ORDER BY COUNT(*) DESC LIMIT 5')
    print('\n📈 Топ статусы:')
    for status, count in cur.fetchall():
        print(f'  • {status}: {count:,}')
    
    conn.close()
    
    print('\n' + '='*80)
    print(f'✅ ИМПОРТ ЗАВЕРШЕН!')
    print(f'   Импортировано: {imported:,}')
    print(f'   Всего в БД: {total:,}')
    print('='*80)

if __name__ == '__main__':
    main()
