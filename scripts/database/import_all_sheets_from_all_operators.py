#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт ВСЕХ листов из ВСЕХ таблиц 35 операторов в PostgreSQL
"""

import sys
import io
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import os
import socket

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

SKIP_SHEETS = {'настройки', 'статистика', 'сводка', 'тренды', 'итого', 'summary', 'settings', 'stats'}

print('\n' + '='*80)
print('ИМПОРТ ВСЕХ ЛИСТОВ ИЗ ВСЕХ ТАБЛИЦ ОПЕРАТОРОВ')
print('='*80 + '\n')

def get_service():
    creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    return build('sheets', 'v4', credentials=creds)

def get_operators():
    service = get_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SPREADSHEET_ID,
        range="Настройки!A2:C100"
    ).execute()
    
    operators = []
    for row in result.get('values', []):
        if len(row) >= 3 and row[2].strip().lower() == 'активен':
            operators.append({'name': row[0].strip(), 'id': row[1].strip()})
    return operators

def get_all_sheets(service, spreadsheet_id):
    """Получить ВСЕ листы из таблицы"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = []
        
        for sheet in spreadsheet.get('sheets', []):
            title = sheet['properties']['title']
            rows = sheet['properties']['gridProperties'].get('rowCount', 0)
            
            if title.lower() not in SKIP_SHEETS and rows > 10:
                sheets.append({'title': title, 'rows': rows})
        
        return sheets
    except Exception as e:
        print(f'  ⚠️  Ошибка получения листов: {e}')
        return []

def read_sheet_data(service, spreadsheet_id, sheet_name, operator_name):
    """Читает данные с одного листа"""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1:Z10000"
        ).execute()
        
        values = result.get('values', [])
        if not values or len(values) < 2:
            return []
        
        headers = [h.lower().strip() for h in values[0]]
        
        # Поиск колонок
        col_card = next((i for i, h in enumerate(headers) if 'код' in h and 'карта' in h), 0)
        col_phone = next((i for i, h in enumerate(headers) if 'тел' in h or 'номер' in h), 1)
        col_date = next((i for i, h in enumerate(headers) if 'дата' in h), 2)
        col_status = next((i for i, h in enumerate(headers) if 'статус' in h), 3)
        
        records = []
        for row in values[1:]:
            if not row or len(row) <= col_card:
                continue
            
            card = str(row[col_card]).strip() if col_card < len(row) else ''
            if not card:
                continue
            
            records.append({
                'operator': operator_name,
                'card': card,
                'phone': str(row[col_phone]) if col_phone < len(row) else '',
                'date': str(row[col_date]) if col_date < len(row) else '',
                'status': str(row[col_status]) if col_status < len(row) else '',
                'sheet': sheet_name
            })
        
        return records
    except HttpError as e:
        if e.resp.status == 429:
            time.sleep(60)
            return read_sheet_data(service, spreadsheet_id, sheet_name, operator_name)
        return []
    except Exception as e:
        return []

def parse_date(date_str):
    if not date_str:
        return None
    
    for fmt in ['%d.%m.%Y %H:%M', '%d.%m.%Y', '%Y-%m-%d']:
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
    errors = 0
    
    for rec in tqdm(records, desc='Импорт'):
        try:
            op_name = rec['operator']
            
            if op_name not in ops_cache:
                cur.execute(
                    'INSERT INTO operators (operator_name) VALUES (%s) ON CONFLICT (operator_name) DO UPDATE SET operator_name = EXCLUDED.operator_name RETURNING operator_id',
                    (op_name,)
                )
                ops_cache[op_name] = cur.fetchone()[0]
            
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
                            operator_id = EXCLUDED.operator_id
                    ''', batch)
                    conn.commit()
                    imported += len(batch)
                except Exception as e:
                    conn.rollback()
                    errors += len(batch)
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
                    operator_id = EXCLUDED.operator_id
            ''', batch)
            conn.commit()
            imported += len(batch)
        except:
            errors += len(batch)
    
    conn.close()
    return imported

def main():
    print('[1/4] Получение списка операторов...')
    operators = get_operators()
    print(f'✅ Операторов: {len(operators)}\n')
    
    print('[2/4] Сбор информации о листах...')
    service = get_service()
    
    total_sheets = 0
    all_records = []
    
    for i, op in enumerate(operators, 1):
        print(f'\n[{i}/{len(operators)}] {op["name"]}')
        
        sheets = get_all_sheets(service, op['id'])
        if not sheets:
            print('  Нет доступных листов')
            continue
        
        print(f'  Листов: {len(sheets)}')
        total_sheets += len(sheets)
        
        for sheet in sheets:
            print(f'    {sheet["title"]} ({sheet["rows"]} строк)...', end=' ')
            records = read_sheet_data(service, op['id'], sheet['title'], op['name'])
            all_records.extend(records)
            print(f'{len(records)} записей')
            time.sleep(0.5)
    
    print(f'\n✅ Всего листов: {total_sheets}')
    print(f'✅ Собрано записей: {len(all_records):,}\n')
    
    print('[3/4] Импорт в PostgreSQL...')
    imported = import_records(all_records)
    
    print(f'\n[4/4] Статистика...')
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    
    cur.execute('SELECT COUNT(*) FROM fixations')
    total = cur.fetchone()[0]
    print(f'📊 Всего записей в БД: {total:,}')
    
    cur.execute('SELECT COUNT(DISTINCT operator_id) FROM fixations')
    ops_count = cur.fetchone()[0]
    print(f'👥 Операторов: {ops_count}')
    
    cur.execute('SELECT MIN(call_date), MAX(call_date) FROM fixations WHERE call_date IS NOT NULL')
    min_d, max_d = cur.fetchone()
    if min_d:
        print(f'📅 Период: {min_d} - {max_d}')
    
    conn.close()
    
    print('\n' + '='*80)
    print(f'✅ ИМПОРТ ЗАВЕРШЕН! Импортировано: {imported:,}')
    print('='*80)

if __name__ == '__main__':
    main()
