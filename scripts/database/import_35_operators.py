#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт данных всех 35 операторов из индивидуальных Google Sheets таблиц в PostgreSQL
"""

import sys
import io
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import os
import socket

# Прокси настройки (обязательно!)
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

import psycopg2
from psycopg2.extras import execute_batch
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
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

CREDENTIALS_FILE = CONFIG_DIR / 'credentials.json'
TOKEN_FILE = CONFIG_DIR / 'token.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

print('\n' + '='*80)
print('ИМПОРТ ДАННЫХ ВСЕХ ОПЕРАТОРОВ ИЗ 35 GOOGLE SHEETS ТАБЛИЦ')
print('='*80 + '\n')

def get_sheets_service():
    """Подключение к Google Sheets API"""
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
    
    return build('sheets', 'v4', credentials=creds)

def get_operators_list(service):
    """Читает список операторов из мастер-таблицы"""
    result = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SPREADSHEET_ID,
        range="Настройки!A2:C100"
    ).execute()
    
    values = result.get('values', [])
    operators = []
    
    for row in values:
        if len(row) >= 3:
            name = row[0].strip()
            spreadsheet_id = row[1].strip()
            status = row[2].strip().lower()
            
            if status == 'активен' and spreadsheet_id:
                operators.append({'name': name, 'spreadsheet_id': spreadsheet_id})
    
    return operators

def read_operator_data(service, operator):
    """Читает данные из таблицы оператора"""
    name = operator['name']
    spreadsheet_id = operator['spreadsheet_id']
    
    try:
        # Получаем список листов
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        # Находим лист с данными (максимальное количество строк)
        data_sheet = None
        max_rows = 0
        
        skip_sheets = {'настройки', 'статистика', 'сводка', 'тренды', 'итого', 'summary'}
        
        for sheet in sheets:
            title = sheet['properties']['title']
            rows = sheet['properties']['gridProperties'].get('rowCount', 0)
            
            if title.lower() not in skip_sheets and rows > max_rows:
                max_rows = rows
                data_sheet = title
        
        if not data_sheet:
            data_sheet = sheets[0]['properties']['title']
        
        # Читаем данные
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{data_sheet}'!A1:Z10000"
        ).execute()
        
        values = result.get('values', [])
        if not values or len(values) < 2:
            return []
        
        headers = [h.lower().strip() for h in values[0]]
        
        # Находим колонки
        col_card = next((i for i, h in enumerate(headers) if 'код' in h and 'карта' in h), 0)
        col_phone = next((i for i, h in enumerate(headers) if 'тел' in h or 'номер' in h), 1)
        col_date = next((i for i, h in enumerate(headers) if 'дата' in h), 2)
        col_status = next((i for i, h in enumerate(headers) if 'статус' in h), 3)
        col_service = next((i for i, h in enumerate(headers) if 'служб' in h or 'сервис' in h), 4)
        col_comments = next((i for i, h in enumerate(headers) if 'комментар' in h), 5)
        
        records = []
        for row in values[1:]:
            if not row or len(row) <= col_card:
                continue
            
            card = row[col_card].strip() if col_card < len(row) else ''
            if not card:
                continue
            
            records.append({
                'operator_name': name,
                'card_number': card,
                'phone_number': row[col_phone] if col_phone < len(row) else '',
                'call_date': row[col_date] if col_date < len(row) else '',
                'status': row[col_status] if col_status < len(row) else '',
                'service': row[col_service] if col_service < len(row) else '',
                'comments': row[col_comments] if col_comments < len(row) else '',
            })
        
        return records
        
    except HttpError as e:
        if e.resp.status == 403:
            print(f'    ⚠️  Нет доступа к таблице {name}')
        elif e.resp.status == 429:
            print(f'    ⚠️  Rate limit, пауза 60 сек...')
            time.sleep(60)
            return read_operator_data(service, operator)
        else:
            print(f'    ⚠️  HTTP {e.resp.status}: {e.reason}')
        return []
    except Exception as e:
        print(f'    ⚠️  Ошибка: {e}')
        return []

def parse_date(date_str):
    """Парсинг даты"""
    if not date_str or not str(date_str).strip():
        return None
    
    date_str = str(date_str).strip()
    formats = ['%d.%m.%Y %H:%M', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y']
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
    return None

def import_to_postgresql(all_records):
    """Импорт в PostgreSQL"""
    if not all_records:
        print('Нет данных для импорта')
        return 0
    
    print(f'\n💾 Импорт {len(all_records):,} записей...')
    
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    
    # Кеш операторов
    operators_cache = {}
    cur.execute('SELECT operator_name, operator_id FROM operators')
    for name, op_id in cur.fetchall():
        operators_cache[name] = op_id
    
    imported = 0
    batch = []
    
    for record in tqdm(all_records, desc='Импорт'):
        try:
            op_name = record['operator_name']
            
            if op_name not in operators_cache:
                cur.execute(
                    'INSERT INTO operators (operator_name) VALUES (%s) ON CONFLICT (operator_name) DO UPDATE SET operator_name = EXCLUDED.operator_name RETURNING operator_id',
                    (op_name,)
                )
                operators_cache[op_name] = cur.fetchone()[0]
            
            batch.append((
                record['card_number'],
                operators_cache[op_name],
                parse_date(record.get('call_date')),
                record.get('phone_number', ''),
                record.get('status', ''),
                record.get('comments', ''),
            ))
            
            if len(batch) >= 5000:
                try:
                    execute_batch(cur, '''
                        INSERT INTO fixations (card_number, operator_id, call_date, phone, status, description)
                        VALUES (%s, %s, %s, %s, %s, %s)
                        ON CONFLICT (card_number, call_date) DO UPDATE SET
                            phone = EXCLUDED.phone,
                            status = EXCLUDED.status,
                            description = EXCLUDED.description,
                            operator_id = EXCLUDED.operator_id
                    ''', batch)
                    conn.commit()
                    imported += len(batch)
                except Exception as e:
                    conn.rollback()
                    print(f'\n⚠️  Ошибка батча: {e}')
                batch = []
        except Exception as e:
            continue
    
    if batch:
        try:
            execute_batch(cur, '''
                INSERT INTO fixations (card_number, phone_number, call_date, call_status, service_name, comments, operator_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)operator_id, call_date, phone, status, description)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (card_number, call_date) DO UPDATE SET
                    phone = EXCLUDED.phone,
                    status = EXCLUDED.status,
                    description = EXCLUDED.descriptiontor_id
            ''', batch)
            conn.commit()
            imported += len(batch)
        except Exception as e:
            conn.rollback()
            print(f'\n⚠️  Ошибка финального батча: {e}')
    
    conn.close()
    return imported

def show_stats():
    """Статистика"""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    
    cur.execute('SELECT COUNT(*) FROM fixations')
    total = cur.fetchone()[0]
    print(f'\n📊 Всего записей: {total:,}')
    
    cur.execute('SELECT status_category, COUNT(*) FROM fixations GROUP BY status_category ORDER BY COUNT(*) DESC')
    print('\n📈 По категориям:')
    for cat, cnt in cur.fetchall():
        pct = (cnt/total*100) if total > 0 else 0
        print(f'   {cat}: {cnt:,} ({pct:.1f}%)')
    
    cur.execute('SELECT COUNT(DISTINCT operator_id) FROM fixations')
    print(f'\n👥 Операторов: {cur.fetchone()[0]}')
    
    conn.close()

def main():
    print('[1/4] Подключение к Google Sheets API...')
    service = get_sheets_service()
    print('✅ Подключено\n')
    
    print('[2/4] Получение списка операторов...')
    operators = get_operators_list(service)
    print(f'✅ Найдено операторов: {len(operators)}\n')
    
    print('[3/4] Сбор данных из таблиц...')
    all_records = []
    
    for i, op in enumerate(operators, 1):
        print(f'  [{i}/{len(operators)}] {op["name"]}...', end=' ')
        records = read_operator_data(service, op)
        all_records.extend(records)
        print(f'{len(records):,} записей')
        time.sleep(1)
    
    print(f'\n✅ Собрано: {len(all_records):,} записей\n')
    
    print('[4/4] Импорт в PostgreSQL...')
    imported = import_to_postgresql(all_records)
    
    if imported > 0:
        show_stats()
        print('\n' + '='*80)
        print(f'✅ ИМПОРТ ЗАВЕРШЕН! Импортировано: {imported:,}')
        print('='*80)

if __name__ == '__main__':
    main()
