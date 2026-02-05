#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт данных из Google Sheets напрямую в PostgreSQL
"""

import sys
import io
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import os
import json
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

# Загрузка конфигурации
load_dotenv(CONFIG_DIR / 'postgresql.env')

DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', 5432)),
    'database': os.getenv('DB_NAME', 'qayta_data'),
    'user': os.getenv('DB_USER', 'qayta_user'),
    'password': os.getenv('DB_PASSWORD', 'qayta_password_2026')
}

# Google Sheets конфигурация
CREDENTIALS_FILE = CONFIG_DIR / 'credentials.json'
TOKEN_FILE = CONFIG_DIR / 'token.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

print('\n' + '='*80)
print('ИМПОРТ ИЗ GOOGLE SHEETS API В POSTGRESQL')
print('='*80)

def get_sheets_service():
    """Подключение к Google Sheets API с OAuth2"""
    if not CREDENTIALS_FILE.exists():
        print(f'\nОшибка: Файл credentials.json не найден в {CONFIG_DIR}')
        print('Скопируйте credentials.json в папку config/')
        return None
    
    try:
        creds = None
        # Загружаем сохраненный токен если есть
        if TOKEN_FILE.exists():
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        
        # Если нет валидных credentials, проходим OAuth flow
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                print('Обновление токена доступа...')
                creds.refresh(Request())
            else:
                print('Необходима авторизация в Google...')
                print('Откроется браузер для входа в Google аккаунт.')
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(CREDENTIALS_FILE), SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Сохраняем credentials для следующих запусков
            with open(TOKEN_FILE, 'w', encoding='utf-8') as token:
                token.write(creds.to_json())
            print('Токен сохранен в token.json')
        
        service = build('sheets', 'v4', credentials=creds)
        return service
    except Exception as e:
        print(f'Ошибка подключения к Google Sheets API: {e}')
        return None

def get_spreadsheets_list():
    """Получить список всех таблиц из info.json"""
    info_file = BASE_DIR / 'docs' / 'info.json'
    
    if not info_file.exists():
        print(f'\nОшибка: Файл info.json не найден')
        return []
    
    try:
        with open(info_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        spreadsheets = []
        for operator_name, sheets_list in data.items():
            if isinstance(sheets_list, list):
                for sheet_info in sheets_list:
                    if isinstance(sheet_info, dict) and 'url' in sheet_info:
                        # Извлекаем ID из URL
                        url = sheet_info['url']
                        if '/spreadsheets/d/' in url:
                            sheet_id = url.split('/spreadsheets/d/')[1].split('/')[0]
                            spreadsheets.append({
                                'operator': operator_name,
                                'id': sheet_id,
                                'name': sheet_info.get('name', 'Без названия')
                            })
        
        return spreadsheets
    
    except Exception as e:
        print(f'Ошибка чтения info.json: {e}')
        return []

def read_sheet_data(service, spreadsheet_id, operator_name):
    """Читает все данные из таблицы"""
    try:
        # Получаем информацию о листах
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', [])
        
        all_data = []
        
        for sheet in sheets:
            sheet_name = sheet['properties']['title']
            
            # Пропускаем служебные листы
            if sheet_name.lower() in ['logs', 'лист4', 'лист9']:
                continue
            
            # Читаем данные
            range_name = f"'{sheet_name}'!A:Z"
            
            try:
                result = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_name
                ).execute()
                
                values = result.get('values', [])
                
                if not values or len(values) < 2:
                    continue
                
                # Первая строка - заголовки
                headers = values[0]
                
                # Определяем индексы нужных колонок
                card_idx = None
                status_idx = None
                date_idx = None
                phone_idx = None
                
                for i, header in enumerate(headers):
                    header_lower = str(header).lower()
                    if 'код карты' in header_lower or 'номер карты' in header_lower:
                        card_idx = i
                    elif 'статус' in header_lower and 'связ' in header_lower:
                        status_idx = i
                    elif 'дата' in header_lower and ('откр' in header_lower or 'фикс' in header_lower):
                        date_idx = i
                    elif 'тел' in header_lower and 'номер' in header_lower:
                        phone_idx = i
                
                if card_idx is None:
                    continue
                
                # Обрабатываем данные
                for row in values[1:]:
                    if not row or len(row) <= card_idx:
                        continue
                    
                    card_number = row[card_idx] if card_idx < len(row) else None
                    if not card_number or str(card_number).strip() == '':
                        continue
                    
                    status = row[status_idx] if status_idx is not None and status_idx < len(row) else None
                    call_date = row[date_idx] if date_idx is not None and date_idx < len(row) else None
                    phone = row[phone_idx] if phone_idx is not None and phone_idx < len(row) else None
                    
                    all_data.append({
                        'card_number': str(card_number).strip(),
                        'operator': operator_name,
                        'status': str(status).strip() if status else None,
                        'call_date': str(call_date).strip() if call_date else None,
                        'phone': str(phone).strip() if phone else None,
                        'sheet_name': sheet_name
                    })
                
                # Небольшая задержка между запросами
                time.sleep(0.3)
                
            except HttpError as e:
                if e.resp.status == 429:  # Too Many Requests
                    print(f'\n  Превышен лимит запросов, ждем 60 секунд...')
                    time.sleep(60)
                continue
        
        return all_data
    
    except Exception as e:
        print(f'\n  Ошибка чтения таблицы: {e}')
        return []

def get_or_create_operator(cursor, operator_name):
    """Получить или создать оператора"""
    if not operator_name:
        return None
    
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

def parse_date(date_str):
    """Парсинг даты из разных форматов"""
    if not date_str:
        return None
    
    import pandas as pd
    try:
        return pd.to_datetime(date_str, dayfirst=True)
    except:
        return None

def import_to_postgresql(data_list):
    """Импорт данных в PostgreSQL"""
    if not data_list:
        print('\nНет данных для импорта')
        return 0
    
    print(f'\n\nНайдено записей: {len(data_list):,}')
    print('\nИмпорт в PostgreSQL...')
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        # Очистка таблицы
        print('  Очистка старых данных...')
        cursor.execute("TRUNCATE TABLE fixations RESTART IDENTITY CASCADE")
        conn.commit()
        
        # Группируем по операторам
        operators_data = {}
        for item in data_list:
            op = item['operator']
            if op not in operators_data:
                operators_data[op] = []
            operators_data[op].append(item)
        
        imported = 0
        errors = 0
        
        print(f'\n  Импорт данных ({len(operators_data)} операторов)...')
        
        for operator_name, items in tqdm(operators_data.items(), desc='  Операторы'):
            operator_id = get_or_create_operator(cursor, operator_name)
            
            batch = []
            for item in items:
                try:
                    call_date = parse_date(item['call_date'])
                    
                    batch.append((
                        item['card_number'],
                        operator_id,
                        call_date,
                        item['status'],
                        item['phone'],
                        item['sheet_name'],
                        datetime.now()
                    ))
                    
                    if len(batch) >= 5000:
                        try:
                            cursor.executemany("""
                                INSERT INTO fixations 
                                (card_number, operator_id, call_date, status, phone, source_file, import_date)
                                VALUES (%s, %s, %s, %s, %s, %s, %s)
                            """, batch)
                            conn.commit()
                            imported += len(batch)
                        except Exception as e:
                            conn.rollback()
                            errors += len(batch)
                        batch = []
                
                except Exception as e:
                    errors += 1
            
            # Остаток
            if batch:
                try:
                    cursor.executemany("""
                        INSERT INTO fixations 
                        (card_number, operator_id, call_date, status, phone, source_file, import_date)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """, batch)
                    conn.commit()
                    imported += len(batch)
                except Exception as e:
                    conn.rollback()
                    errors += len(batch)
        
        cursor.close()
        conn.close()
        
        print(f'\n  Импортировано: {imported:,}')
        if errors > 0:
            print(f'  Ошибок: {errors}')
        
        return imported
    
    except Exception as e:
        print(f'\nОшибка импорта в PostgreSQL: {e}')
        return 0

def show_statistics():
    """Показать статистику"""
    print('\n' + '='*80)
    print('СТАТИСТИКА')
    print('='*80)
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM fixations")
        total = cursor.fetchone()[0]
        print(f'\nВсего записей: {total:,}')
        
        cursor.execute("""
            SELECT operator_name, total_fixations, positive_percentage
            FROM v_operator_statistics 
            WHERE total_fixations > 0
            ORDER BY total_fixations DESC 
            LIMIT 10
        """)
        
        operators = cursor.fetchall()
        if operators:
            print('\nТОП-10 операторов:')
            for i, (name, count, pct) in enumerate(operators, 1):
                print(f'  {i:2}. {name:<40} {count:>7,} ({pct or 0:>5.1f}%)')
        
        cursor.execute("""
            SELECT status_category, COUNT(*) 
            FROM fixations 
            WHERE status_category IS NOT NULL
            GROUP BY status_category 
            ORDER BY COUNT(*) DESC
        """)
        
        categories = cursor.fetchall()
        if categories:
            print('\nПо категориям:')
            for cat, count in categories:
                pct = (count / total * 100) if total > 0 else 0
                print(f'  {cat:<20} {count:>10,} ({pct:>5.1f}%)')
        
        cursor.close()
        conn.close()
        print('\n' + '='*80)
    
    except Exception as e:
        print(f'Ошибка статистики: {e}')

def main():
    """Главная функция"""
    
    # Подключение к Google Sheets API
    print('\n[1/4] Подключение к Google Sheets API...')
    service = get_sheets_service()
    
    if not service:
        return
    
    print('  OK')
    
    # Получение списка таблиц
    print('\n[2/4] Получение списка таблиц из info.json...')
    spreadsheets = get_spreadsheets_list()
    
    if not spreadsheets:
        print('  Нет таблиц для импорта')
        return
    
    print(f'  Найдено таблиц: {len(spreadsheets)}')
    
    # Чтение данных
    print('\n[3/4] Чтение данных из Google Sheets...')
    all_data = []
    
    for sheet_info in tqdm(spreadsheets, desc='  Таблицы'):
        data = read_sheet_data(service, sheet_info['id'], sheet_info['operator'])
        all_data.extend(data)
    
    print(f'\n  Прочитано записей: {len(all_data):,}')
    
    # Импорт в PostgreSQL
    print('\n[4/4] Импорт в PostgreSQL...')
    imported = import_to_postgresql(all_data)
    
    if imported > 0:
        show_statistics()
        
        print('\n' + '='*80)
        print('ИМПОРТ ЗАВЕРШЕН!')
        print('='*80)
        print(f'\nИмпортировано: {imported:,} записей')
        print('\nДоступ к данным:')
        print('  - pgAdmin:    http://localhost:5050')
        print('  - PostgreSQL: localhost:5432')
        print('='*80 + '\n')

if __name__ == '__main__':
    main()
