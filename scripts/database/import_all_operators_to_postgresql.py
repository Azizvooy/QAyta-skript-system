#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт всех данных операторов из Google Sheets в PostgreSQL
Читает 35 листов из мастер-таблицы и импортирует в БД
"""

import sys
import io
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import os
import socket

# Прокси настройки
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

# ID мастер-таблицы со всеми операторами
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

# Служебные листы, которые не содержат данных операторов
SKIP_SHEETS = ['Настройки', 'Статистика', 'Сводка', 'Тренды', 'Итого', 'Summary', 'Settings']

print('\n' + '='*80)
print('МАССОВЫЙ ИМПОРТ ДАННЫХ ВСЕХ ОПЕРАТОРОВ В POSTGRESQL')
print('='*80)
print(f'Мастер-таблица: {MASTER_SPREADSHEET_ID}')
print(f'Credentials: {CREDENTIALS_FILE}')
print()

def get_sheets_service():
    """Подключение к Google Sheets API с OAuth2"""
    if not CREDENTIALS_FILE.exists():
        print(f'\nОшибка: Файл credentials.json не найден в {CONFIG_DIR}')
        return None
    
    try:
        creds = None
        if TOKEN_FILE.exists():
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        
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
            
            with open(TOKEN_FILE, 'w', encoding='utf-8') as token:
                token.write(creds.to_json())
            print('✅ Токен сохранен')
        
        service = build('sheets', 'v4', credentials=creds)
        return service
    except Exception as e:
        print(f"Ошибка подключения к Google Sheets API: {e}")
        import traceback
        traceback.print_exc()
        return None

def get_operator_sheets(service):
    """Получить список всех листов операторов"""
    try:
        print('Запрос метаданных таблицы...')
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=MASTER_SPREADSHEET_ID).execute()
        
        title = spreadsheet.get('properties', {}).get('title', 'Неизвестная таблица')
        print(f'Таблица: {title}')
        
        sheets = spreadsheet.get('sheets', [])
        print(f'Всего листов: {len(sheets)}')
        
        operator_sheets = []
        for sheet in sheets:
            sheet_title = sheet['properties']['title']
            if sheet_title not in SKIP_SHEETS:
                operator_sheets.append(sheet_title)
            else:
                print(f'  Пропускаем служебный лист: {sheet_title}')
        
        return operator_she1:Z (с заголовками и до конца)
        range_name = f"'{operator_name}'!A1:Z10000"
        
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=range_name,
            valueRenderOption='UNFORMATTED_VALUE',
            dateTimeRenderOption='FORMATTED_STRING'с доступом. Проверьте что таблица открыта для вашего Google аккаунта.')
        elif e.resp.status == 404:
            print('Таблица не найдена. Проверьте ID таблицы.')
        import traceback
        traceback.print_exc()
        return []
    except Exception as e:
        print(f'❌ Ошибка получения списка листов: {e}')
        import traceback
        traceback.print_exc()
        return []

def read_operator_sheet(service, operator_name):
    """Читает данные с листа оператора"""
    try:
        # Читаем диапазон A2:Z (с заголовками и до конца)
        range_name = f"'{operator_name}'!A1:Z10000"
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values or len(values) < 2:
            return []
        
        # Первая строка - заголовки
        headers = values[0]
        data_rows = values[1:]
        
        # Ищем нужные колонки
        col_indices = {}
        for i, header in enumerate(headers):
            header_lower = header.lower().strip()
            if 'код' in header_lower and 'карта' in header_lower:
                col_indices['card_number'] = i
            elif 'тел' in header_lower or 'номер' in header_lower:
                col_indices['phone_number'] = i
            elif 'дата' in header_lower:
                col_indices['call_date'] = i
            elif 'статус' in header_lower:
                col_indices['status'] = i
            elif 'служб' in header_lower or 'сервис' in header_lower:
                col_indices['service'] = i
            elif 'комментар' in header_lower:
                col_indices['comments'] = i
        
        # Собираем записи
        records = []
        for row in data_rows:
            if not row or len(row) == 0:
                continue
            
            # Проверяем что есть хотя бы код карты
            card_number = row[col_indices.get('card_number', 0)] if col_indices.get('card_number', 0) < len(row) else ''
            if not card_number or card_number.strip() == '':
                continue
            
            record = {
                'operator_name': operator_name,
                'card_number': card_number.strip(),
                'phone_number': row[col_indices.get('phone_number', 1)] if col_indices.get('phone_number', 1) < len(row) else '',
                'call_date': row[col_indices.get('call_date', 2)] if col_indices.get('call_date', 2) < len(row) else '',
                'status': row[col_indices.get('status', 3)] if col_indices.get('status', 3) < len(row) else '',
                'service': row[col_indices.get('service', 4)] if col_indices.get('service', 4) < len(row) else '',
                'comments': row[col_indices.get('comments', 5)] if col_indices.get('comments', 5) < len(row) else '',
            }
            records.append(record)
        
        return records
        
    except HttpError as e:
        if e.resp.status == 429:
            print(f'⚠️  Rate limit для {operator_name}, ждем 60 секунд...')
            time.sleep(60)
            return read_operator_sheet(service, operator)
        elif e.resp.status == 403:
            print(f'⚠️  Нет доступа к таблице {operator_name} (ID: {spreadsheet_id[:20]}...)')
            return []
        else:
            print(f'⚠️  Ошибка чтения таблицы {operator_name}: {e}')
            return []
    except Exception as e:
        print(f'⚠️  Ошибка обработки таблицы {operator_name}: {e}')
        return []

def parse_date(date_str):
    """Парсит дату из различных форматов"""
    if not date_str or date_str.strip() == '':
        return None
    
    date_str = date_str.strip()
    
    # Пробуем разные форматы
    formats = [
        '%d.%m.%Y %H:%M',
        '%d.%m.%Y',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d',
        '%d/%m/%Y %H:%M',
        '%d/%m/%Y',
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
    
    return None

def import_to_postgresql(all_records):
    """Импорт всех записей в PostgreSQL"""
    if not all_records:
        print('Нет данных для импорта')
        return 0
    
    print(f'\n💾 Импорт {len(all_records):,} записей в PostgreSQL...')
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        
        # Получаем кеш операторов
        operators_cache = {}
        cur.execute('SELECT name, id FROM operators')
        for name, op_id in cur.fetchall():
            operators_cache[name] = op_id
        
        imported_count = 0
        skipped_count = 0
        error_count = 0
        batch_data = []
        
        for record in tqdm(all_records, desc='Импорт записей'):
            try:
                operator_name = record['operator_name']
                card_number = record['card_number']
                
                if not operator_name or not card_number:
                    skipped_count += 1
                    continue
                
                # Создаем оператора если нужно
                if operator_name not in operators_cache:
                    cur.execute(
                        'INSERT INTO operators (name) VALUES (%s) ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name RETURNING id',
                        (operator_name,)
                    )
                    operators_cache[operator_name] = cur.fetchone()[0]
                
                operator_id = operators_cache[operator_name]
                
                # Парсим дату
                call_date = parse_date(record.get('call_date'))
                
                batch_data.append((
                    card_number,
                    record.get('phone_number', ''),
                    call_date,
                    record.get('status', ''),
                    record.get('service', ''),
                    record.get('comments', ''),
                    operator_id
                ))
                
                # Коммитим батчами по 5000
                if len(batch_data) >= 5000:
                    execute_batch(cur, '''
                        INSERT INTO fixations (card_number, phone_number, call_date, call_status, service_name, comments, operator_id)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        ON CONFLICT (card_number, call_date) DO UPDATE SET
                            phone_number = EXCLUDED.phone_number,
                            call_status = EXCLUDED.call_status,
                            service_name = EXCLUDED.service_name,
                            comments = EXCLUDED.comments,
                            operator_id = EXCLUDED.operator_id
                    ''', batch_data)
                    imported_count += len(batch_data)
                    batch_data = []
                    conn.commit()
            
            except Exception as e:
                error_count += 1
                if error_count <= 5:
                    print(f'\n⚠️  Ошибка: {e}')
                continue
        
        # Импортируем остаток
        if batch_data:
            execute_batch(cur, '''
                INSERT INTO fixations (card_number, phone_number, call_date, call_status, service_name, comments, operator_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (card_number, call_date) DO UPDATE SET
                    phone_number = EXCLUDED.phone_number,
                    call_status = EXCLUDED.call_status,
                    service_name = EXCLUDED.service_name,
                    comments = EXCLUDED.comments,
                    operator_id = EXCLUDED.operator_id
            ''', batch_data)
            imported_count += len(batch_data)
            conn.commit()
        
        conn.close()
        
        print(f'\n✅ Импорт завершен!')
        print(f'   Импортировано: {imported_count:,}')
        if skipped_count > 0:
            print(f'   Пропущено: {skipped_count:,}')
        if error_count > 0:
            print(f'   Ошибок: {error_count:,}')
        
        return imported_count
        
    except Exception as e:
        print(f'\n❌ Ошибка импорта: {e}')
        import traceback
        traceback.print_exc()
        return 0

def show_statistics():
    """Показать статистику по импортированным данным"""
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        
        print('\n' + '='*80)
        print('СТАТИСТИКА')
        print('='*80)
        
        cur.execute('SELECT COUNT(*) FROM fixations')
        total = cur.fetchone()[0]
        print(f'\n📊 Всего записей: {total:,}')
        
        cur.execute('''
            SELECT status_category, COUNT(*) as cnt 
            FROM fixations 
            GROUP BY status_category 
            ORDER BY cnt DESC
        ''')
        print('\n📈 По категориям:')
        for category, count in cur.fetchall():
            percentage = (count / total * 100) if total > 0 else 0
            print(f'   {category}: {count:,} ({percentage:.1f}%)')
        
        cur.execute('SELECT COUNT(DISTINCT operator_id) FROM fixations')
        ops_count = cur.fetchone()[0]
        print(f'\n👥 Операторов: {ops_count}')
        
        cur.execute('''
            SELECT o.name, COUNT(*) as cnt 
            FROM fixations f
            JOIN operators o ON f.operator_id = o.id
            GROUP BY o.name
            ORDER BY cnt DESC
            LIMIT 10
        ''')
        print('\n🏆 Топ-10 операторов:')
        for name, count in cur.fetchall():
            print(f'   {name}: {count:,}')
        
        cur.execute('SELECT MIN(call_date), MAX(call_date) FROM fixations WHERE call_date IS NOT NULL')
        min_date, max_date = cur.fetchone()
        if min_date and max_date:
            print(f'\n📅 Диапазон дат: {min_date} - {max_date}')
        
        conn.close()
        
    except Exception as e:
        print(f'Ошибка получения статистики: {e}')

def main():
    """Главная функция"""
    
    # Подключение к Google Sheets API
    print('\n[1/4] Подключение к Google Sheets API...')
    service = get_sheets_service()
    
    if not service:
        print('❌ Не удалось подключиться к Google Sheets API')
        return
    
    print('✅ Подключено к Google Sheets API')
    
    # Получение списка операторов
    print('\n[2/4] Получение списка операторов из мастер-таблицы...')
    operators = get_operator_sheets(service)
    
    if not operators:
        print('❌ Не найдены операторы')
        return
    
    print(f'✅ Найдено операторов: {len(operators)}')
    print(f'   Первые 5: {", ".join([op["name"] for op in operators[:5]])}...')
    
    # Сбор данных со всех таблиц операторов
    print(f'\n[3/4] Сбор данных из {len(operators)} таблиц...')
    all_records = []
    
    for i, operator in enumerate(tqdm(operators, desc='Чтение таблиц'), 1):
        records = read_operator_sheet(service, operator)
        if records:
            all_records.extend(records)
            print(f'  [{i}/{len(operators)}] {operator["name"]}: {len(records):,} записей')
        else:
            print(f'  [{i}/{len(operators)}] {operator["name"]}: нет данных')
        time.sleep(1.0)  # Пауза между запросами
    
    print(f'\n✅ Собрано записей: {len(all_records):,}')
    
    # Импорт в PostgreSQL
    print('\n[4/4] Импорт в PostgreSQL...')
    imported = import_to_postgresql(all_records)
    
    if imported > 0:
        show_statistics()
        
        # Обновляем контекст
        try:
            conn = psycopg2.connect(**DB_CONFIG)
            cur = conn.cursor()
            cur.execute('''
                INSERT INTO conversation_context (context_key, context_value, description)
                VALUES ('all_operators_imported', 'true', %s)
                ON CONFLICT (context_key) DO UPDATE SET
                    context_value = EXCLUDED.context_value,s)} операторов ({datetime.now().isoformat()})',))
            
            cur.execute('''
                INSERT INTO action_history (action_type, action_name, status, details)
                VALUES ('import', 'All Operators Import', 'success', %s)
            ''', (f'Импортировано {imported:,} записей от {len(operator
                INSERT INTO action_history (action_type, action_name, status, details)
                VALUES ('import', 'All Operators Import', 'success', %s)
            ''', (f'Импортировано {imported:,} записей от {len(operator_sheets)} операторов',))
            
            conn.commit()
            conn.close()
        except:
            pass
        
        print('\n' + '='*80)
        print('✅ ИМПОРТ ЗАВЕРШЕН УСПЕШНО!')
        print('='*80)
    else:
        print('\n❌ Импорт не выполнен')

if __name__ == '__main__':
    main()
