#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт данных из Google Docs в PostgreSQL
Читает JSON Lines из Google Docs и записывает в БД
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

# Google API конфигурация
CREDENTIALS_FILE = CONFIG_DIR / 'credentials.json'
TOKEN_FILE = CONFIG_DIR / 'token.json'
SCOPES = [
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

print('\n' + '='*80)
print('ИМПОРТ ИЗ GOOGLE DOCS В POSTGRESQL')
print('='*80)

def get_google_service(service_name, version):
    """Получение сервиса Google API с OAuth2"""
    if not CREDENTIALS_FILE.exists():
        print(f'\nОшибка: Файл credentials.json не найден в {CONFIG_DIR}')
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
            print('✅ Токен сохранен')
        
        service = build(service_name, version, credentials=creds)
        return service
    except Exception as e:
        print(f"Ошибка подключения к Google API: {e}")
        import traceback
        traceback.print_exc()
        return None

def read_docs_data(docs_service, document_id):
    """Читает данные из Google Docs (JSON Lines format)"""
    print(f'\n📄 Чтение документа {document_id}...')
    
    try:
        document = docs_service.documents().get(documentId=document_id).execute()
        content = document.get('body').get('content')
        
        records = []
        line_count = 0
        error_count = 0
        
        # Проходим по всем параграфам
        for element in content:
            if 'paragraph' in element:
                paragraph = element.get('paragraph')
                elements = paragraph.get('elements', [])
                
                for elem in elements:
                    if 'textRun' in elem:
                        text = elem.get('textRun').get('content').strip()
                        
                        # Пропускаем пустые строки и не-JSON
                        if not text or not text.startswith('{'):
                            continue
                        
                        # Парсим JSON
                        try:
                            record = json.loads(text)
                            records.append(record)
                            line_count += 1
                        except json.JSONDecodeError as e:
                            error_count += 1
                            if error_count <= 5:  # Показываем первые 5 ошибок
                                print(f'Ошибка парсинга JSON: {str(e)[:100]}')
        
        print(f'✅ Прочитано {line_count:,} записей из документа')
        if error_count > 0:
            print(f'⚠️  Ошибок парсинга: {error_count}')
        
        return records
    except HttpError as e:
        print(f'❌ Ошибка доступа к документу: {e}')
        print('Проверьте:')
        print('1. ID документа правильный')
        print('2. У вас есть доступ к документу')
        print('3. Google Docs API включен в проекте')
        return []
    except Exception as e:
        print(f'❌ Ошибка чтения документа: {e}')
        import traceback
        traceback.print_exc()
        return []

def import_to_postgresql(records):
    """Импорт записей в PostgreSQL"""
    if not records:
        print('Нет данных для импорта')
        return 0
    
    print(f'\n💾 Импорт {len(records):,} записей в PostgreSQL...')
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        
        # Получаем или создаем операторов
        operators_cache = {}
        
        # Загружаем существующих операторов
        cur.execute('SELECT name, id FROM operators')
        for name, op_id in cur.fetchall():
            operators_cache[name] = op_id
        
        imported_count = 0
        skipped_count = 0
        error_count = 0
        batch_data = []
        
        for record in tqdm(records, desc='Обработка записей'):
            try:
                operator_name = record.get('operator')
                card_number = record.get('card_number')
                call_date = record.get('call_date')
                phone_number = record.get('phone_number')
                status = record.get('status')
                service_name = record.get('service')
                comments = record.get('comments', '')
                
                # Пропускаем записи без ключевых полей
                if not operator_name or not card_number:
                    skipped_count += 1
                    continue
                
                # Получаем или создаем оператора
                if operator_name not in operators_cache:
                    cur.execute(
                        'INSERT INTO operators (name) VALUES (%s) ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name RETURNING id',
                        (operator_name,)
                    )
                    operators_cache[operator_name] = cur.fetchone()[0]
                
                operator_id = operators_cache[operator_name]
                
                # Парсим дату
                parsed_date = None
                if call_date:
                    try:
                        # Пробуем разные форматы дат
                        for date_format in ['%d.%m.%Y %H:%M', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']:
                            try:
                                parsed_date = datetime.strptime(call_date, date_format)
                                break
                            except:
                                continue
                    except:
                        pass
                
                batch_data.append((
                    card_number,
                    phone_number,
                    parsed_date,
                    status,
                    service_name,
                    comments,
                    operator_id
                ))
                
                # Коммитим батчами по 5000 записей
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
                if error_count <= 5:  # Показываем первые 5 ошибок
                    print(f'\n⚠️  Ошибка обработки записи: {e}')
                continue
        
        # Импортируем оставшиеся записи
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
        print(f'   Пропущено: {skipped_count:,}')
        if error_count > 0:
            print(f'   Ошибок: {error_count:,}')
        
        return imported_count
        
    except Exception as e:
        print(f'\n❌ Ошибка импорта: {e}')
        import traceback
        traceback.print_exc()
        return 0

def show_statistics(conn):
    """Показать статистику по импортированным данным"""
    try:
        cur = conn.cursor()
        
        print('\n' + '='*80)
        print('СТАТИСТИКА')
        print('='*80)
        
        # Общее количество  
        cur.execute('SELECT COUNT(*) FROM fixations')
        total = cur.fetchone()[0]
        print(f'\n📊 Всего записей: {total:,}')
        
        # По категориям
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
        
        # Топ-5 операторов
        cur.execute('''
            SELECT o.name, COUNT(*) as cnt 
            FROM fixations f
            JOIN operators o ON f.operator_id = o.id
            GROUP BY o.name
            ORDER BY cnt DESC
            LIMIT 5
        ''')
        print('\n🏆 Топ-5 операторов:')
        for name, count in cur.fetchall():
            print(f'   {name}: {count:,}')
        
        conn.close()
        
    except Exception as e:
        print(f'Ошибка получения статистики: {e}')

def main():
    """Главная функция"""
    
    # Запрашиваем ID документа
    print('\n📝 Введите ID документа Google Docs')
    print('(ID находится в URL: https://docs.google.com/document/d/[THIS_IS_ID]/edit)')
    print('\nОставьте пустым если хотите использовать тестовый документ:')
    
    docs_id = input('ID документа: ').strip()
    
    if not docs_id:
        print('❌ ID документа обязателен')
        return
    
    # Подключение к Google Docs API
    print('\n[1/3] Подключение к Google Docs API...')
    docs_service = get_google_service('docs', 'v1')
    
    if not docs_service:
        print('❌ Не удалось подключиться к Google Docs API')
        return
    
    print('✅ Подключено к Google Docs API')
    
    # Чтение данных из документа
    print('\n[2/3] Чтение данных из Google Docs...')
    records = read_docs_data(docs_service, docs_id)
    
    if not records:
        print('❌ Не удалось прочитать данные из документа')
        return
    
    # Импорт в PostgreSQL
    print('\n[3/3] Импорт в PostgreSQL...')
    imported = import_to_postgresql(records)
    
    if imported > 0:
        # Показываем статистику
        conn = psycopg2.connect(**DB_CONFIG)
        show_statistics(conn)
        
        print('\n' + '='*80)
        print('✅ ИМПОРТ ЗАВЕРШЕН УСПЕШНО!')
        print('='*80)
    else:
        print('\n❌ Импорт не выполнен')

if __name__ == '__main__':
    main()
