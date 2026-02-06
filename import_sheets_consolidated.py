#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
КОНСОЛИДИРОВАННЫЙ ИМПОРТ ДАННЫХ ИЗ GOOGLE SHEETS
=============================================================================
Импортирует колонки 2-12 из всех листов и объединяет в единую таблицу
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

# Прокси (закомментировано для работы в Codespaces)
# os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
# os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
# socket.setdefaulttimeout(120)

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
TOKEN_FILE = BASE_DIR / 'config' / 'token.json'
CREDENTIALS_FILE = BASE_DIR / 'config' / 'credentials.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Список ID документов для импорта
SPREADSHEET_IDS = [
    "18y_QSol_XIZiaKGdoc64-tqerxYXg1kwmO7mmxo21rQ",
    "1JG9RqC64MZbP63HCSa0avp55zXP8nEqzsiuQbbSi9Do",
    "1Ld37ljkNect6iLV0X0QztTJC7vhgPDkaRXTJGVYOtfg",
    "16AK6_7FcWeksg2KaLfuzxo6sOGxsORw9hU-lycgeGiM",
    "1h-EZltuaagu2dyFj0wf-lch2qIhXKaycyCM4MYwLa4U",
    "18WH4ocx371qXsuPIf00hvKR1Kc7u1cd2ZTOUsWS377U",
    "1iSH66bKucYRU7St8Ch_-NThYay3C-tD_p_VlLpVS7WA",
    "1toIzc0CpyQIditC9KqeMVf3II2CWAW4c96pqp1b1t6s",
    "1nrlXxHexPwEBJCyXUt6ehLSMUoC11A6tTr8B5MK1Tuw",
    "11cK2I6pQ_hrHMrbs0AhKmqLLiCbOV00R7tboG0CdrS0",
    "1-T0dKoRATQ4uWmYYli6qU7Jq9FvblD7AGW2wQZFoUB0",
    "10aFfdXjkLNlt_H0D9e4VJnxFJalGX5R9zLzAf8XpQDw",
    "1Gb0RJDqr-Z34D9dHXfPVcGJguKP46YKgPq1qgT_A4nI",
    "1-BdML7lK0fW3vrcxl8yTpgL9FpeM2NUotigf6JIN4CQ",
    "12EVYbShzbwujbqm42rkXZACrRbblo_2Ls-8XmUnID2I",
    "1YmNqdrkLeQBH5Nnq_gVdeVs3LDHArWXfHjo1BBYfpD4",
    "1mJfVK1dCSIMV4ME2lHXUyT1NXD0qaQgr_1pKqMPgCe4",
    "1MLsLwaimRSR6Gdhcm9fuatNs1B-kTgVDfUfbm6BSCjY",
    "12hSjYYlTj9DrVq4PTkI9IwTD4radEgG-Z0bYpgddXiU",
    "1_Ch5eolJHF5JQdSH7uLweBYvIZ1DomhED7pdWRdn7oI",
    "1uhvZlw1GEbMdDi5sGRJoPkd5tyvKnvob0XU0uKl99Z4",
    "1xohqcQR6vpLcmvWPNgbIPO1UCDQJmPILQ29yUpnH_A0",
    "1yDlr5nqVkoEpzPDdKRFyHTxSYJwnlyzBQ_NaJOev300",
    "1D6EIWhpH-QgjL1HvVqX54cQQl_52GJs0oAXWNQF8FKc",
    "1mP6RJtA918WUi8zq7N2RmPh4jMJfGa41UZT0mu3U4nQ",
    "1LbRFZb3830m77GKIBVipifmW6D0kCxWLP-Etku6sYts",
    "1XmQDC7hk0VYV1TQ9ETf9ZJfll5On41ZD2742H5fRmT0",
    "1j_VMVVb8CkM883y8nw2b1BHbTCCFw_43KkMZWhEK1SM",
    "1QMCAddnW5qn5OG9awAyI7Jqeo5Jzdh3mvEHuwqVBHsU",
    "1LjNHy0nsNqjeHRoAfCbGIRh_0QsLoSWSLPwRt4pik58",
    "1Ii1LlQRHtq8dqyZHCtFkCNti37fFK5ff9qbPxNHhcVw",
    "1S7oJBkx9NjsXramXxYeDq36zvN-3a--9f9KhhilfiNc",
    "1jJ8nz7lzFOgz40bN12kX1cyjkhCcquba6H9QHn8kldE",
    "1UaWDeG1pcNbGvvMrVSaRKGX6nqNdEmpQ92mueII1VRE",
    "1YuRCpm_iZkuK-eVqJ3EN5rCgEBOSgaY3qg5shaszb9A"
]

# ФИО для поиска листов
FIO_LIST = [
    "Narziyeva Gavxar Atxamjanovna",
    "Xoshimov Akromjon Axmadjon o'g'li",
    "Farxodov Xusniddin Murodjon o'g'li",
    "Rahimjonov Kamoliddin Olimjon o'g'li",
    "Xusniddinova Shaxnoza Akramovna",
    "Qosimov Firdavs Nuriddin o'g'li",
    "Abdullayev Dilmurod Xayrulla o'g'li",
    "Sirojiddinov Ismoilbek Shavkat o'g'li",
    "Mavlyanova Dilobar Rustam qizi",
    "Zokirjonova Surayyo Rustam qizi",
    "Payziyeva Shoxista Navro'zjon qizi",
    "Muxamadaliyeva Mufazzal Abduqaxxorovna",
    "Turg'unboyeva Azizaxon Shuxrat qizi",
    "Ruziyeva Dilnoza Xoshimjonovna",
    "Sobirjonova Umidaxon Rustamovna",
    "Karimova Durdona Toir qizi",
    "Xasanova Maftuna Askar qizi",
    "Sagdullayeva Moxinur Asqar qizi",
    "Mirbabayeva Shirin Kaxramonovna"
]

# Листы для импорта
SHEETS_TO_IMPORT = ["FIKSA"]
for fio in FIO_LIST:
    SHEETS_TO_IMPORT.append(f"{fio} 01.2026")

print('\n' + '='*80)
print('📥 КОНСОЛИДИРОВАННЫЙ ИМПОРТ ДАННЫХ ИЗ GOOGLE SHEETS')
print('='*80)
print(f'📋 Документов: {len(SPREADSHEET_IDS)}')
print(f'📄 Листов для поиска: {len(SHEETS_TO_IMPORT)}')
print('='*80)

def is_target_sheet(sheet_name):
    """Фильтр нужных листов: FIKSA, FIKSA(...), ФИО 01.2026"""
    if not sheet_name:
        return False
    name = str(sheet_name).strip()
    if name == 'FIKSA':
        return True
    if name.startswith('FIKSA'):
        return True
    if '01.2026' in name:
        return True
    return False

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
            print('[AUTH] Первичная авторизация (консольный режим)...')
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            auth_url, _ = flow.authorization_url(prompt='consent')
            print('\nОткройте ссылку и вставьте код авторизации:')
            print(auth_url)
            code = os.environ.get('GOOGLE_AUTH_CODE')
            if not code:
                code = input('\nКод авторизации: ').strip()
            flow.fetch_token(code=code)
            creds = flow.credentials
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return creds

# =============================================================================
# ПОЛУЧЕНИЕ ДАННЫХ
# =============================================================================

def get_sheet_data(service, spreadsheet_id, sheet_name, max_rows=10000):
    """Получить данные с указанного листа"""
    try:
        # Читаем весь лист
        range_name = f"'{sheet_name}'!A1:L{max_rows}"
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        return values
    except Exception as e:
        # Тихо пропускаем ошибки (лист может не существовать)
        return []

def get_all_sheet_names(service, spreadsheet_id):
    """Получить список всех листов в документе с размером"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        return [
            {
                'title': sheet['properties']['title'],
                'rowCount': sheet['properties'].get('gridProperties', {}).get('rowCount', 10000)
            }
            for sheet in sheets
        ]
    except:
        return []

def get_spreadsheet_title(service, spreadsheet_id):
    """Получить название документа"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return spreadsheet.get('properties', {}).get('title', 'Без названия')
    except:
        return f"ID: {spreadsheet_id[:8]}..."

def process_sheet_data_consolidated(values, sheet_name, doc_title, doc_id):
    """
    Обработать данные с листа - взять колонки 2-12 (индексы 1-11)
    ИМПОРТИРУЕМ ВСЕ СТРОКИ БЕЗ ФИЛЬТРАЦИИ
    """
    if not values or len(values) < 2:
        return []
    
    records = []
    
    # Пропускаем заголовок и берём данные со строки 2
    for row_idx, row in enumerate(values[1:], start=2):
        if not row or not any(row):  # Пропускаем только полностью пустые строки
            continue
        
        # Берём колонки 2-12 (индексы 1-11 в Python)
        record = []
        for col_idx in range(1, 12):  # Колонки 2-12 (индексы 1-11)
            if col_idx < len(row):
                record.append(row[col_idx])
            else:
                record.append('')  # Пустое значение если колонки нет
        
        # Добавляем метку импорта в колонку L
        record.append('Да')  # Импортирован

        # Добавляем мета-информацию в конец
        record.append(doc_title)  # Название документа
        record.append(sheet_name)  # Название листа
        record.append(doc_id)  # ID документа
        record.append(row_idx)  # Номер строки
        
        records.append(record)
    
    return records

# =============================================================================
# ОСНОВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    try:
        # Аутентификация
        print('\n[1/3] Подключение к Google Sheets...')
        creds = authenticate()
        service = build('sheets', 'v4', credentials=creds)
        print('✅ Подключено')
        
        # Сбор данных
        print(f'\n[2/3] Импорт данных из {len(SPREADSHEET_IDS)} документов...')
        all_records = []
        stats = {
            'processed': 0,
            'success': 0,
            'errors': 0,
            'total_records': 0,
            'sheets_found': {}
        }
        
        for idx, spreadsheet_id in enumerate(SPREADSHEET_IDS, start=1):
            try:
                # Получаем название документа
                doc_title = get_spreadsheet_title(service, spreadsheet_id)
                print(f'\n  [{idx}/{len(SPREADSHEET_IDS)}] {doc_title}')
                
                # Получаем ВСЕ листы в документе
                all_sheets = get_all_sheet_names(service, spreadsheet_id)
                
                doc_records = 0
                
                # Импортируем данные с КАЖДОГО листа в документе
                for sheet_info in all_sheets:
                    sheet_name = sheet_info['title']
                    max_rows = sheet_info['rowCount']
                    values = get_sheet_data(service, spreadsheet_id, sheet_name, max_rows=max_rows)
                    
                    if values:
                        records = process_sheet_data_consolidated(
                            values, sheet_name, doc_title, spreadsheet_id
                        )
                        all_records.extend(records)
                        doc_records += len(records)
                        
                        # Статистика по листам
                        if sheet_name not in stats['sheets_found']:
                            stats['sheets_found'][sheet_name] = 0
                        stats['sheets_found'][sheet_name] += len(records)
                        
                        if len(records) > 0:
                            print(f'    ✓ {sheet_name}: {len(records)} записей')
                
                stats['processed'] += 1
                stats['success'] += 1
                stats['total_records'] += doc_records
                
                if doc_records > 0:
                    print(f'    Итого: {doc_records} записей')
                
            except Exception as e:
                print(f'    ❌ Ошибка: {str(e)}')
                stats['errors'] += 1
        
        # Сохранение
        print(f'\n[3/3] Сохранение данных...')
        
        if all_records:
            # Создаём DataFrame с правильными колонками
            columns = [
                'Колонка_2', 'Колонка_3', 'Колонка_4', 'Колонка_5', 'Колонка_6',
                'Колонка_7', 'Колонка_8', 'Колонка_9', 'Колонка_10', 'Колонка_11', 'Колонка_12',
                'Импортирован',
                'Документ', 'Лист', 'ID_Документа', 'Номер_Строки'
            ]
            
            df = pd.DataFrame(all_records, columns=columns)
            
            # Сохраняем в CSV
            timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            csv_file = BASE_DIR / 'data' / f'КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_{timestamp}.csv'
            df.to_csv(csv_file, index=False, encoding='utf-8-sig')
            
            print(f'\n✅ Сохранено в CSV: {csv_file.name}')
            print(f'   Всего строк: {len(df)}')
            print(f'   Колонок: {len(columns)}')

            # Сохраняем в SQLite
            try:
                DB_PATH.parent.mkdir(parents=True, exist_ok=True)
                conn = sqlite3.connect(DB_PATH)
                cur = conn.cursor()
                cur.execute('DROP TABLE IF EXISTS sheets_data')
                cur.execute('''
                    CREATE TABLE sheets_data (
                        Колонка_2 TEXT, Колонка_3 TEXT, Колонка_4 TEXT, Колонка_5 TEXT, Колонка_6 TEXT,
                        Колонка_7 TEXT, Колонка_8 TEXT, Колонка_9 TEXT, Колонка_10 TEXT, Колонка_11 TEXT, Колонка_12 TEXT,
                        Импортирован TEXT,
                        Документ TEXT, Лист TEXT, ID_Документа TEXT, Номер_Строки INTEGER
                    )
                ''')
                conn.commit()

                batch_size = 5000
                rows = df.values.tolist()
                for i in range(0, len(rows), batch_size):
                    cur.executemany(
                        'INSERT INTO sheets_data VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                        rows[i:i+batch_size]
                    )
                    conn.commit()
                conn.close()
                print(f'✅ Сохранено в SQLite: {DB_PATH.name}')
            except Exception as e:
                print(f'⚠️  Ошибка при сохранении в SQLite: {e}')
        
        # Итоговая статистика
        print('\n' + '='*80)
        print('✅ ИМПОРТ ЗАВЕРШЁН')
        print('='*80)
        print(f'\n📊 Статистика:')
        print(f'  Обработано документов: {stats["processed"]}/{len(SPREADSHEET_IDS)}')
        print(f'  Успешно: {stats["success"]}')
        print(f'  Ошибок: {stats["errors"]}')
        print(f'  Всего записей импортировано: {stats["total_records"]}')
        
        # Статистика по листам
        if stats['sheets_found']:
            print(f'\n📄 СТАТИСТИКА ПО ЛИСТАМ:')
            for sheet_name, count in sorted(stats['sheets_found'].items(), key=lambda x: x[1], reverse=True):
                if count > 0:
                    print(f'  {sheet_name[:50]:<50} - {count:>6} записей')
        
        print('\n' + '='*80)
        
    except Exception as e:
        print(f'\n❌ КРИТИЧЕСКАЯ ОШИБКА: {str(e)}')
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
