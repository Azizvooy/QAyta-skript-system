#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ИМПОРТ ДАННЫХ ИЗ МНОЖЕСТВА GOOGLE SHEETS
=============================================================================
Импортирует данные из листов "ФИО 01.2026" и "FIKSA" из всех указанных документов
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
SHEETS_TO_IMPORT = ["FIKSA"]  # Базовые листы
# Добавляем листы с ФИО + " 01.2026"
for fio in FIO_LIST:
    SHEETS_TO_IMPORT.append(f"{fio} 01.2026")

print('\n' + '='*80)
print('📥 ИМПОРТ ДАННЫХ ИЗ МНОЖЕСТВА GOOGLE SHEETS')
print('='*80)
print(f'📋 Документов для обработки: {len(SPREADSHEET_IDS)}')
print(f'📄 Листов для поиска: {len(SHEETS_TO_IMPORT)} ({len(FIO_LIST)} ФИО + FIKSA)')
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

def get_sheet_data(service, spreadsheet_id, sheet_name):
    """Получить данные с указанного листа"""
    try:
        # Читаем весь лист
        range_name = f"'{sheet_name}'!A1:Z10000"
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        return values
    except Exception as e:
        print(f"    ⚠️  Ошибка чтения листа '{sheet_name}': {str(e)}")
        return []

def get_spreadsheet_title(service, spreadsheet_id):
    """Получить название документа"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return spreadsheet.get('properties', {}).get('title', 'Без названия')
    except:
        return f"ID: {spreadsheet_id[:8]}..."

def process_sheet_data(values, sheet_name, doc_title):
    """Обработать данные с листа"""
    if not values or len(values) < 2:
        return []
    
    records = []
    headers = values[0] if values else []
    
    for idx, row in enumerate(values[1:], start=2):
        if not row or not any(row):  # Пропускаем пустые строки
            continue
        
        # Создаём запись с основными данными
        record = {
            'источник': doc_title,
            'лист': sheet_name,
            'строка': idx,
            'дата_импорта': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Добавляем данные из столбцов
        for col_idx, header in enumerate(headers):
            if col_idx < len(row):
                record[header] = row[col_idx]
        
        records.append(record)
    
    return records

# =============================================================================
# СОХРАНЕНИЕ В БД
# =============================================================================

def save_to_database(all_records):
    """Сохранить все записи в базу данных"""
    if not all_records:
        print('\n⚠️  Нет данных для сохранения')
        return
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Создаём таблицу для импортированных данных если не существует
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS imported_sheets_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            источник TEXT,
            лист TEXT,
            строка INTEGER,
            дата_импорта TEXT,
            данные TEXT
        )
    ''')
    
    # Сохраняем записи
    added = 0
    for record in all_records:
        try:
            # Преобразуем запись в JSON-строку
            import json
            data_json = json.dumps(record, ensure_ascii=False)
            
            cursor.execute('''
                INSERT INTO imported_sheets_data 
                (источник, лист, строка, дата_импорта, данные)
                VALUES (?, ?, ?, ?, ?)
            ''', (
                record.get('источник', ''),
                record.get('лист', ''),
                record.get('строка', 0),
                record.get('дата_импорта', ''),
                data_json
            ))
            added += 1
        except Exception as e:
            print(f"    ⚠️  Ошибка при сохранении записи: {str(e)}")
    
    conn.commit()
    conn.close()
    
    print(f'\n✅ Сохранено в базу данных: {added} записей')

# =============================================================================
# ЭКСПОРТ В CSV
# =============================================================================

def export_to_csv(all_records):
    """Экспорт всех данных в CSV файл"""
    if not all_records:
        return
    
    # Создаём DataFrame
    df = pd.DataFrame(all_records)
    
    # Сохраняем в CSV
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    csv_file = BASE_DIR / 'data' / f'ИМПОРТ_МНОЖЕСТВЕННЫХ_SHEETS_{timestamp}.csv'
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    
    print(f'\n📄 Экспортировано в CSV: {csv_file.name}')
    print(f'   Всего строк: {len(df)}')

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
            'total_records': 0
        }
        
        for idx, spreadsheet_id in enumerate(SPREADSHEET_IDS, start=1):
            try:
                # Получаем название документа
                doc_title = get_spreadsheet_title(service, spreadsheet_id)
                print(f'\n  [{idx}/{len(SPREADSHEET_IDS)}] {doc_title}')
                
                doc_records = 0
                
                # Импортируем данные с каждого листа
                for sheet_name in SHEETS_TO_IMPORT:
                    values = get_sheet_data(service, spreadsheet_id, sheet_name)
                    
                    if values:
                        records = process_sheet_data(values, sheet_name, doc_title)
                        all_records.extend(records)
                        doc_records += len(records)
                        print(f'    ✓ {sheet_name}: {len(records)} записей')
                    else:
                        print(f'    - {sheet_name}: нет данных')
                
                stats['processed'] += 1
                stats['success'] += 1
                stats['total_records'] += doc_records
                print(f'    Итого из документа: {doc_records} записей')
                
            except Exception as e:
                print(f'    ❌ Ошибка: {str(e)}')
                stats['errors'] += 1
        
        # Сохранение
        print(f'\n[3/3] Сохранение данных...')
        if all_records:
            save_to_database(all_records)
            export_to_csv(all_records)
        
        # Итоговая статистика
        print('\n' + '='*80)
        print('✅ ИМПОРТ ЗАВЕРШЁН')
        print('='*80)
        print(f'\n📊 Статистика:')
        print(f'  Обработано документов: {stats["processed"]}/{len(SPREADSHEET_IDS)}')
        print(f'  Успешно: {stats["success"]}')
        print(f'  Ошибок: {stats["errors"]}')
        print(f'  Всего записей импортировано: {stats["total_records"]}')
        
        # ТОП-10 по количеству записей
        if all_records:
            from collections import Counter
            sources = [r['источник'] for r in all_records]
            top_sources = Counter(sources).most_common(10)
            
            print(f'\n🏆 ТОП-10 ДОКУМЕНТОВ ПО КОЛИЧЕСТВУ ЗАПИСЕЙ:')
            for source, count in top_sources:
                print(f'  {source[:50]:<50} - {count} записей')
        
        print('\n' + '='*80)
        
    except Exception as e:
        print(f'\n❌ КРИТИЧЕСКАЯ ОШИБКА: {str(e)}')
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
