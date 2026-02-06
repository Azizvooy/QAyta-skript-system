#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Импорт через Service Account (без OAuth)
"""

import os
import time
import pandas as pd
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime

BASE_DIR = Path(__file__).parent
SERVICE_ACCOUNT_FILE = BASE_DIR / 'config' / 'service_account.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

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

def authenticate():
    """Аутентификация через Service Account"""
    if not SERVICE_ACCOUNT_FILE.exists():
        print(f"\n❌ Файл {SERVICE_ACCOUNT_FILE} не найден!")
        print("\nИнструкция по созданию Service Account:")
        print("1. Открой: https://console.cloud.google.com/")
        print("2. APIs & Services → Credentials")
        print("3. Create Credentials → Service Account")
        print("4. Заполни имя, нажми Create and Continue")
        print("5. Нажми Continue, затем Done")
        print("6. Кликни на созданный Service Account")
        print("7. Keys → Add Key → Create New Key → JSON")
        print("8. Сохрани файл как config/service_account.json")
        print("9. ВАЖНО: Дай доступ к каждому Google Sheets документу:")
        print("   - Открой каждый документ")
        print("   - Нажми Share")
        print("   - Добавь email из service_account.json (поле 'client_email')")
        print("   - Дай права 'Viewer'\n")
        return None
    
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return credentials

def get_all_sheets(service, spreadsheet_id):
    """Получить все листы документа"""
    try:
        metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = metadata.get('sheets', [])
        return [{'title': s['properties']['title'], 
                 'rowCount': s['properties']['gridProperties'].get('rowCount', 0)}
                for s in sheets]
    except Exception as e:
        print(f"❌ Ошибка получения листов: {e}")
        return []

def request_with_retry(func, max_retries=6, base_delay=1.5):
    for attempt in range(1, max_retries + 1):
        try:
            return func()
        except HttpError as e:
            status = e.resp.status if e.resp else None
            if status in (429, 500, 502, 503, 504):
                sleep_time = base_delay * attempt
                print(f"  ⏳ Лимит/ошибка {status}. Жду {sleep_time:.1f}с и повторяю...", flush=True)
                time.sleep(sleep_time)
                continue
            raise
    return None

def should_import_sheet(title: str) -> bool:
    name = title.lower()
    is_fiksa = "fiksa" in name and "state" not in name
    is_jan_2026 = "01.2026" in name
    return is_fiksa or is_jan_2026

def get_sheet_data(service, spreadsheet_id, sheet_name, max_rows=None):
    """Получить данные листа (колонки B-L)"""
    try:
        range_name = f"'{sheet_name}'!B:L"
        if max_rows:
            range_name = f"'{sheet_name}'!B2:L{max_rows}"
        
        def do_get():
            return service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()

        result = request_with_retry(do_get)
        if result is None:
            return pd.DataFrame()
        
        values = result.get('values', [])
        if not values:
            return pd.DataFrame()
        
        df = pd.DataFrame(values)
        df.columns = [f'Колонка_{i}' for i in range(2, 2 + len(df.columns))]
        
        return df
    except Exception as e:
        print(f"  ⚠️  Ошибка чтения '{sheet_name}': {e}")
        return pd.DataFrame()

def main():
    print("="*80)
    print("📥 ИМПОРТ ИЗ GOOGLE SHEETS (Service Account)")
    print("="*80)
    
    # Авторизация
    print("\n[1/3] Подключение...")
    creds = authenticate()
    if not creds:
        return
    
    service = build('sheets', 'v4', credentials=creds)
    print("✅ Подключено!")
    
    # Импорт
    print(f"\n[2/3] Импорт из {len(SPREADSHEET_IDS)} документов...")
    print("Фильтр: листы FIKSA и ФИО 01.2026")
    all_data = []
    
    for idx, sheet_id in enumerate(SPREADSHEET_IDS, 1):
        print(f"\n[{idx}/{len(SPREADSHEET_IDS)}] Документ {sheet_id[:8]}...")
        
        sheets = get_all_sheets(service, sheet_id)
        if not sheets:
            continue
        
        for sheet in sheets:
            name = sheet['title']
            rows = sheet['rowCount']
            if not should_import_sheet(name):
                continue

            print(f"  📄 {name} ({rows:,} строк)...", end=' ')
            
            df = get_sheet_data(service, sheet_id, name, max_rows=rows)
            if not df.empty:
                df['Источник_документ'] = sheet_id
                df['Источник_лист'] = name
                all_data.append(df)
                print(f"✅ {len(df):,}")
            else:
                print("⚠️ Пусто")
            time.sleep(0.4)
    
    # Сохранение
    if not all_data:
        print("\n❌ Нет данных для сохранения!")
        return
    
    print(f"\n[3/3] Объединение {len(all_data)} листов...")
    df_final = pd.concat(all_data, ignore_index=True)

    if 'Импортирован' not in df_final.columns:
        if 'Колонка_12' in df_final.columns:
            insert_at = df_final.columns.get_loc('Колонка_12') + 1
            df_final.insert(insert_at, 'Импортирован', 'Да')
        else:
            df_final['Импортирован'] = 'Да'
    
    # Фильтр по датам (по умолчанию выключен)
    if 'Колонка_4' in df_final.columns and str(os.environ.get('APPLY_DATE_FILTER', '')).strip() == '1':
        print("Фильтрация по датам 04-31.01.2026...")
        df_final['Дата_temp'] = pd.to_datetime(df_final['Колонка_4'], 
                                               format='%d.%m.%Y', 
                                               errors='coerce')
        df_final = df_final[
            (df_final['Дата_temp'] >= '2026-01-04') & 
            (df_final['Дата_temp'] <= '2026-01-31')
        ]
        df_final = df_final.drop(columns=['Дата_temp'])
    
    # Сохранение
    output_dir = BASE_DIR / 'data'
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_path = output_dir / f'КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_{timestamp}.csv'
    
    df_final.to_csv(csv_path, index=False, encoding='utf-8-sig')
    
    print(f"\n{'='*80}")
    print(f"✅ ГОТОВО! Импортировано {len(df_final):,} записей")
    print(f"📁 Файл: {csv_path.name}")
    print(f"{'='*80}\n")

if __name__ == '__main__':
    main()
