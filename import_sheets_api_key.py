#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ИМПОРТ ИЗ GOOGLE SHEETS ЧЕРЕЗ API KEY
Работает с публичными документами БЕЗ авторизации
Импортирует ВСЕ строки без ограничений
"""

import os
import pandas as pd
from pathlib import Path
from datetime import datetime
import requests
import time

BASE_DIR = Path(__file__).parent

# API KEY из переменной окружения или файла
API_KEY_FILE = BASE_DIR / 'config' / 'api_key.txt'

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

def get_api_key():
    """Получить API ключ"""
    # Из переменной окружения
    api_key = os.environ.get('GOOGLE_API_KEY')
    if api_key:
        return api_key
    
    # Из файла
    if API_KEY_FILE.exists():
        return API_KEY_FILE.read_text().strip()
    
    print("\n" + "="*80)
    print("📝 НАСТРОЙКА API KEY (ОДИН РАЗ)")
    print("="*80)
    print("\n1. Открой: https://console.cloud.google.com/")
    print("2. APIs & Services → Credentials")
    print("3. Create Credentials → API Key")
    print("4. Скопируй ключ")
    print("5. Создай файл config/api_key.txt и вставь туда ключ")
    print("\nИЛИ передай через переменную: export GOOGLE_API_KEY='твой_ключ'\n")
    return None

def request_with_retry(url, params, max_retries=6, base_delay=1.5):
    """Запрос с повторными попытками при 429/5xx"""
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, params=params, timeout=60)
            if response.status_code in (429, 500, 502, 503, 504):
                raise requests.exceptions.HTTPError(response=response)
            return response
        except requests.exceptions.HTTPError as e:
            status = getattr(e.response, 'status_code', None)
            if status in (429, 500, 502, 503, 504):
                sleep_time = base_delay * attempt
                print(f"  ⏳ Лимит/ошибка {status}. Жду {sleep_time:.1f}с и повторяю...", flush=True)
                time.sleep(sleep_time)
                continue
            raise
    return None

def get_all_sheets(spreadsheet_id, api_key):
    """Получить все листы документа"""
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
    params = {'key': api_key}
    
    try:
        response = request_with_retry(url, params)
        if response is None:
            print("  ❌ Не удалось получить список листов после повторов")
            return []
        response.raise_for_status()
        data = response.json()
        
        sheets = []
        for sheet in data.get('sheets', []):
            props = sheet['properties']
            sheets.append({
                'title': props['title'],
                'sheetId': props['sheetId'],
                'rowCount': props['gridProperties'].get('rowCount', 0),
                'colCount': props['gridProperties'].get('columnCount', 0)
            })
        return sheets
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 403:
            print(f"  ❌ Доступ запрещен - документ должен быть публичным!")
            print(f"     Открой: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
            print(f"     Нажми Share → Get link → Anyone with the link (Viewer)")
        else:
            print(f"  ❌ HTTP ошибка: {e}")
        return []
    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        return []

def get_sheet_data(spreadsheet_id, sheet_name, api_key):
    """Получить ВСЕ данные листа (колонки B-L) БЕЗ ограничений"""
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{sheet_name}!B:L"
    params = {
        'key': api_key,
        'valueRenderOption': 'UNFORMATTED_VALUE',
        'dateTimeRenderOption': 'FORMATTED_STRING'
    }
    
    try:
        response = request_with_retry(url, params)
        if response is None:
            return pd.DataFrame()
        response.raise_for_status()
        data = response.json()
        
        values = data.get('values', [])
        if not values:
            return pd.DataFrame()
        
        # Создаем DataFrame
        df = pd.DataFrame(values)
        
        # Добавляем недостающие колонки если нужно
        expected_cols = 11  # B-L = 11 колонок
        while len(df.columns) < expected_cols:
            df[len(df.columns)] = None
        
        # Переименовываем
        df.columns = [f'Колонка_{i}' for i in range(2, 2 + len(df.columns))]
        
        return df
        
    except Exception as e:
        print(f"  ⚠️  Ошибка чтения '{sheet_name}': {e}")
        return pd.DataFrame()

def main():
    print("="*80)
    print("📥 ИМПОРТ ИЗ GOOGLE SHEETS (API KEY - БЕЗ ОГРАНИЧЕНИЙ)")
    print("="*80)
    
    # Получаем API ключ
    api_key = get_api_key()
    if not api_key:
        return
    
    print(f"\n✅ API Key найден")
    print(f"📋 Документов: {len(SPREADSHEET_IDS)}")
    print(f"⚡ Импорт БЕЗ ограничений по строкам\n")
    
    all_data = []
    total_sheets = 0
    
    for idx, sheet_id in enumerate(SPREADSHEET_IDS, 1):
        print(f"[{idx}/{len(SPREADSHEET_IDS)}] Документ {sheet_id[:8]}...")
        
        # Получаем список листов
        sheets = get_all_sheets(sheet_id, api_key)
        if not sheets:
            continue
        
        # Импортируем каждый лист
        for sheet in sheets:
            name = sheet['title']
            rows = sheet['rowCount']
            
            print(f"  📄 {name} ({rows:,} строк)...", end=' ', flush=True)
            
            df = get_sheet_data(sheet_id, name, api_key)
            
            if not df.empty:
                df['Источник_документ'] = sheet_id
                df['Источник_лист'] = name
                all_data.append(df)
                total_sheets += 1
                print(f"✅ {len(df):,}")
            else:
                print("⚠️ Пусто")
            
            # Пауза чтобы не превысить лимит API
            time.sleep(0.4)
    
    if not all_data:
        print("\n❌ Нет данных для сохранения!")
        return
    
    print(f"\n{'='*80}")
    print(f"Объединение {total_sheets} листов...")
    df_final = pd.concat(all_data, ignore_index=True)

    if 'Импортирован' not in df_final.columns:
        if 'Колонка_12' in df_final.columns:
            insert_at = df_final.columns.get_loc('Колонка_12') + 1
            df_final.insert(insert_at, 'Импортирован', 'Да')
        else:
            df_final['Импортирован'] = 'Да'
    
    print(f"Всего записей: {len(df_final):,}")
    
    # Фильтрация по датам января 2026 (04-31)
    if 'Колонка_4' in df_final.columns:
        print("\nФильтрация по датам 04-31.01.2026...")
        
        # Преобразуем даты
        df_final['Дата_temp'] = pd.to_datetime(
            df_final['Колонка_4'], 
            format='%d.%m.%Y',
            errors='coerce'
        )
        
        before = len(df_final)
        df_final = df_final[
            (df_final['Дата_temp'] >= '2026-01-04') & 
            (df_final['Дата_temp'] <= '2026-01-31')
        ]
        after = len(df_final)
        
        df_final = df_final.drop(columns=['Дата_temp'])
        print(f"До фильтрации: {before:,}")
        print(f"После фильтрации: {after:,}")
        print(f"Отфильтровано: {before - after:,}")
    
    # Сохранение
    output_dir = BASE_DIR / 'data'
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_path = output_dir / f'КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_{timestamp}.csv'
    
    df_final.to_csv(csv_path, index=False, encoding='utf-8-sig')
    
    print(f"\n{'='*80}")
    print(f"✅ ГОТОВО!")
    print(f"📊 Импортировано: {len(df_final):,} записей")
    print(f"📁 Файл: {csv_path.name}")
    print(f"{'='*80}\n")

if __name__ == '__main__':
    main()
