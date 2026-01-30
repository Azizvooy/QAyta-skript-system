"""
=============================================================================
СБОРЩИК ДАННЫХ В EXCEL - ВСЕ СТРОКИ ИЗ ВСЕХ АРХИВНЫХ ЛИСТОВ
=============================================================================
Собирает все данные из архивных листов всех операторов в один Excel файл
=============================================================================
"""

import json
import os
from datetime import datetime
from typing import List, Dict
import time

# Настройка прокси для корпоративной сети
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import socket

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("❌ Установите pandas: pip install pandas openpyxl")
    exit(1)

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

# =============================================================================
# НАСТРОЙКИ
# =============================================================================

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# ID таблицы со списком операторов
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

# Имя листа с настройками
SETTINGS_SHEET_NAME = "Настройки"

# Служебные листы (НЕ архивы)
SKIP_SHEETS = ["Статистика", "Предыдущий месяц", "Сводка по дням", "Настройки"]

# Файлы
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'
OUTPUT_EXCEL = 'ALL_DATA.xlsx'

# Настройки
socket.setdefaulttimeout(120)
BATCH_SIZE = 1000  # Читать по 1000 строк за раз

# =============================================================================
# АУТЕНТИФИКАЦИЯ
# =============================================================================

def authenticate():
    """Аутентификация в Google API"""
    creds = None
    
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Обновление токена...")
            # Создаем requests session с прокси
            import requests
            session = requests.Session()
            session.proxies = {
                'http': 'http://10.145.62.76:3128',
                'https': 'http://10.145.62.76:3128',
            }
            from google.auth.transport.requests import Request as GoogleRequest
            request = GoogleRequest(session=session)
            creds.refresh(request)
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print("❌ Файл credentials.json не найден!")
                return None
            
            print("🔐 Авторизация (откроется браузер)...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print("✅ Токен сохранен")
    
    return creds

# =============================================================================
# ЧТЕНИЕ ДАННЫХ
# =============================================================================

def get_operator_list(service) -> List[Dict[str, str]]:
    """Получить список ВСЕХ операторов (включая уволенных, с пустыми ФИО)"""
    print(f"\n📋 Чтение списка операторов...")
    
    result = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SPREADSHEET_ID,
        range=f"{SETTINGS_SHEET_NAME}!A2:C100"  # Читаем до 100 строки
    ).execute()
    
    values = result.get('values', [])
    operators = []
    
    for idx, row in enumerate(values, start=2):
        # Пропускаем полностью пустые строки
        if len(row) == 0:
            continue
            
        # Главное - наличие ID таблицы (даже если ФИО пустое)
        spreadsheet_id = row[1].strip() if len(row) > 1 and row[1] else ""
        
        # Читаем ВСЕ строки где есть ID таблицы
        if spreadsheet_id and spreadsheet_id != "ID таблицы":
            name = row[0].strip() if len(row) > 0 and row[0] else f"Оператор {idx}"
            status = row[2].strip() if len(row) > 2 and row[2] else "не указан"
            
            operators.append({
                'name': name if name else f"Оператор {idx}",
                'spreadsheet_id': spreadsheet_id,
                'status': status
            })
    
    print(f"✅ Найдено операторов: {len(operators)} (все строки с ID таблицы)")
    return operators

def get_sheet_list(service, spreadsheet_id) -> List[str]:
    """Получить список архивных листов (только с ФИО или датами)"""
    try:
        metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = metadata.get('sheets', [])
        
        # Фильтруем: оставляем только листы с ФИО или датами
        sheet_names = []
        for sheet in sheets:
            title = sheet['properties']['title']
            title_lower = title.lower()
            
            # Пропускаем служебные листы
            if title in SKIP_SHEETS:
                continue
            
            # Пропускаем листы с определенными словами
            skip_words = ['setting', 'аризалар', 'аризалары', 'настройки', 'сводка', 'статистика']
            if any(word in title_lower for word in skip_words):
                continue
            
            # Добавляем лист (если в названии есть буквы - предполагаем что это ФИО или дата)
            if title.strip():
                sheet_names.append(title)
        
        return sheet_names
    except Exception as e:
        print(f"  ❌ Ошибка получения листов: {e}")
        return []

def read_sheet_data(service, spreadsheet_id, sheet_name) -> List[List]:
    """Читать все данные из листа"""
    try:
        # Читаем большой диапазон
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!B2:I20000",  # Читаем до 20k строк
            valueRenderOption='FORMATTED_VALUE'
        ).execute()
        
        return result.get('values', [])
    except Exception as e:
        print(f"    ⚠️ Ошибка чтения {sheet_name}: {e}")
        return []

# =============================================================================
# ОБРАБОТКА
# =============================================================================

def collect_all_data(service, operators):
    """Собрать все данные от всех операторов"""
    all_rows = []
    
    print(f"\n🚀 Начало сбора данных...\n")
    
    iterator = tqdm(operators, desc="Операторы") if HAS_TQDM else operators
    
    for operator in iterator:
        operator_name = operator['name']
        spreadsheet_id = operator['spreadsheet_id']
        
        if not HAS_TQDM:
            print(f"▶ {operator_name}")
        
        # Получаем список листов
        sheets = get_sheet_list(service, spreadsheet_id)
        
        if not HAS_TQDM:
            print(f"  Найдено архивных листов: {len(sheets)}")
        
        # Читаем каждый лист
        for sheet_name in sheets:
            if not HAS_TQDM:
                print(f"    📄 {sheet_name}...", end=" ")
            
            rows = read_sheet_data(service, spreadsheet_id, sheet_name)
            
            # Добавляем информацию об операторе и листе
            for row in rows:
                if len(row) > 0 and row[0]:  # Если есть номер карты
                    # Добавляем: Оператор | Лист | Данные из B-I
                    extended_row = [operator_name, sheet_name] + row
                    all_rows.append(extended_row)
            
            if not HAS_TQDM:
                print(f"✓ {len(rows)} строк")
            
            time.sleep(0.1)  # Пауза между запросами
    
    return all_rows

# =============================================================================
# СОХРАНЕНИЕ
# =============================================================================

def save_to_excel(data, filename):
    """Сохранить данные в CSV и Excel (если возможно)"""
    print(f"\n💾 Сохранение данных...")
    
    # Создаем DataFrame
    columns = [
        'Оператор',
        'Архивный лист',
        'Номер карты',
        'Номер телефона',
        'Дата открытия карты',
        'Статус',
        'Служба',
        'Комментарий',
        'Оператор фиксировавший',
        'Дата фиксации'
    ]
    
    df = pd.DataFrame(data, columns=columns)
    
    print(f"   Всего строк: {len(df):,}")
    
    # Сохраняем в CSV (без ограничений по размеру)
    csv_filename = filename.replace('.xlsx', '.csv')
    df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
    print(f"✅ CSV сохранен: {csv_filename}")
    
    # Пытаемся сохранить в Excel, если < 1 млн строк
    if len(df) < 1000000:
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Все данные', index=False)
            print(f"✅ Excel сохранен: {filename}")
        except Exception as e:
            print(f"⚠️ Excel не создан (слишком большой файл)")
    else:
        print(f"⚠️ Excel не создан: {len(df):,} строк > лимита 1,048,576")
        print(f"   Используйте CSV файл: {csv_filename}")
    
    # Статистика
    print(f"\n📈 СТАТИСТИКА:")
    print(f"   Всего строк: {len(df):,}")
    print(f"   Операторов: {df['Оператор'].nunique()}")
    print(f"   Уникальных карт: {df['Номер карты'].nunique():,}")
    print(f"\n📁 Файл: {os.path.abspath(csv_filename)}")

# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 80)
    print("СБОРЩИК ВСЕХ ДАННЫХ В EXCEL")
    print("=" * 80)
    
    # Аутентификация
    creds = authenticate()
    if not creds:
        return
    
    # Создаем сервис
    try:
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        print("✅ Google Sheets API подключен")
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return
    
    # Получаем список операторов
    operators = get_operator_list(service)
    if not operators:
        print("⚠️ Нет операторов для обработки")
        return
    
    # Собираем данные
    start_time = time.time()
    all_data = collect_all_data(service, operators)
    elapsed = time.time() - start_time
    
    print(f"\n⏱️ Время сбора: {elapsed/60:.1f} минут")
    
    if not all_data:
        print("⚠️ Нет данных для сохранения")
        return
    
    # Сохраняем в Excel
    save_to_excel(all_data, OUTPUT_EXCEL)
    
    print("\n✅ ГОТОВО!")
    print(f"📁 Откройте файл: {OUTPUT_EXCEL}")

if __name__ == '__main__':
    main()
