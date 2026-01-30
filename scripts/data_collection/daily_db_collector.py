"""
=============================================================================
ЕЖЕДНЕВНЫЙ СБОР ДАННЫХ О ФИКСАЦИИ В БАЗУ ДАННЫХ
=============================================================================
Собирает данные из всех таблиц операторов и сохраняет в SQLite БД
=============================================================================
"""

import os
import sqlite3
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import socket
import json

# Настройка прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

# Настройки
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]
TOKEN_FILE = 'config/token.json'
CREDENTIALS_FILE = 'config/credentials.json'
DB_PATH = 'data/fiksa_database.db'
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"

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
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print("❌ Файл credentials.json не найден!")
                return None
            
            print("🔐 Авторизация...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print("✅ Токен сохранен")
    
    return creds

# =============================================================================
# БАЗА ДАННЫХ
# =============================================================================

def init_database():
    """Создает структуру базы данных"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Таблица с данными о фиксациях
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fiksa_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            collection_date DATE NOT NULL,
            operator_name TEXT NOT NULL,
            card_number TEXT,
            full_name TEXT,
            phone TEXT,
            address TEXT,
            status TEXT,
            call_date DATETIME,
            notes TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Таблица со статистикой по операторам
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS operator_stats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            collection_date DATE NOT NULL,
            operator_name TEXT NOT NULL,
            total_records INTEGER DEFAULT 0,
            status_counts TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(collection_date, operator_name)
        )
    ''')
    
    # Индексы для быстрого поиска
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_collection_date ON fiksa_records(collection_date)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_operator ON fiksa_records(operator_name)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_status ON fiksa_records(status)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_card ON fiksa_records(card_number)')
    
    conn.commit()
    conn.close()
    print("✅ База данных инициализирована")

# =============================================================================
# ПОЛУЧЕНИЕ ДАННЫХ ИЗ GOOGLE SHEETS
# =============================================================================

def get_operators_list(sheets_service):
    """Получает список операторов из мастер-таблицы"""
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=f"{SETTINGS_SHEET_NAME}!A2:C100"
        ).execute()
        
        values = result.get('values', [])
        operators = []
        
        for row in values:
            if len(row) >= 2:
                name = row[0].strip() if len(row) > 0 else ""
                spreadsheet_id = row[1].strip() if len(row) > 1 else ""
                
                if name and spreadsheet_id and spreadsheet_id != "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
                    operators.append({
                        "name": name,
                        "spreadsheet_id": spreadsheet_id
                    })
        
        return operators
        
    except HttpError as error:
        print(f"❌ Ошибка получения списка операторов: {error}")
        return []

def get_operator_data(sheets_service, spreadsheet_id, operator_name):
    """Получает данные с листа FIKSA оператора"""
    try:
        # Получаем структуру таблицы
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        # Находим лист с точным названием FIKSA (не старые листы с датами)
        sheet_name = None
        all_sheets = []
        for sheet in sheets:
            title = sheet['properties']['title']
            all_sheets.append(title)
            # Ищем ТОЧНОЕ совпадение "FIKSA" (без дат и других приставок)
            if title.upper() == 'FIKSA':
                sheet_name = title
                break
        
        if not sheet_name:
            # Если нет FIKSA, пропускаем
            print(f"  ⚠️  Нет листа FIKSA. Доступные листы: {', '.join(all_sheets)}")
            return []
        
        print(f"  📄 Лист: {sheet_name}")
        
        # Читаем данные (увеличиваем лимит до 10000 строк)
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A2:Z10000"
        ).execute()
        
        values = result.get('values', [])
        
        # Преобразуем в структурированный формат
        records = []
        for row in values:
            if len(row) >= 5:  # Минимум должно быть несколько колонок
                card_number = row[0].strip() if len(row) > 0 else ''
                full_name = row[1].strip() if len(row) > 1 else ''
                
                # Пропускаем пустые строки (нет ни карты, ни имени)
                if not card_number and not full_name:
                    continue
                
                record = {
                    'operator_name': operator_name,
                    'card_number': card_number,
                    'full_name': full_name,
                    'phone': row[2].strip() if len(row) > 2 else '',
                    'address': row[3].strip() if len(row) > 3 else '',
                    'status': row[4].strip() if len(row) > 4 else '',  # Колонка E
                    'call_date': row[5].strip() if len(row) > 5 else '',
                    'notes': row[6].strip() if len(row) > 6 else ''
                }
                records.append(record)
        
        return records
        
    except HttpError as error:
        print(f"  ⚠️  Ошибка: {error}")
        return []

# =============================================================================
# СОХРАНЕНИЕ В БД
# =============================================================================

def save_to_database(records, collection_date):
    """Сохраняет записи в базу данных"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    saved_count = 0
    
    for record in records:
        try:
            cursor.execute('''
                INSERT INTO fiksa_records 
                (collection_date, operator_name, card_number, full_name, phone, 
                 address, status, call_date, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                collection_date,
                record['operator_name'],
                record['card_number'],
                record['full_name'],
                record['phone'],
                record['address'],
                record['status'],
                record['call_date'],
                record['notes']
            ))
            saved_count += 1
        except sqlite3.Error as e:
            print(f"  ⚠️  Ошибка сохранения записи: {e}")
    
    conn.commit()
    conn.close()
    
    return saved_count

def save_operator_stats(operator_name, total_records, status_counts, collection_date):
    """Сохраняет статистику по оператору"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    status_json = json.dumps(status_counts, ensure_ascii=False)
    
    cursor.execute('''
        INSERT OR REPLACE INTO operator_stats 
        (collection_date, operator_name, total_records, status_counts)
        VALUES (?, ?, ?, ?)
    ''', (collection_date, operator_name, total_records, status_json))
    
    conn.commit()
    conn.close()

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    print("=" * 80)
    print("📊 ЕЖЕДНЕВНЫЙ СБОР ДАННЫХ О ФИКСАЦИИ В БД")
    print("=" * 80)
    
    # Инициализация БД
    init_database()
    
    # Аутентификация
    creds = authenticate()
    if not creds:
        return
    
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    
    # Текущая дата
    collection_date = datetime.now().date()
    print(f"📅 Дата сбора: {collection_date}")
    
    # Получаем список операторов
    print("\n📋 Получение списка операторов...")
    operators = get_operators_list(service)
    print(f"✅ Найдено операторов: {len(operators)}")
    
    # Собираем данные
    print(f"\n🚀 Начинаем сбор данных...\n")
    
    total_saved = 0
    
    for i, operator in enumerate(operators, 1):
        # Пропускаем операторов с пустым именем
        if not operator['name'] or operator['name'].strip() == '-':
            print(f"[{i}/{len(operators)}] (пропущен - пустое имя)")
            continue
            
        print(f"[{i}/{len(operators)}] {operator['name']}")
        print(f"[{i}/{len(operators)}] {operator['name']}")
        
        # Получаем данные
        records = get_operator_data(service, operator['spreadsheet_id'], operator['name'])
        
        if records:
            # Сохраняем в БД
            saved = save_to_database(records, collection_date)
            total_saved += saved
            
            # Подсчет статистики
            status_counts = {}
            for record in records:
                status = record['status']
                status_counts[status] = status_counts.get(status, 0) + 1
            
            # Сохраняем статистику
            save_operator_stats(operator['name'], len(records), status_counts, collection_date)
            
            print(f"  ✅ Сохранено записей: {saved}")
        else:
            print(f"  ⚠️  Нет данных")
    
    # Итоги
    print("\n" + "=" * 80)
    print(f"📊 ИТОГО сохранено записей: {total_saved}")
    print(f"💾 База данных: {DB_PATH}")
    print("=" * 80)

if __name__ == "__main__":
    main()
