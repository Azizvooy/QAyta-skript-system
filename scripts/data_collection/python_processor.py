"""
=============================================================================
PYTHON ОБРАБОТЧИК ДАННЫХ ИЗ GOOGLE DOCS
=============================================================================
Версия: 1.0
Дата: 01.12.2025

📋 НАЗНАЧЕНИЕ:
Читает данные из Google Docs через API, обрабатывает их и записывает
результаты обратно в Google Sheets

🔄 ПРОЦЕСС:
1. Читает JSON Lines из Google Docs
2. Обрабатывает данные (статистика, фильтрация, группировка)
3. Записывает результаты в Google Sheets через API

📦 УСТАНОВКА:
pip install google-auth google-auth-oauthlib google-auth-httplib2
pip install google-api-python-client pandas

🔑 НАСТРОЙКА:
1. Создайте проект в Google Cloud Console
2. Включите APIs: Google Docs API, Google Sheets API
3. Создайте OAuth 2.0 credentials
4. Скачайте credentials.json
5. Поместите в папку со скриптом

🚀 ИСПОЛЬЗОВАНИЕ:
python python_processor.py
=============================================================================
"""

import json
import os
from datetime import datetime
from collections import defaultdict
from typing import List, Dict, Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# =============================================================================
# НАСТРОЙКИ
# =============================================================================

# Области доступа (scopes)
SCOPES = [
    'https://www.googleapis.com/auth/documents.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# ID документа Google Docs с данными
DOCS_ID = "ВСТАВЬТЕ_ID_ДОКУМЕНТА_ЗДЕСЬ"

# ID таблицы Google Sheets для результатов
SHEETS_ID = "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ"

# Файлы для хранения токенов
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'

# =============================================================================
# АУТЕНТИФИКАЦИЯ
# =============================================================================

def authenticate():
    """
    Аутентификация в Google API
    
    При первом запуске откроется браузер для авторизации
    Токен сохраняется в token.json для последующих запусков
    """
    creds = None
    
    # Проверяем существующий токен
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    # Если токена нет или он недействителен - авторизуемся
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Обновление токена...")
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print("❌ Файл credentials.json не найден!")
                print("\nИнструкция:")
                print("1. Перейдите: https://console.cloud.google.com")
                print("2. Создайте проект")
                print("3. Включите Google Docs API и Google Sheets API")
                print("4. Создайте OAuth 2.0 credentials (Desktop app)")
                print("5. Скачайте JSON и сохраните как credentials.json")
                return None
            
            print("🔐 Авторизация (откроется браузер)...")
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Сохраняем токен
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print("✅ Токен сохранен")
    
    return creds

# =============================================================================
# ЧТЕНИЕ ДАННЫХ ИЗ GOOGLE DOCS
# =============================================================================

def read_docs_data(docs_service, document_id: str) -> List[Dict[str, Any]]:
    """
    Читает данные из Google Docs
    
    Args:
        docs_service: Сервис Google Docs API
        document_id: ID документа
        
    Returns:
        Список записей (словарей)
    """
    print(f"\n📄 Чтение документа {document_id}...")
    
    try:
        # Читаем документ
        document = docs_service.documents().get(documentId=document_id).execute()
        
        content = document.get('body').get('content')
        
        records = []
        line_count = 0
        
        # Проходим по всем параграфам
        for element in content:
            if 'paragraph' in element:
                paragraph = element.get('paragraph')
                elements = paragraph.get('elements')
                
                for elem in elements:
                    if 'textRun' in elem:
                        text = elem.get('textRun').get('content').strip()
                        
                        # Пропускаем пустые строки и заголовки
                        if not text or not text.startswith('{'):
                            continue
                        
                        # Парсим JSON
                        try:
                            record = json.loads(text)
                            records.append(record)
                            line_count += 1
                        except json.JSONDecodeError:
                            continue
        
        print(f"✅ Прочитано записей: {len(records)}")
        return records
        
    except HttpError as error:
        print(f"❌ Ошибка чтения документа: {error}")
        return []

# =============================================================================
# ОБРАБОТКА ДАННЫХ
# =============================================================================

def process_data(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Обрабатывает данные и создает статистику
    
    Args:
        records: Список записей
        
    Returns:
        Словарь со статистикой
    """
    print("\n📊 Обработка данных...")
    
    # Статистика по операторам
    operator_stats = defaultdict(lambda: {
        'total_records': 0,
        'unique_cards': set(),
        'statuses': defaultdict(int),
        'by_month': defaultdict(int)
    })
    
    # Статистика по месяцам
    monthly_stats = defaultdict(lambda: {
        'total_records': 0,
        'unique_cards': set(),
        'operators': set(),
        'statuses': defaultdict(int)
    })
    
    # Обрабатываем каждую запись
    for record in records:
        operator = record.get('operator', 'Неизвестно')
        card = record.get('card', '')
        status = record.get('status', 'Неизвестно')
        date_str = record.get('date', '')
        
        # Парсим дату
        try:
            date_obj = datetime.strptime(date_str, '%d.%m.%Y %H:%M:%S')
            month_key = date_obj.strftime('%m.%Y')
        except:
            month_key = 'Неизвестно'
        
        # Статистика по оператору
        operator_stats[operator]['total_records'] += 1
        operator_stats[operator]['unique_cards'].add(card)
        operator_stats[operator]['statuses'][status] += 1
        operator_stats[operator]['by_month'][month_key] += 1
        
        # Статистика по месяцам
        monthly_stats[month_key]['total_records'] += 1
        monthly_stats[month_key]['unique_cards'].add(card)
        monthly_stats[month_key]['operators'].add(operator)
        monthly_stats[month_key]['statuses'][status] += 1
    
    print(f"✅ Обработано записей: {len(records)}")
    print(f"   Операторов: {len(operator_stats)}")
    print(f"   Месяцев: {len(monthly_stats)}")
    
    return {
        'operators': operator_stats,
        'monthly': monthly_stats,
        'total_records': len(records)
    }

# =============================================================================
# ЗАПИСЬ В GOOGLE SHEETS
# =============================================================================

def write_to_sheets(sheets_service, spreadsheet_id: str, stats: Dict[str, Any]):
    """
    Записывает статистику в Google Sheets
    
    Args:
        sheets_service: Сервис Google Sheets API
        spreadsheet_id: ID таблицы
        stats: Статистика для записи
    """
    print(f"\n📝 Запись результатов в таблицу {spreadsheet_id}...")
    
    try:
        # Лист 1: Статистика по операторам
        operator_data = [
            ['ФИО оператора', 'Всего записей', 'Уникальных карт', 'Статусы']
        ]
        
        for operator, data in sorted(stats['operators'].items()):
            operator_data.append([
                operator,
                data['total_records'],
                len(data['unique_cards']),
                ', '.join([f"{k}: {v}" for k, v in data['statuses'].items()])
            ])
        
        # Записываем в лист "Статистика по операторам"
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range='Статистика по операторам!A1',
            valueInputOption='RAW',
            body={'values': operator_data}
        ).execute()
        
        print(f"✅ Записано операторов: {len(operator_data) - 1}")
        
        # Лист 2: Статистика по месяцам
        monthly_data = [
            ['Месяц', 'Всего записей', 'Уникальных карт', 'Операторов', 'Статусы']
        ]
        
        for month, data in sorted(stats['monthly'].items(), reverse=True):
            monthly_data.append([
                month,
                data['total_records'],
                len(data['unique_cards']),
                len(data['operators']),
                ', '.join([f"{k}: {v}" for k, v in data['statuses'].items()])
            ])
        
        # Записываем в лист "Статистика по месяцам"
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range='Статистика по месяцам!A1',
            valueInputOption='RAW',
            body={'values': monthly_data}
        ).execute()
        
        print(f"✅ Записано месяцев: {len(monthly_data) - 1}")
        
        # Лист 3: Общая информация
        summary_data = [
            ['Параметр', 'Значение'],
            ['Дата обработки', datetime.now().strftime('%d.%m.%Y %H:%M:%S')],
            ['Всего записей', stats['total_records']],
            ['Операторов', len(stats['operators'])],
            ['Месяцев', len(stats['monthly'])]
        ]
        
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range='Общая информация!A1',
            valueInputOption='RAW',
            body={'values': summary_data}
        ).execute()
        
        print("✅ Записана общая информация")
        
    except HttpError as error:
        print(f"❌ Ошибка записи в таблицу: {error}")

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    """
    Основная функция обработки
    """
    print("=" * 80)
    print("PYTHON ОБРАБОТЧИК ДАННЫХ ИЗ GOOGLE DOCS")
    print("=" * 80)
    
    # Проверка конфигурации
    if DOCS_ID == "ВСТАВЬТЕ_ID_ДОКУМЕНТА_ЗДЕСЬ":
        print("\n❌ Необходимо настроить DOCS_ID в коде")
        return
    
    if SHEETS_ID == "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
        print("\n❌ Необходимо настроить SHEETS_ID в коде")
        return
    
    # Аутентификация
    creds = authenticate()
    if not creds:
        return
    
    # Создаем сервисы
    try:
        docs_service = build('docs', 'v1', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        print("✅ Сервисы API подключены")
    except Exception as e:
        print(f"❌ Ошибка подключения к API: {e}")
        return
    
    # Читаем данные
    records = read_docs_data(docs_service, DOCS_ID)
    
    if not records:
        print("\n⚠️ Нет данных для обработки")
        return
    
    # Обрабатываем данные
    stats = process_data(records)
    
    # Записываем результаты
    write_to_sheets(sheets_service, SHEETS_ID, stats)
    
    print("\n" + "=" * 80)
    print("✅ ОБРАБОТКА ЗАВЕРШЕНА")
    print("=" * 80)

if __name__ == '__main__':
    main()
