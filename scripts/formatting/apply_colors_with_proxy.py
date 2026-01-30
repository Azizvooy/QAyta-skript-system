"""
=============================================================================
ПРИМЕНЕНИЕ ЦВЕТОВ К СТАТУСАМ (С ПРОКСИ)
=============================================================================
Автоматическое изменение цветов в колонке E для всех таблиц операторов
=============================================================================
"""

import os
import socket
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Настройка прокси (как в export_all_sheets.py)
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

# Увеличиваем таймаут
socket.setdefaulttimeout(120)

# Настройки
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"

# Цветовая схема (светлая палитра)
STATUS_COLORS = {
    "отрицательный": {"red": 1.0, "green": 0.4, "blue": 0.4},      # #ff6666
    "положительный": {"red": 0.6, "green": 1.0, "blue": 0.6},      # #99ff99
    "тишине": {"red": 1.0, "green": 0.85, "blue": 0.85},           # #ffd9d9
    "соед прервано": {"red": 1.0, "green": 0.85, "blue": 0.85},    # #ffd9d9
    "НЕТ ОТВЕТА (ЗАНЯТО)": {"red": 1.0, "green": 1.0, "blue": 0.6}, # #ffff99
    "заявка закрыта": {"red": 0.85, "green": 0.85, "blue": 0.85},  # #d9d9d9
    "открыть карту": {"red": 0.6, "green": 0.85, "blue": 1.0},     # #99d9ff
    "тиббиёт ходими аризаси": {"red": 0.7, "green": 0.9, "blue": 1.0} # #b3e6ff
}

# =============================================================================
# АУТЕНТИФИКАЦИЯ
# =============================================================================

def authenticate():
    """Аутентификация в Google API (как в export_all_sheets.py)"""
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
# ПОЛУЧЕНИЕ СПИСКА ОПЕРАТОРОВ
# =============================================================================

def get_operator_list(sheets_service):
    """Читает список операторов из мастер-таблицы"""
    print(f"\n📋 Чтение списка операторов...")
    
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
                status = row[2].strip() if len(row) > 2 else "активен"
                
                if name and spreadsheet_id and spreadsheet_id != "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
                    operators.append({
                        "name": name,
                        "spreadsheet_id": spreadsheet_id,
                        "status": status
                    })
        
        print(f"✅ Найдено операторов: {len(operators)}")
        return operators
        
    except HttpError as error:
        print(f"❌ Ошибка: {error}")
        return []

# =============================================================================
# СОЗДАНИЕ ПРАВИЛ ФОРМАТИРОВАНИЯ
# =============================================================================

def create_conditional_format_requests(sheet_id):
    """Создает правила условного форматирования для колонки E"""
    requests = []
    
    for status_text, color in STATUS_COLORS.items():
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": sheet_id,
                        "startColumnIndex": 4,  # Колонка E
                        "endColumnIndex": 5
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{"userEnteredValue": status_text}]
                        },
                        "format": {
                            "backgroundColor": color
                        }
                    }
                },
                "index": 0
            }
        })
    
    return requests

def create_data_validation_request(sheet_id, setting_sheet_id):
    """Создает правило валидации данных (выпадающий список из листа SETTING)"""
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startColumnIndex": 4,  # Колонка E
                "endColumnIndex": 5,
                "startRowIndex": 1  # Со второй строки (пропускаем заголовок)
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{
                        "userEnteredValue": "=SETTING!$E$2:$E$20"  # Диапазон со статусами
                    }]
                },
                "showCustomUi": True,
                "strict": True
            }
        }
    }

# =============================================================================
# ПРИМЕНЕНИЕ ФОРМАТИРОВАНИЯ
# =============================================================================

def apply_formatting_to_spreadsheet(sheets_service, spreadsheet_id, operator_name):
    """Применяет форматирование к одной таблице"""
    try:
        # Получаем информацию о таблице
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        if not sheets:
            print(f"  ⚠️  Нет листов в таблице")
            return False
        
        # Находим листы SETTING и основной лист
        setting_sheet_id = None
        main_sheet_id = None
        main_sheet_title = None
        
        for sheet in sheets:
            props = sheet.get('properties', {})
            title = props.get('title', '')
            sheet_id = props.get('sheetId', 0)
            
            if title == 'SETTING':
                setting_sheet_id = sheet_id
            elif main_sheet_id is None and title not in ['Статистика', 'Предыдущий месяц', 'Сводка по дням', 'Настройки']:
                main_sheet_id = sheet_id
                main_sheet_title = title
        
        if main_sheet_id is None:
            # Если не нашли основной лист, берем первый
            main_sheet_id = sheets[0]['properties']['sheetId']
            main_sheet_title = sheets[0]['properties']['title']
        
        print(f"  📄 Лист: {main_sheet_title}")
        
        # Формируем запросы
        requests = []
        
        # 1. Добавляем выпадающий список (только если есть SETTING)
        if setting_sheet_id is not None:
            requests.append(create_data_validation_request(main_sheet_id, setting_sheet_id))
        
        # 2. Добавляем условное форматирование
        requests.extend(create_conditional_format_requests(main_sheet_id))
        
        # Выполняем batchUpdate
        body = {'requests': requests}
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        
        print(f"  ✅ Форматирование применено")
        return True
        
    except HttpError as error:
        print(f"  ❌ Ошибка: {error}")
        return False
        print(f"  ✅ Форматирование применено")
        return True
        
    except HttpError as error:
        print(f"  ❌ Ошибка: {error}")
        return False

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    print("=" * 80)
    print("🎨 ПРИМЕНЕНИЕ ЦВЕТОВ К СТАТУСАМ (ВСЕ ОПЕРАТОРЫ)")
    print("=" * 80)
    
    # Аутентификация
    creds = authenticate()
    if not creds:
        return
    
    # Создаем сервис
    print("\n🔧 Создание сервиса Google Sheets...")
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    print("✅ Сервис создан")
    
    # Получаем список операторов
    operators = get_operator_list(service)
    
    if not operators:
        print("\n❌ Операторы не найдены!")
        return
    
    # Применяем форматирование
    print(f"\n🚀 Начинаем применение форматирования к {len(operators)} таблицам...\n")
    
    success_count = 0
    fail_count = 0
    
    for i, operator in enumerate(operators, 1):
        print(f"\n[{i}/{len(operators)}] {operator['name']}")
        
        if apply_formatting_to_spreadsheet(service, operator['spreadsheet_id'], operator['name']):
            success_count += 1
        else:
            fail_count += 1
    
    # Итоги
    print("\n" + "=" * 80)
    print("📊 ИТОГИ:")
    print(f"   ✅ Успешно: {success_count}")
    print(f"   ❌ Ошибки: {fail_count}")
    print("=" * 80)

if __name__ == "__main__":
    main()
