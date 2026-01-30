"""
Синхронизатор цветового форматирования для всех таблиц операторов
Подключается через Google Sheets API и применяет:
1. Выпадающий список со статусами в колонке E
2. Условное форматирование с цветами для каждого статуса
"""

import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Области доступа
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ID мастер-таблицы
MASTER_SPREADSHEET_ID = "1wlqqSCV3HW5ZgfYUT6IS2Ne466jJQeEKH1Nl4Tx2jdc"

# Список статусов для выпадающего списка
STATUS_LIST = [
    "отрицательный",
    "положительный",
    "тишине",
    "соед прервано",
    "НЕТ ОТВЕТА (ЗАНЯТО)",
    "заявка закрыта (не удалось дозвониться)",
    "открыть карту",
    "тиббиёт ходими аризаси"
]

# Цветовая схема (RGB 0-1)
def hex_to_rgb(hex_color):
    """Конвертирует HEX в RGB (0-1)"""
    hex_color = hex_color.lstrip('#')
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return {"red": r/255.0, "green": g/255.0, "blue": b/255.0}

STATUS_COLORS = {
    "отрицательный": {"bg": "#ff6666", "fg": "#000000"},
    "положительный": {"bg": "#99ff99", "fg": "#000000"},
    "тишине": {"bg": "#ffd9d9", "fg": "#000000"},
    "соед прервано": {"bg": "#ffd9d9", "fg": "#000000"},
    "НЕТ ОТВЕТА (ЗАНЯТО)": {"bg": "#ffff99", "fg": "#000000"},
    "заявка закрыта (не удалось дозвониться)": {"bg": "#d9d9d9", "fg": "#000000"},
    "открыть карту": {"bg": "#99d9ff", "fg": "#000000"},
    "тиббиёт ходими аризаси": {"bg": "#b3e6ff", "fg": "#000000"}
}

def authenticate():
    """Аутентификация в Google API"""
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def get_sheet_id(service, spreadsheet_id, sheet_name="FIKSA"):
    """Получает ID листа по имени"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        return 0
    except:
        return 0

def create_data_validation_request(sheet_id):
    """Создает запрос для выпадающего списка"""
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "startColumnIndex": 4,
                "endColumnIndex": 5
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": status} for status in STATUS_LIST]
                },
                "showCustomUi": True,
                "strict": False
            }
        }
    }

def create_conditional_format_requests(sheet_id):
    """Создает запросы условного форматирования"""
    requests = []
    
    for status, colors in STATUS_COLORS.items():
        bg_color = hex_to_rgb(colors["bg"])
        fg_color = hex_to_rgb(colors["fg"])
        
        request = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "startColumnIndex": 4,
                        "endColumnIndex": 5
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{"userEnteredValue": status}]
                        },
                        "format": {
                            "backgroundColor": bg_color,
                            "textFormat": {"foregroundColor": fg_color}
                        }
                    }
                },
                "index": 0
            }
        }
        requests.append(request)
    
    return requests

def clear_existing_rules(service, spreadsheet_id, sheet_id):
    """Удаляет существующие правила условного форматирования для колонки E"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        delete_requests = []
        for sheet in sheets:
            if sheet['properties']['sheetId'] == sheet_id:
                rules = sheet.get('conditionalFormats', [])
                for i, rule in enumerate(rules):
                    ranges = rule.get('ranges', [])
                    for r in ranges:
                        if r.get('startColumnIndex') == 4 and r.get('endColumnIndex') == 5:
                            delete_requests.append({
                                "deleteConditionalFormatRule": {
                                    "sheetId": sheet_id,
                                    "index": i
                                }
                            })
                            break
        
        if delete_requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": delete_requests}
            ).execute()
    except:
        pass

def apply_formatting(service, spreadsheet_id, operator_name):
    """Применяет форматирование к одной таблице"""
    try:
        sheet_id = get_sheet_id(service, spreadsheet_id, "FIKSA")
        
        # Удаляем старые правила
        clear_existing_rules(service, spreadsheet_id, sheet_id)
        
        # Создаем запросы
        requests = []
        
        # Добавляем выпадающий список
        requests.append(create_data_validation_request(sheet_id))
        
        # Добавляем условное форматирование
        requests.extend(create_conditional_format_requests(sheet_id))
        
        # Применяем все изменения
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
        
        print(f"   ✅ {operator_name}")
        return True
        
    except HttpError as e:
        print(f"   ❌ {operator_name}: {e}")
        return False
    except Exception as e:
        print(f"   ❌ {operator_name}: {e}")
        return False

def get_operators(service):
    """Получает список операторов из мастер-таблицы"""
    try:
        import socket
        socket.setdefaulttimeout(60)  # 60 секунд таймаут
        
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range='Настройки!A2:C100'
        ).execute()
        
        values = result.get('values', [])
        operators = []
        
        for row in values:
            if len(row) >= 3:
                name = row[0].strip()
                spreadsheet_id = row[1].strip()
                status = row[2].strip().lower()
                
                if status == "активен" and spreadsheet_id != "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
                    operators.append({
                        'name': name,
                        'id': spreadsheet_id
                    })
        
        return operators
    except Exception as e:
        print(f"❌ Ошибка получения списка: {e}")
        return []

def main():
    print("=" * 70)
    print("🎨 СИНХРОНИЗАЦИЯ ЦВЕТОВ СТАТУСОВ")
    print("=" * 70)
    print()
    
    print("🔐 Авторизация...")
    try:
        creds = authenticate()
        service = build('sheets', 'v4', credentials=creds)
        print("✅ Авторизация успешна")
    except Exception as e:
        print(f"❌ Ошибка авторизации: {e}")
        return
    
    print()
    print("📋 Получение списка операторов...")
    operators = get_operators(service)
    
    if not operators:
        print("❌ Нет операторов для обработки")
        return
    
    print(f"✅ Найдено операторов: {len(operators)}")
    print()
    
    print("🚀 Применение форматирования...")
    print()
    
    success = 0
    failed = 0
    
    for i, op in enumerate(operators, 1):
        print(f"[{i}/{len(operators)}] {op['name']}...", end=" ")
        if apply_formatting(service, op['id'], op['name']):
            success += 1
        else:
            failed += 1
    
    print()
    print("=" * 70)
    print("📊 РЕЗУЛЬТАТЫ:")
    print(f"   ✅ Успешно: {success}")
    print(f"   ❌ Ошибок: {failed}")
    print("=" * 70)
    print()
    print("🎨 Применено:")
    print("   • Выпадающий список со статусами в колонке E")
    print("   • Цветное форматирование (светлая гамма)")
    print()

if __name__ == '__main__':
    main()
