"""
Простой синхронизатор цветов - работает оффлайн
Просто укажите ID таблиц ниже и запустите
"""

import pickle
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials  
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ===== УКАЖИТЕ ID ТАБЛИЦ ЗДЕСЬ =====
# Формат: "Имя оператора": "ID_таблицы"
OPERATORS = {
    # Пример (замените на реальные ID):
    # "Abdullayeva": "1abc123...",
    # "Karimova": "1def456...",
}

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

def hex_to_rgb(hex_color):
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

def auth():
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

def apply(service, spreadsheet_id, name):
    try:
        # Получаем sheet_id
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = 0
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == 'FIKSA':
                sheet_id = sheet['properties']['sheetId']
                break
        
        requests = []
        
        # Выпадающий список
        requests.append({
            "setDataValidation": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "startColumnIndex": 4, "endColumnIndex": 5},
                "rule": {
                    "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": s} for s in STATUS_LIST]},
                    "showCustomUi": True, "strict": False
                }
            }
        })
        
        # Цвета
        for status, colors in STATUS_COLORS.items():
            requests.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, "startColumnIndex": 4, "endColumnIndex": 5}],
                        "booleanRule": {
                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": status}]},
                            "format": {
                                "backgroundColor": hex_to_rgb(colors["bg"]),
                                "textFormat": {"foregroundColor": hex_to_rgb(colors["fg"])}
                            }
                        }
                    },
                    "index": 0
                }
            })
        
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
        print(f"   ✅ {name}")
        return True
    except Exception as e:
        print(f"   ❌ {name}: {e}")
        return False

def main():
    if not OPERATORS:
        print("❌ Добавьте ID таблиц в переменную OPERATORS в этом файле!")
        print("   Формат: \"Имя\": \"ID_таблицы\"")
        return
    
    print("🔐 Авторизация...")
    creds = auth()
    service = build('sheets', 'v4', credentials=creds)
    print("✅ Готово\n")
    
    print(f"🚀 Обработка {len(OPERATORS)} таблиц...\n")
    
    success = 0
    for name, spreadsheet_id in OPERATORS.items():
        if apply(service, spreadsheet_id, name):
            success += 1
    
    print(f"\n✅ Успешно: {success}/{len(OPERATORS)}")

if __name__ == '__main__':
    main()
