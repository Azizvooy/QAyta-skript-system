"""
Быстрый тест - применяет форматирование к одной таблице
"""

import pickle
import os
import socket
import httplib2
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Увеличиваем таймауты
socket.setdefaulttimeout(120)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Тестовая таблица (мастер-таблица)
TEST_SPREADSHEET_ID = "1wlqqSCV3HW5ZgfYUT6IS2Ne466jJQeEKH1Nl4Tx2jdc"

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

service = build('sheets', 'v4', credentials=creds)

print("Поиск листа FIKSA...")
spreadsheet = service.spreadsheets().get(spreadsheetId=TEST_SPREADSHEET_ID).execute()
sheet_id = None

for sheet in spreadsheet.get('sheets', []):
    title = sheet['properties']['title']
    print(f"   Найден лист: {title}")
    if title == 'FIKSA':
        sheet_id = sheet['properties']['sheetId']
        print(f"   OK Лист FIKSA найден, ID: {sheet_id}")

if not sheet_id:
    print("ОШИБКА: Лист FIKSA не найден!")
    exit()

print("\nСоздание запросов...")
requests = []

# Выпадающий список
requests.append({
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
                "values": [{"userEnteredValue": s} for s in STATUS_LIST]
            },
            "showCustomUi": True,
            "strict": False
        }
    }
})

print(f"   OK Выпадающий список: {len(STATUS_LIST)} статусов")

# Цвета
for status, colors in STATUS_COLORS.items():
    requests.append({
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
                        "backgroundColor": hex_to_rgb(colors["bg"]),
                        "textFormat": {"foregroundColor": hex_to_rgb(colors["fg"])}
                    }
                }
            },
            "index": 0
        }
    })

print(f"   OK Правила форматирования: {len(STATUS_COLORS)}")

print("\nПрименение изменений...")
response = service.spreadsheets().batchUpdate(
    spreadsheetId=TEST_SPREADSHEET_ID,
    body={"requests": requests}
).execute()

print(f"\nГОТОВО! Применено {len(requests)} изменений")
print(f"   Выпадающий список со статусами")
print(f"   Цветное форматирование (светлая гамма)")
