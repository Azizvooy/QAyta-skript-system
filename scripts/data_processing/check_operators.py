"""
Проверка таблицы настроек - сколько там операторов
"""

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os

os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"
TOKEN_FILE = 'token.json'

creds = Credentials.from_authorized_user_file(TOKEN_FILE)
service = build('sheets', 'v4', credentials=creds)

# Читаем все строки
result = service.spreadsheets().values().get(
    spreadsheetId=MASTER_SPREADSHEET_ID,
    range=f"{SETTINGS_SHEET_NAME}!A1:C200"
).execute()

values = result.get('values', [])

print(f"Всего строк в таблице настроек: {len(values)}")
print("\n" + "="*80)
print("СПИСОК ВСЕХ ОПЕРАТОРОВ:")
print("="*80)

count = 0
for idx, row in enumerate(values, 1):
    if idx == 1:  # Заголовок
        print(f"\nЗаголовки: {row}")
        print("-"*80)
        continue
    
    if len(row) >= 2:
        name = row[0].strip() if len(row) > 0 else ""
        spreadsheet_id = row[1].strip() if len(row) > 1 else ""
        status = row[2].strip() if len(row) > 2 else ""
        
        if name and spreadsheet_id and spreadsheet_id != "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
            count += 1
            print(f"{count:2}. {name:50} | {status:15} | {spreadsheet_id[:20]}...")

print(f"\n{'='*80}")
print(f"ИТОГО операторов с заполненными ID: {count}")
