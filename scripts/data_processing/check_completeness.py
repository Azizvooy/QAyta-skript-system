"""
Проверка полноты собранных данных
Сравниваем что в CSV с тем что в исходных таблицах
"""
import pandas as pd
import os
from collections import defaultdict

# Настройка прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import socket
socket.setdefaulttimeout(120)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
TOKEN_FILE = 'token.json'
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"
SKIP_SHEETS = ["Статистика", "Предыдущий месяц", "Сводка по дням", "Настройки"]

print("=" * 80)
print("ПРОВЕРКА ПОЛНОТЫ ДАННЫХ")
print("=" * 80)

# 1. Читаем CSV
print("\n📖 Читаем собранные данные (CSV)...")
df = pd.read_csv('ALL_DATA.csv', encoding='utf-8-sig')
print(f"✅ В CSV: {len(df):,} строк")

# Группируем по операторам
csv_stats = df.groupby(['Оператор', 'Архивный лист']).size().to_dict()
csv_by_operator = df.groupby('Оператор').size().to_dict()

print(f"\n📊 Статистика CSV:")
print(f"   Операторов: {df['Оператор'].nunique()}")
print(f"   Листов: {df.groupby(['Оператор', 'Архивный лист']).ngroups}")

# 2. Подключаемся к Google Sheets
print("\n🔗 Подключение к Google Sheets...")
creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

# 3. Читаем список операторов
print("\n📋 Читаем список операторов...")
result = service.spreadsheets().values().get(
    spreadsheetId=MASTER_SPREADSHEET_ID,
    range=f"{SETTINGS_SHEET_NAME}!A2:C100"
).execute()

values = result.get('values', [])
operators = []

for idx, row in enumerate(values, start=2):
    if len(row) == 0:
        continue
    spreadsheet_id = row[1].strip() if len(row) > 1 and row[1] else ""
    if spreadsheet_id and spreadsheet_id != "ID таблицы":
        name = row[0].strip() if len(row) > 0 and row[0] else f"Оператор {idx}"
        operators.append({
            'name': name if name else f"Оператор {idx}",
            'spreadsheet_id': spreadsheet_id
        })

print(f"✅ Найдено операторов в настройках: {len(operators)}")

# 4. Проверяем каждого оператора
print("\n🔍 Детальная проверка по операторам:\n")

missing_data = []
total_original = 0
total_csv = len(df)

for op in operators:
    operator_name = op['name']
    spreadsheet_id = op['spreadsheet_id']
    
    print(f"▶ {operator_name}")
    
    try:
        # Получаем список листов
        metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = metadata.get('sheets', [])
        
        operator_sheets = []
        for sheet in sheets:
            title = sheet['properties']['title']
            title_lower = title.lower()
            
            if title in SKIP_SHEETS:
                continue
            
            skip_words = ['setting', 'аризалар', 'аризалары', 'настройки', 'сводка', 'статистика']
            if any(word in title_lower for word in skip_words):
                continue
            
            if title.strip():
                operator_sheets.append(title)
        
        print(f"  Листов в таблице: {len(operator_sheets)}")
        
        # Подсчитываем строки в каждом листе
        operator_total = 0
        for sheet_name in operator_sheets:
            try:
                result = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{sheet_name}'!B2:B20000",
                    valueRenderOption='FORMATTED_VALUE'
                ).execute()
                
                rows = result.get('values', [])
                # Считаем непустые строки
                non_empty = sum(1 for row in rows if len(row) > 0 and row[0])
                operator_total += non_empty
                
                # Проверяем в CSV
                csv_key = (operator_name, sheet_name)
                csv_count = csv_stats.get(csv_key, 0)
                
                if csv_count != non_empty:
                    status = "⚠️"
                    missing_data.append({
                        'operator': operator_name,
                        'sheet': sheet_name,
                        'original': non_empty,
                        'csv': csv_count,
                        'diff': non_empty - csv_count
                    })
                else:
                    status = "✓"
                
                print(f"    {status} {sheet_name}: {non_empty} строк (CSV: {csv_count})")
                
            except Exception as e:
                print(f"    ❌ {sheet_name}: ошибка чтения")
        
        total_original += operator_total
        csv_operator_count = csv_by_operator.get(operator_name, 0)
        
        if operator_total != csv_operator_count:
            print(f"  ⚠️ Итого: {operator_total} (CSV: {csv_operator_count}) РАСХОЖДЕНИЕ!")
        else:
            print(f"  ✅ Итого: {operator_total} ✓")
        
    except Exception as e:
        print(f"  ❌ Ошибка доступа: {e}")
    
    print()

# 5. Итоговый отчет
print("=" * 80)
print("📊 ИТОГОВЫЙ ОТЧЕТ")
print("=" * 80)
print(f"Всего строк в исходных таблицах: {total_original:,}")
print(f"Всего строк в CSV:               {total_csv:,}")
print(f"Разница:                         {total_original - total_csv:,}")

if len(missing_data) > 0:
    print(f"\n⚠️ НАЙДЕНЫ РАСХОЖДЕНИЯ: {len(missing_data)}")
    print("\nДетали:")
    for item in missing_data:
        print(f"  • {item['operator']} / {item['sheet']}")
        print(f"    Исходник: {item['original']:,} | CSV: {item['csv']:,} | Разница: {item['diff']:,}")
else:
    print("\n✅ ДАННЫЕ СОВПАДАЮТ ПОЛНОСТЬЮ!")

print("\n" + "=" * 80)
