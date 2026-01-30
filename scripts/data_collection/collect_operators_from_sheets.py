"""
Сбор всех уникальных ФИО операторов из 8-й колонки всех листов
"""
import os

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
print("СБОР ВСЕХ УНИКАЛЬНЫХ ФИО ОПЕРАТОРОВ ИЗ 8-Й КОЛОНКИ")
print("=" * 80)

# Подключение
print("\n🔗 Подключение к Google Sheets...")
creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

# Получаем список операторов
print("\n📋 Получение списка таблиц...")
result = service.spreadsheets().values().get(
    spreadsheetId=MASTER_SPREADSHEET_ID,
    range=f"{SETTINGS_SHEET_NAME}!A2:C100"
).execute()

values = result.get('values', [])
operators = []

for row in values:
    if len(row) == 0:
        continue
    spreadsheet_id = row[1].strip() if len(row) > 1 and row[1] else ""
    if spreadsheet_id and spreadsheet_id != "ID таблицы":
        name = row[0].strip() if len(row) > 0 and row[0] else "Без имени"
        operators.append({
            'name': name,
            'spreadsheet_id': spreadsheet_id
        })

print(f"✅ Найдено таблиц: {len(operators)}")

# Собираем всех уникальных операторов
all_operators = set()
total_checked = 0

print("\n🔍 Сбор уникальных ФИО...\n")

for idx, op in enumerate(operators, 1):
    print(f"[{idx}/{len(operators)}] {op['name']}")
    
    try:
        # Получаем листы
        metadata = service.spreadsheets().get(spreadsheetId=op['spreadsheet_id']).execute()
        sheets = metadata.get('sheets', [])
        
        for sheet in sheets:
            title = sheet['properties']['title']
            
            # Пропускаем служебные
            if title in SKIP_SHEETS:
                continue
            
            title_lower = title.lower()
            skip_words = ['setting', 'аризалар', 'настройки', 'сводка', 'статистика']
            if any(word in title_lower for word in skip_words):
                continue
            
            try:
                # Читаем 8-ю колонку (I - это 9-я, но нам нужна 8-я = H)
                # Колонки: B C D E F G H I
                # B=карта, C=тел, D=дата откр, E=статус, F=служба, G=коммент, H=оператор фикс, I=дата фикс
                result = service.spreadsheets().values().get(
                    spreadsheetId=op['spreadsheet_id'],
                    range=f"'{title}'!H2:H5000",  # 8-я колонка
                    valueRenderOption='FORMATTED_VALUE'
                ).execute()
                
                rows = result.get('values', [])
                for row in rows:
                    if len(row) > 0 and row[0] and str(row[0]).strip():
                        operator_name = str(row[0]).strip()
                        if operator_name and operator_name != "":
                            all_operators.add(operator_name)
                
                total_checked += 1
                
            except Exception as e:
                continue
        
    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        continue

# Выводим результат
print("\n" + "=" * 80)
print(f"📊 РЕЗУЛЬТАТ: Найдено {len(all_operators)} уникальных операторов")
print("=" * 80)

sorted_operators = sorted(all_operators)
for i, op in enumerate(sorted_operators, 1):
    print(f"{i:3}. {op}")

# Сохраняем в файл
with open('OPERATORS_LIST.txt', 'w', encoding='utf-8') as f:
    f.write("СПИСОК ВСЕХ УНИКАЛЬНЫХ ОПЕРАТОРОВ\n")
    f.write("=" * 80 + "\n\n")
    for i, op in enumerate(sorted_operators, 1):
        f.write(f"{i}. {op}\n")

print(f"\n✅ Список сохранен в: OPERATORS_LIST.txt")
print(f"   Проверено листов: {total_checked}")
print("=" * 80)
