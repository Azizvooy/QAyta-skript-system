"""
Поиск проблемных строк в Google Sheets:
1. Пустая дата фиксации
2. Неполные строки (пустые ячейки в важных полях)
"""
import os

# Настройка прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import socket
socket.setdefaulttimeout(120)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = 'token.json'
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"
SKIP_SHEETS = ["Статистика", "Предыдущий месяц", "Сводка по дням", "Настройки", "setting", "аризалар"]

print("=" * 100)
print("ПОИСК ПРОБЛЕМНЫХ СТРОК В GOOGLE SHEETS")
print("=" * 100)

# Подключение
print("\n🔗 Подключение к Google Sheets...")
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow

creds = None
if os.path.exists(TOKEN_FILE):
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        from google.auth.transport.requests import Request
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    
    with open(TOKEN_FILE, 'w') as token:
        token.write(creds.to_json())

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

# Собираем проблемные строки
problems = []
total_checked = 0
total_problems = 0

print("\n🔍 Сканирование таблиц...\n")

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
                # Читаем данные
                # B=карта, C=тел, D=дата откр, E=статус, F=служба, G=коммент, H=оператор фикс, I=дата фикс
                result = service.spreadsheets().values().get(
                    spreadsheetId=op['spreadsheet_id'],
                    range=f"'{title}'!B2:I20000",
                    valueRenderOption='FORMATTED_VALUE'
                ).execute()
                
                rows = result.get('values', [])
                
                for row_idx, row in enumerate(rows, 2):  # Начинаем с 2 (строка 1 - заголовок)
                    total_checked += 1
                    
                    # Проверяем на проблемы
                    issues = []
                    
                    # Проверка 1: Неполная строка (меньше 8 колонок)
                    if len(row) < 8:
                        issues.append(f"Неполная строка ({len(row)} колонок из 8)")
                    
                    # Если строка полная, проверяем содержимое
                    if len(row) >= 8:
                        # Проверка 2: Пустая дата фиксации (колонка I, индекс 7)
                        if not row[7] or str(row[7]).strip() == '':
                            issues.append("Пустая дата фиксации")
                        
                        # Проверка 3: Пустой оператор фиксировавший (колонка H, индекс 6)
                        if not row[6] or str(row[6]).strip() == '':
                            issues.append("Пустой оператор фиксировавший")
                        
                        # Проверка 4: Пустой номер карты (колонка B, индекс 0)
                        if not row[0] or str(row[0]).strip() == '':
                            issues.append("Пустой номер карты")
                        
                        # Проверка 5: Пустая дата открытия (колонка D, индекс 2)
                        if len(row) > 2 and (not row[2] or str(row[2]).strip() == ''):
                            issues.append("Пустая дата открытия")
                    
                    # Если есть проблемы - добавляем
                    if issues:
                        total_problems += 1
                        problems.append({
                            'operator': op['name'],
                            'spreadsheet_id': op['spreadsheet_id'],
                            'sheet': title,
                            'row': row_idx,
                            'issues': issues,
                            'data': row[:8] if len(row) >= 8 else row
                        })
                
            except Exception as e:
                print(f"  ⚠️ Ошибка на листе {title}: {e}")
                continue
        
    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        continue

# Сохраняем результаты
print("\n" + "=" * 100)
print(f"📊 РЕЗУЛЬТАТЫ СКАНИРОВАНИЯ")
print("=" * 100)
print(f"Всего проверено строк: {total_checked:,}")
print(f"Найдено проблемных строк: {total_problems:,}")

if total_problems > 0:
    print(f"\n💾 Сохранение списка проблемных строк...")
    
    with open('ПРОБЛЕМНЫЕ_СТРОКИ.txt', 'w', encoding='utf-8') as f:
        f.write("СПИСОК ПРОБЛЕМНЫХ СТРОК В GOOGLE SHEETS\n")
        f.write("=" * 100 + "\n\n")
        f.write(f"Всего проблемных строк: {total_problems:,}\n\n")
        
        current_operator = None
        for problem in problems:
            if problem['operator'] != current_operator:
                current_operator = problem['operator']
                f.write("\n" + "=" * 100 + "\n")
                f.write(f"ОПЕРАТОР: {current_operator}\n")
                f.write(f"Таблица ID: {problem['spreadsheet_id']}\n")
                f.write("=" * 100 + "\n\n")
            
            f.write(f"Лист: {problem['sheet']}, Строка: {problem['row']}\n")
            f.write(f"Проблемы: {', '.join(problem['issues'])}\n")
            f.write(f"Данные: {problem['data']}\n")
            f.write("-" * 100 + "\n")
    
    print(f"✅ Список сохранен: ПРОБЛЕМНЫЕ_СТРОКИ.txt")
    
    # Статистика по типам проблем
    print("\n📊 Статистика по типам проблем:")
    issue_counts = {}
    for problem in problems:
        for issue in problem['issues']:
            issue_counts[issue] = issue_counts.get(issue, 0) + 1
    
    for issue, count in sorted(issue_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"  {issue}: {count:,} строк ({count/total_problems*100:.1f}%)")
    
    # ТОП-10 операторов с проблемами
    print("\n👥 ТОП-10 операторов с проблемами:")
    operator_counts = {}
    for problem in problems:
        op = problem['operator']
        operator_counts[op] = operator_counts.get(op, 0) + 1
    
    for i, (op, count) in enumerate(sorted(operator_counts.items(), key=lambda x: x[1], reverse=True)[:10], 1):
        print(f"  {i}. {op}: {count:,} проблемных строк")

else:
    print("\n✅ Проблемных строк не найдено!")

print("\n" + "=" * 100)
print("✅ СКАНИРОВАНИЕ ЗАВЕРШЕНО")
print("=" * 100)
