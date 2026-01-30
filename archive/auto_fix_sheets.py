"""
Автоматическое исправление проблемных строк в Google Sheets:
1. Удаление полностью пустых строк (где все ячейки B-I пусты)
2. Заполнение пустых "Оператор фиксировавший" из колонки "Оператор"
"""
import os

# Настройка прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import socket
import time
socket.setdefaulttimeout(120)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = 'token.json'
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"
SKIP_SHEETS = ["Статистика", "Предыдущий месяц", "Сводка по дням", "Настройки", "setting", "аризалар"]

print("=" * 100)
print("АВТОМАТИЧЕСКОЕ ИСПРАВЛЕНИЕ ПРОБЛЕМНЫХ СТРОК")
print("=" * 100)

# Подключение
print("\n🔗 Подключение к Google Sheets...")
import os.path

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

# Статистика
total_deleted = 0
total_fixed = 0
total_errors = 0

print("\n🔧 Исправление данных...\n")

for idx, op in enumerate(operators, 1):
    print(f"[{idx}/{len(operators)}] {op['name']}")
    
    try:
        # Получаем листы
        metadata = service.spreadsheets().get(spreadsheetId=op['spreadsheet_id']).execute()
        sheets = metadata.get('sheets', [])
        
        for sheet in sheets:
            title = sheet['properties']['title']
            sheet_id = sheet['properties']['sheetId']
            
            # Пропускаем служебные
            if title in SKIP_SHEETS:
                continue
            
            title_lower = title.lower()
            skip_words = ['setting', 'аризалар', 'настройки', 'сводка', 'статистика']
            if any(word in title_lower for word in skip_words):
                continue
            
            try:
                # Читаем данные включая колонку A (для проверки формул)
                result = service.spreadsheets().values().get(
                    spreadsheetId=op['spreadsheet_id'],
                    range=f"'{title}'!A2:I20000",
                    valueRenderOption='FORMATTED_VALUE'
                ).execute()
                
                rows = result.get('values', [])
                if not rows:
                    continue
                
                # Собираем изменения
                updates = []
                rows_to_delete = []
                
                for row_idx, row in enumerate(rows, 2):  # Начинаем с 2 (строка 1 - заголовок)
                    # Проверяем колонки B-I (индексы 1-8 в массиве, т.к. A=0)
                    if len(row) <= 1:
                        # Строка пустая или только колонка A
                        rows_to_delete.append(row_idx)
                        continue
                    
                    # Проверяем все ли ячейки B-I пусты
                    data_cols = row[1:9] if len(row) > 1 else []
                    all_empty = all(not str(cell).strip() for cell in data_cols)
                    
                    if all_empty:
                        # Все колонки B-I пусты - удаляем
                        rows_to_delete.append(row_idx)
                        continue
                    
                    # Если строка не пустая, проверяем "Оператор фиксировавший" (колонка H, индекс 7)
                    if len(row) >= 8:
                        operator_fix = row[7] if len(row) > 7 else ''
                        
                        if not operator_fix or not str(operator_fix).strip():
                            # Берем значение из колонки A (оператор)
                            operator_name = op['name']
                            
                            # Добавляем обновление
                            updates.append({
                                'range': f"'{title}'!H{row_idx}",
                                'values': [[operator_name]]
                            })
                
                # Применяем обновления
                if updates:
                    batch_data = {
                        'valueInputOption': 'RAW',
                        'data': updates
                    }
                    service.spreadsheets().values().batchUpdate(
                        spreadsheetId=op['spreadsheet_id'],
                        body=batch_data
                    ).execute()
                    
                    total_fixed += len(updates)
                    print(f"  ✓ Заполнено пустых операторов: {len(updates)}")
                
                # Удаляем пустые строки (с конца, чтобы индексы не сбивались)
                if rows_to_delete:
                    rows_to_delete.sort(reverse=True)
                    
                    # Группируем последовательные строки для массового удаления
                    delete_requests = []
                    
                    i = 0
                    while i < len(rows_to_delete):
                        start_row = rows_to_delete[i]
                        end_row = start_row
                        
                        # Ищем последовательные строки
                        j = i + 1
                        while j < len(rows_to_delete) and rows_to_delete[j] == rows_to_delete[j-1] - 1:
                            end_row = rows_to_delete[j]
                            j += 1
                        
                        # Добавляем запрос на удаление диапазона
                        delete_requests.append({
                            'deleteDimension': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'dimension': 'ROWS',
                                    'startIndex': end_row - 1,  # 0-based
                                    'endIndex': start_row  # exclusive
                                }
                            }
                        })
                        
                        i = j
                    
                    # Применяем удаление пачками (максимум 10 за раз)
                    for batch_start in range(0, len(delete_requests), 10):
                        batch = delete_requests[batch_start:batch_start + 10]
                        
                        service.spreadsheets().batchUpdate(
                            spreadsheetId=op['spreadsheet_id'],
                            body={'requests': batch}
                        ).execute()
                        
                        time.sleep(0.5)  # Небольшая пауза между запросами
                    
                    total_deleted += len(rows_to_delete)
                    print(f"  ✓ Удалено пустых строк: {len(rows_to_delete)}")
                
                # Небольшая пауза между листами
                time.sleep(0.3)
                
            except Exception as e:
                total_errors += 1
                print(f"  ⚠️ Ошибка на листе {title}: {e}")
                continue
        
    except Exception as e:
        total_errors += 1
        print(f"  ❌ Ошибка: {e}")
        continue

# Итоговая статистика
print("\n" + "=" * 100)
print("📊 ИТОГИ")
print("=" * 100)
print(f"Удалено пустых строк: {total_deleted:,}")
print(f"Заполнено пустых операторов: {total_fixed:,}")
print(f"Ошибок: {total_errors}")

print("\n" + "=" * 100)
print("✅ ИСПРАВЛЕНИЕ ЗАВЕРШЕНО")
print("=" * 100)
print("\n💡 Рекомендации:")
print("   1. Запустите сбор данных заново: python collect_to_excel.py")
print("   2. Создайте новые отчеты: python create_qayta_report.py")
print("=" * 100)
