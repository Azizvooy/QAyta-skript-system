"""
=============================================================================
МАССОВЫЙ ЭКСПОРТ GOOGLE SHEETS → CSV
=============================================================================
Скачивает все листы из таблиц операторов в CSV формат
и затем обрабатывает их локально
=============================================================================
"""

import os
import io
import time
from datetime import datetime
from collections import defaultdict
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from tqdm import tqdm

# Настройка прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

# Области доступа
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

# Настройки
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SETTINGS_SHEET_NAME = "Настройки"
SKIP_SHEETS = ["Статистика", "Предыдущий месяц", "Сводка по дням", "Настройки", 
               "FIKSA", "_FIKSA_STATE", "Аризалар", "GAI", "SETTING", "GRAFIK"]

TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'
EXPORT_FOLDER = 'exported_sheets'

# =============================================================================
# АУТЕНТИФИКАЦИЯ
# =============================================================================

def authenticate():
    """Аутентификация в Google API"""
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
    """Читает список операторов"""
    print(f"\n📋 Чтение списка операторов...")
    
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=f"{SETTINGS_SHEET_NAME}!A2:C100"  # Увеличил до 100 строк
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
        
        print(f"✅ Найдено ВСЕХ операторов: {len(operators)}")
        return operators  # Возвращаем ВСЕ, не только активных
        
    except HttpError as error:
        print(f"❌ Ошибка: {error}")
        return []

# =============================================================================
# ПОЛУЧЕНИЕ СПИСКА ЛИСТОВ
# =============================================================================

def get_sheet_gids(sheets_service, spreadsheet_id):
    """Получает список листов с их GID"""
    try:
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        sheet_list = []
        for sheet in sheets:
            props = sheet.get('properties', {})
            title = props.get('title', '')
            gid = props.get('sheetId', 0)
            
            # Пропускаем служебные листы
            if title not in SKIP_SHEETS:
                sheet_list.append({
                    'title': title,
                    'gid': gid
                })
        
        time.sleep(1)  # Задержка
        return sheet_list
        
    except HttpError as error:
        print(f"  ⚠️  Ошибка получения листов: {error}")
        return []

# =============================================================================
# ЭКСПОРТ ЛИСТА В CSV
# =============================================================================

def export_sheet_to_csv(spreadsheet_id, gid, output_path):
    """
    Экспортирует лист в CSV через прямую ссылку
    Это НЕ требует Drive API и работает быстрее
    """
    try:
        import requests
        
        # Получаем токен доступа
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        
        # URL для экспорта конкретного листа в CSV
        export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
        
        headers = {
            'Authorization': f'Bearer {creds.token}'
        }
        
        proxies = {
            'http': 'http://10.145.62.76:3128',
            'https': 'http://10.145.62.76:3128',
        }
        
        response = requests.get(export_url, headers=headers, proxies=proxies, timeout=60)
        
        if response.status_code == 200:
            with open(output_path, 'wb') as f:
                f.write(response.content)
            return True
        else:
            print(f"    ⚠️  Ошибка экспорта: HTTP {response.status_code}")
            return False
            
    except Exception as e:
        print(f"    ⚠️  Ошибка: {e}")
        return False

# =============================================================================
# ОСНОВНОЙ ПРОЦЕСС
# =============================================================================

def main():
    print("="*80)
    print("ЭКСПОРТ GOOGLE SHEETS → CSV")
    print("="*80)
    
    # Создаем папку для экспорта
    os.makedirs(EXPORT_FOLDER, exist_ok=True)
    
    # Аутентификация
    creds = authenticate()
    if not creds:
        return
    
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    # Получаем список операторов
    operators = get_operator_list(sheets_service)
    if not operators:
        print("❌ Не найдено операторов")
        return
    
    print(f"\n🚀 Начало экспорта: {len(operators)} операторов")
    
    exported_files = []
    
    # Обрабатываем каждого оператора
    for idx, operator in enumerate(operators, 1):
        operator_name = operator['name']
        spreadsheet_id = operator['spreadsheet_id']
        
        print(f"\n[{idx}/{len(operators)}] {operator_name}")
        
        # Создаем папку для оператора
        operator_folder = os.path.join(EXPORT_FOLDER, operator_name.replace('/', '_'))
        os.makedirs(operator_folder, exist_ok=True)
        
        # Получаем список листов
        sheets = get_sheet_gids(sheets_service, spreadsheet_id)
        print(f"  Листов для экспорта: {len(sheets)}")
        
        # Экспортируем каждый лист
        for sheet in sheets:
            title = sheet['title']
            gid = sheet['gid']
            
            # Безопасное имя файла
            safe_title = title.replace('/', '_').replace('\\', '_').replace(':', '_')
            csv_path = os.path.join(operator_folder, f"{safe_title}.csv")
            
            print(f"    📄 {title}...", end=" ", flush=True)
            
            if export_sheet_to_csv(spreadsheet_id, gid, csv_path):
                print(f"✅ ({os.path.getsize(csv_path) / 1024:.1f} KB)")
                exported_files.append({
                    'operator': operator_name,
                    'sheet': title,
                    'path': csv_path
                })
                time.sleep(1.5)  # Задержка между запросами
            else:
                print("❌")
    
    print(f"\n✅ Экспорт завершен!")
    print(f"📁 Экспортировано файлов: {len(exported_files)}")
    print(f"📂 Папка: {os.path.abspath(EXPORT_FOLDER)}")
    
    # Теперь обрабатываем все CSV файлы
    print("\n" + "="*80)
    print("ОБРАБОТКА CSV ФАЙЛОВ")
    print("="*80)
    
    all_data = []
    
    for file_info in tqdm(exported_files, desc="Обработка файлов"):
        try:
            df = pd.read_csv(file_info['path'], encoding='utf-8-sig')
            
            # Добавляем информацию об операторе и листе
            df['Оператор'] = file_info['operator']
            df['Архивный лист'] = file_info['sheet']
            
            all_data.append(df)
            
        except Exception as e:
            print(f"⚠️  Ошибка обработки {file_info['path']}: {e}")
    
    if all_data:
        # Объединяем все данные
        print("\n🔄 Объединение данных...")
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Сохраняем
        output_file = "ALL_DATA_COLLECTED.csv"
        combined_df.to_csv(output_file, index=False, encoding='utf-8-sig')
        
        print(f"\n✅ ГОТОВО!")
        print(f"📊 Всего записей: {len(combined_df):,}")
        print(f"📁 Файл: {output_file}")
        print(f"💾 Размер: {os.path.getsize(output_file) / (1024*1024):.2f} МБ")
        
        # Статистика
        print(f"\n📈 СТАТИСТИКА:")
        print(f"  Уникальных карт: {combined_df['Номер карты'].nunique():,}")
        print(f"  Операторов: {combined_df['Оператор'].nunique()}")
        
        if 'Статус' in combined_df.columns:
            print(f"\n  Топ-5 статусов:")
            for status, count in combined_df['Статус'].value_counts().head(5).items():
                print(f"    {status}: {count:,}")

if __name__ == "__main__":
    main()
