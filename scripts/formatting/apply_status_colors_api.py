"""
Скрипт для применения цветового форматирования к колонке E (Статусы)
через Google Sheets API.

Применяет:
1. Выпадающий список с цветными вариантами
2. Условное форматирование с более светлой гаммой цветов
"""

import os
import json
import httplib2
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Отключаем проверку прокси для локального подключения
os.environ['NO_PROXY'] = 'localhost,127.0.0.1'
os.environ['no_proxy'] = 'localhost,127.0.0.1'

# Области доступа для API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ID мастер-таблицы со списком операторов
MASTER_SPREADSHEET_ID = "1wlqqSCV3HW5ZgfYUT6IS2Ne466jJQeEKH1Nl4Tx2jdc"

# Цветовая схема (более светлая гамма)
# RGB значения от 0 до 1
STATUS_COLORS = {
    "отрицательный": {
        "bg": {"red": 1.0, "green": 0.4, "blue": 0.4},      # Светло-красный
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "положительный": {
        "bg": {"red": 0.6, "green": 1.0, "blue": 0.6},      # Светло-зеленый
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "тишине": {
        "bg": {"red": 1.0, "green": 0.85, "blue": 0.85},    # Очень светло-розовый
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "соед прервано": {
        "bg": {"red": 1.0, "green": 0.85, "blue": 0.85},    # Очень светло-розовый
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "НЕТ ОТВЕТА (ЗАНЯТО)": {
        "bg": {"red": 1.0, "green": 1.0, "blue": 0.6},      # Светло-желтый
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "заявка закрыта (не удалось дозвониться)": {
        "bg": {"red": 0.85, "green": 0.85, "blue": 0.85},   # Светло-серый
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "открыть карту": {
        "bg": {"red": 0.6, "green": 0.85, "blue": 1.0},     # Светло-небесный
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    },
    "тиббиёт ходими аризаси": {
        "bg": {"red": 0.7, "green": 0.9, "blue": 1.0},      # Нежно-голубой
        "fg": {"red": 0.0, "green": 0.0, "blue": 0.0}       # Черный текст
    }
}

def get_credentials():
    """Получает или создает учетные данные для Google API"""
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Ошибка обновления токена: {e}")
                os.remove('token.json')
                creds = None
        
        if not creds:
            if not os.path.exists('credentials.json'):
                print("❌ Файл credentials.json не найден!")
                print("Создайте его в Google Cloud Console:")
                print("https://console.cloud.google.com/apis/credentials")
                return None
            
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def get_operator_list(service):
    """Получает список операторов из мастер-таблицы"""
    try:
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
                
                if name and spreadsheet_id and spreadsheet_id != "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
                    if status == "активен":
                        operators.append({
                            'name': name,
                            'spreadsheet_id': spreadsheet_id
                        })
        
        print(f"📋 Найдено активных операторов: {len(operators)}")
        return operators
        
    except HttpError as error:
        print(f"❌ Ошибка при получении списка операторов: {error}")
        return []

def create_conditional_format_rules():
    """Создает правила условного форматирования для всех статусов"""
    rules = []
    
    for status_text, colors in STATUS_COLORS.items():
        rule = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": 0,  # ID листа FIKSA (обычно 0)
                        "startColumnIndex": 4,  # Колонка E (0-indexed)
                        "endColumnIndex": 5,
                        "startRowIndex": 1  # Начиная со строки 2
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{
                                "userEnteredValue": status_text
                            }]
                        },
                        "format": {
                            "backgroundColor": colors["bg"],
                            "textFormat": {
                                "foregroundColor": colors["fg"]
                            }
                        }
                    }
                },
                "index": 0
            }
        }
        rules.append(rule)
    
    return rules

def create_data_validation_rule():
    """Создает правило выпадающего списка для колонки E"""
    status_list = list(STATUS_COLORS.keys())
    
    return {
        "setDataValidation": {
            "range": {
                "sheetId": 0,  # ID листа FIKSA
                "startColumnIndex": 4,  # Колонка E
                "endColumnIndex": 5,
                "startRowIndex": 1  # Начиная со строки 2
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": status} for status in status_list]
                },
                "showCustomUi": True,
                "strict": True
            }
        }
    }

def get_sheet_id(service, spreadsheet_id, sheet_name="FIKSA"):
    """Получает ID листа по его имени"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        for sheet in sheets:
            properties = sheet.get('properties', {})
            if properties.get('title') == sheet_name:
                return properties.get('sheetId')
        
        return 0  # По умолчанию первый лист
        
    except Exception as e:
        print(f"   ⚠️  Не удалось получить ID листа: {e}")
        return 0

def apply_formatting_to_spreadsheet(service, spreadsheet_id, operator_name):
    """Применяет форматирование к одной таблице оператора"""
    try:
        # Получаем ID листа FIKSA
        sheet_id = get_sheet_id(service, spreadsheet_id, "FIKSA")
        
        # Подготавливаем запросы
        requests = []
        
        # 1. Удаляем старые правила условного форматирования для колонки E
        requests.append({
            "deleteConditionalFormatRule": {
                "sheetId": sheet_id,
                "index": 0
            }
        })
        
        # 2. Создаем выпадающий список
        validation_rule = create_data_validation_rule()
        validation_rule["setDataValidation"]["range"]["sheetId"] = sheet_id
        requests.append(validation_rule)
        
        # 3. Добавляем правила условного форматирования
        format_rules = create_conditional_format_rules()
        for rule in format_rules:
            rule["addConditionalFormatRule"]["rule"]["ranges"][0]["sheetId"] = sheet_id
            requests.append(rule)
        
        # Применяем все изменения одним batch-запросом
        body = {"requests": requests}
        
        try:
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            print(f"   ✅ {operator_name}: форматирование применено ({len(requests)} операций)")
            return True
            
        except HttpError as error:
            error_message = str(error)
            
            # Если ошибка из-за попытки удалить несуществующее правило, повторяем без удаления
            if "Invalid requests[0].deleteConditionalFormatRule" in error_message:
                print(f"   ℹ️  {operator_name}: старых правил нет, создаем новые...")
                requests = requests[1:]  # Убираем запрос на удаление
                body = {"requests": requests}
                
                response = service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=body
                ).execute()
                
                print(f"   ✅ {operator_name}: форматирование применено ({len(requests)} операций)")
                return True
            else:
                raise error
        
    except HttpError as error:
        print(f"   ❌ {operator_name}: ошибка - {error}")
        return False
    except Exception as e:
        print(f"   ❌ {operator_name}: неожиданная ошибка - {e}")
        return False

def main():
    """Основная функция"""
    print("=" * 60)
    print("🎨 ПРИМЕНЕНИЕ ЦВЕТОВОГО ФОРМАТИРОВАНИЯ К КОЛОНКЕ E")
    print("=" * 60)
    print()
    
    # Получаем учетные данные
    print("🔐 Авторизация в Google API...")
    creds = get_credentials()
    
    if not creds:
        print("❌ Не удалось получить учетные данные!")
        return
    
    try:
        # Создаем сервис Google Sheets API
        service = build('sheets', 'v4', credentials=creds)
        
        # Получаем список операторов
        print("📋 Получение списка операторов...")
        operators = get_operator_list(service)
        
        if not operators:
            print("❌ Нет активных операторов для обработки!")
            return
        
        print()
        print(f"🚀 Начинаем обработку {len(operators)} таблиц...")
        print()
        
        # Применяем форматирование к каждой таблице
        success_count = 0
        failed_count = 0
        
        for i, operator in enumerate(operators, 1):
            print(f"[{i}/{len(operators)}] {operator['name']}...")
            
            if apply_formatting_to_spreadsheet(
                service, 
                operator['spreadsheet_id'], 
                operator['name']
            ):
                success_count += 1
            else:
                failed_count += 1
        
        print()
        print("=" * 60)
        print("📊 РЕЗУЛЬТАТЫ:")
        print(f"   ✅ Успешно обработано: {success_count}")
        print(f"   ❌ Ошибок: {failed_count}")
        print("=" * 60)
        
        if success_count > 0:
            print()
            print("🎨 Применены цвета (светлая гамма):")
            print("   • Отрицательный → Светло-красный")
            print("   • Положительный → Светло-зеленый")
            print("   • Тишине → Нежно-розовый")
            print("   • Соед прервано → Нежно-розовый")
            print("   • НЕТ ОТВЕТА (ЗАНЯТО) → Светло-желтый")
            print("   • Заявка закрыта → Светло-серый")
            print("   • Открыть карту → Светло-небесный")
            print("   • Тиббиёт ходими аризаси → Нежно-голубой")
        
    except HttpError as error:
        print(f"❌ Ошибка API: {error}")
    except Exception as e:
        print(f"❌ Неожиданная ошибка: {e}")

if __name__ == '__main__':
    main()
