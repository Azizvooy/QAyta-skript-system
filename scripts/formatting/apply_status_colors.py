"""
Скрипт для применения цветового форматирования к колонке E (Статусы) через Google Sheets API
Автор: GitHub Copilot
Дата: 07.01.2026
"""

import os
import sys
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

# Настройка прокси (если требуется)
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

# Область доступа для Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Цветовая схема для статусов (RGB в формате 0-1)
STATUS_COLORS = {
    "отрицательный": {
        "background": {"red": 1.0, "green": 0.0, "blue": 0.0},  # Красный
        "foreground": {"red": 1.0, "green": 1.0, "blue": 1.0}   # Белый текст
    },
    "положительный": {
        "background": {"red": 0.0, "green": 1.0, "blue": 0.0},  # Зеленый
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}   # Черный текст
    },
    "тишине": {
        "background": {"red": 1.0, "green": 0.8, "blue": 0.8},  # Нежно-красный
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}   # Черный текст
    },
    "соед прервано": {
        "background": {"red": 1.0, "green": 0.8, "blue": 0.8},  # Нежно-красный
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}   # Черный текст
    },
    "НЕТ ОТВЕТА (ЗАНЯТО)": {
        "background": {"red": 1.0, "green": 1.0, "blue": 0.0},  # Желтый
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}   # Черный текст
    },
    "заявка закрыта (не удалось дозвониться)": {
        "background": {"red": 0.8, "green": 0.8, "blue": 0.8},  # Серый
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}   # Черный текст
    },
    "открыть карту": {
        "background": {"red": 0.53, "green": 0.81, "blue": 0.92},  # Небесный (#87ceeb)
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}      # Черный текст
    },
    "тиббиёт ходими аризаси": {
        "background": {"red": 0.68, "green": 0.85, "blue": 0.90},  # Голубой нежный (#add8e6)
        "foreground": {"red": 0.0, "green": 0.0, "blue": 0.0}      # Черный текст
    }
}


def get_credentials():
    """Получает учетные данные для Google Sheets API"""
    creds = None
    
    # Проверяем наличие сохраненного токена
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # Если нет валидных учетных данных
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("❌ Файл credentials.json не найден!")
                print("Скачайте его из Google Cloud Console")
                sys.exit(1)
            
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Сохраняем учетные данные
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds


def create_conditional_format_rule(status_text, color_scheme, sheet_id, start_row=2, end_row=1000):
    """
    Создает правило условного форматирования для конкретного статуса
    
    Args:
        status_text: текст статуса для сравнения
        color_scheme: словарь с фоновым и текстовым цветами
        sheet_id: ID листа (не путать с ID таблицы!)
        start_row: начальная строка (по умолчанию 2)
        end_row: конечная строка (по умолчанию 1000)
    
    Returns:
        Словарь с правилом условного форматирования
    """
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": sheet_id,
                    "startRowIndex": start_row - 1,  # API использует 0-based индексы
                    "endRowIndex": end_row,
                    "startColumnIndex": 4,  # Колонка E (0-based)
                    "endColumnIndex": 5
                }],
                "booleanRule": {
                    "condition": {
                        "type": "TEXT_EQ",
                        "values": [{
                            "userEnteredValue": status_text
                        }]
                    },
                    "format": {
                        "backgroundColor": color_scheme["background"],
                        "textFormat": {
                            "foregroundColor": color_scheme["foreground"]
                        }
                    }
                }
            },
            "index": 0
        }
    }


def delete_existing_rules_for_column_e(service, spreadsheet_id, sheet_id):
    """
    Удаляет существующие правила условного форматирования для колонки E
    
    Args:
        service: объект Google Sheets API service
        spreadsheet_id: ID таблицы
        sheet_id: ID листа
    """
    try:
        # Получаем текущие правила
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        sheets = spreadsheet.get('sheets', [])
        target_sheet = None
        
        for sheet in sheets:
            if sheet['properties']['sheetId'] == sheet_id:
                target_sheet = sheet
                break
        
        if not target_sheet:
            return
        
        conditional_rules = target_sheet.get('conditionalFormats', [])
        
        # Собираем индексы правил для колонки E
        rules_to_delete = []
        for idx, rule in enumerate(conditional_rules):
            ranges = rule.get('ranges', [])
            for range_obj in ranges:
                # Проверяем, относится ли правило к колонке E (индекс 4)
                if range_obj.get('startColumnIndex') == 4 and range_obj.get('endColumnIndex') == 5:
                    rules_to_delete.append(idx)
                    break
        
        # Удаляем правила (в обратном порядке, чтобы индексы не сбивались)
        if rules_to_delete:
            requests = []
            for idx in sorted(rules_to_delete, reverse=True):
                requests.append({
                    "deleteConditionalFormatRule": {
                        "sheetId": sheet_id,
                        "index": idx
                    }
                })
            
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()
            
            print(f"  ✓ Удалено старых правил: {len(rules_to_delete)}")
    
    except Exception as e:
        print(f"  ⚠️  Ошибка при удалении старых правил: {e}")


def apply_formatting_to_sheet(service, spreadsheet_id, sheet_name="FIKSA"):
    """
    Применяет цветовое форматирование к листу FIKSA
    
    Args:
        service: объект Google Sheets API service
        spreadsheet_id: ID таблицы
        sheet_name: название листа (по умолчанию "FIKSA")
    """
    try:
        # Получаем информацию о таблице
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        # Находим ID листа FIKSA
        sheet_id = None
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break
        
        if sheet_id is None:
            print(f"  ⚠️  Лист '{sheet_name}' не найден")
            return False
        
        print(f"  📄 Лист '{sheet_name}' найден (ID: {sheet_id})")
        
        # Удаляем существующие правила для колонки E
        delete_existing_rules_for_column_e(service, spreadsheet_id, sheet_id)
        
        # Создаем новые правила условного форматирования
        requests = []
        
        for status_text, color_scheme in STATUS_COLORS.items():
            rule = create_conditional_format_rule(status_text, color_scheme, sheet_id)
            requests.append(rule)
        
        # Применяем все правила одним запросом
        body = {
            'requests': requests
        }
        
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        
        print(f"  ✓ Применено правил форматирования: {len(STATUS_COLORS)}")
        return True
        
    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        return False


def read_operators_list(service):
    """
    Читает список операторов из мастер-таблицы Google Sheets
    """
    MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
    SETTINGS_SHEET_NAME = "Настройки"
    
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=f"{SETTINGS_SHEET_NAME}!A2:C100"
        ).execute()
        
        values = result.get('values', [])
        operators = []
        
        for row in values:
            if len(row) >= 2:
                name = row[0].strip() if len(row) > 0 else ""
                spreadsheet_id = row[1].strip() if len(row) > 1 else ""
                status = row[2].strip() if len(row) > 2 else "активен"
                
                if name and spreadsheet_id and spreadsheet_id != "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ":
                    if status.lower() == "активен":
                        operators.append({
                            'name': name,
                            'id': spreadsheet_id
                        })
        
        return operators
        
    except Exception as e:
        print(f"⚠️  Ошибка чтения мастер-таблицы: {e}")
        return []


def main():
    """Основная функция"""
    print("=" * 70)
    print("🎨 ПРИМЕНЕНИЕ ЦВЕТОВОГО ФОРМАТИРОВАНИЯ К КОЛОНКЕ E (СТАТУСЫ)")
    print("=" * 70)
    print()
    
    # Получаем учетные данные
    print("🔐 Авторизация в Google Sheets API...")
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    print("✅ Авторизация успешна\n")
    
    # Читаем список операторов
    print("📋 Чтение списка операторов из мастер-таблицы...")
    operators = read_operators_list(service)
    
    if not operators:
        print("❌ Нет активных операторов в списке")
        print("\nВведите ID таблицы вручную (или нажмите Enter для выхода):")
        manual_id = input("ID таблицы: ").strip()
        
        if manual_id:
            operators = [{'name': 'Ручной ввод', 'id': manual_id}]
        else:
            return
    
    print(f"✅ Найдено операторов: {len(operators)}\n")
    
    # Применяем форматирование к каждой таблице
    success_count = 0
    fail_count = 0
    
    for idx, operator in enumerate(operators, 1):
        print(f"[{idx}/{len(operators)}] {operator['name']}")
        print(f"  🔗 ID: {operator['id']}")
        
        if apply_formatting_to_sheet(service, operator['id']):
            success_count += 1
        else:
            fail_count += 1
        
        print()
    
    # Итоги
    print("=" * 70)
    print("📊 РЕЗУЛЬТАТЫ:")
    print(f"  ✅ Успешно обработано: {success_count}")
    print(f"  ❌ Ошибок: {fail_count}")
    print("=" * 70)
    print()
    print("🎨 Цветовая схема:")
    print("  🔴 Отрицательный - красный")
    print("  🟢 Положительный - зеленый")
    print("  🩷 Тишине / Соед прервано - нежно-красный")
    print("  🟡 НЕТ ОТВЕТА (ЗАНЯТО) - желтый")
    print("  ⚪ Заявка закрыта - серый")
    print("  🔵 Открыть карту - небесный")
    print("  💙 Тиббиёт ходими аризаси - голубой нежный")
    print()


if __name__ == "__main__":
    main()
