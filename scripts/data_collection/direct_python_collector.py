"""
=============================================================================
ПРЯМОЙ СБОРЩИК ДАННЫХ - PYTHON → GOOGLE SHEETS API
=============================================================================
Версия: 2.0 (для больших объемов)
Дата: 01.12.2025

📋 НАЗНАЧЕНИЕ:
Напрямую читает данные из таблиц операторов через Google Sheets API
и обрабатывает их в Python. БЕЗ промежуточного Google Docs.

💪 ВОЗМОЖНОСТИ:
- Обработка миллионов записей (35 операторов × 5 листов × 15k строк)
- Пакетная обработка (batch processing)
- Прогресс-бар в реальном времени
- Сохранение промежуточных результатов
- Возобновление после прерывания

📦 УСТАНОВКА:
pip install google-auth google-auth-oauthlib google-auth-httplib2
pip install google-api-python-client pandas tqdm

🚀 ИСПОЛЬЗОВАНИЕ:
python direct_python_collector.py
=============================================================================
"""

import json
import os
import pickle
from datetime import datetime
from collections import defaultdict
from typing import List, Dict, Any, Optional
import time

# Настройка прокси для корпоративной сети
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import socket
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

# Опциональные библиотеки
try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    print("⚠️  Установите tqdm для прогресс-бара: pip install tqdm")

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("⚠️  Установите pandas для расширенной аналитики: pip install pandas")

# =============================================================================
# НАСТРОЙКИ
# =============================================================================

# Области доступа
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ID таблицы с настройками (список операторов)
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

# ID таблицы для результатов
OUTPUT_SPREADSHEET_ID = "1Tu7LXTDZ2G_DxPiWJ3CLtORA6Uhma7yMe6vXAMgZDXA"

# Лист с настройками операторов
SETTINGS_SHEET_NAME = "Настройки"

# Служебные листы (пропускать при сборе архивов)
SKIP_SHEETS = ["Статистика", "Предыдущий месяц", "Сводка по дням", "Настройки"]

# Файлы
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'
CACHE_FILE = 'collection_cache.pkl'

# Настройки обработки
BATCH_SIZE = 100  # Обрабатываем по 100 строк за раз из одного листа
MAX_WORKERS = 5   # Параллельных запросов к API
RETRY_ATTEMPTS = 3  # Повторные попытки при ошибках
RETRY_DELAY = 2   # Задержка между повторами (секунды)

# Увеличенные таймауты для медленного соединения
socket.setdefaulttimeout(120)  # 2 минуты таймаут для всех сокетов

# =============================================================================
# АУТЕНТИФИКАЦИЯ
# =============================================================================

def authenticate():
    """Аутентификация в Google API с поддержкой прокси"""
    creds = None
    
    # Создаем requests session с прокси
    session = requests.Session()
    session.proxies = {
        'http': 'http://10.145.62.76:3128',
        'https': 'http://10.145.62.76:3128',
    }
    
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Обновление токена...")
            # Используем requests для обновления токена
            from google.auth.transport.requests import Request as GoogleRequest
            import google.auth.transport.requests
            
            # Патчим Request для использования нашей сессии с прокси
            original_session = google.auth.transport.requests.AuthorizedSession
            request = GoogleRequest(session=session)
            creds.refresh(request)
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print("❌ Файл credentials.json не найден!")
                print("\n📋 Инструкция:")
                print("1. https://console.cloud.google.com")
                print("2. Включите Google Sheets API")
                print("3. Создайте OAuth 2.0 credentials (Desktop)")
                print("4. Скачайте JSON → credentials.json")
                return None
            
            print("🔐 Авторизация (откроется браузер)...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print("✅ Токен сохранен")
    
    return creds

# =============================================================================
# ЧТЕНИЕ СПИСКА ОПЕРАТОРОВ
# =============================================================================

def get_operator_list(service) -> List[Dict[str, str]]:
    """
    Читает список операторов из листа Настройки
    
    Returns:
        List[Dict]: [{"name": "ФИО", "spreadsheet_id": "ID", "status": "активен"}]
    """
    print(f"\n📋 Чтение списка операторов из {MASTER_SPREADSHEET_ID}...")
    
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=f"{SETTINGS_SHEET_NAME}!A2:C1000"
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
        
        active_operators = [op for op in operators if op["status"].lower() == "активен"]
        
        print(f"✅ Найдено операторов: {len(operators)} (активных: {len(active_operators)})")
        return active_operators
        
    except HttpError as error:
        print(f"❌ Ошибка чтения списка операторов: {error}")
        return []

# =============================================================================
# ЧТЕНИЕ ДАННЫХ ИЗ ТАБЛИЦЫ ОПЕРАТОРА
# =============================================================================

def get_sheet_list(service, spreadsheet_id: str) -> List[str]:
    """Получает список листов в таблице"""
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheets = spreadsheet.get('sheets', [])
            
            sheet_names = []
            for sheet in sheets:
                title = sheet.get('properties', {}).get('title', '')
                if title not in SKIP_SHEETS:
                    sheet_names.append(title)
            
            # Задержка после успешного запроса
            time.sleep(1.5)
            
            return sheet_names
            
        except HttpError as error:
            if error.resp.status == 429:  # Rate limit exceeded
                wait_time = 20 * (attempt + 1)
                print(f"  ⏳ Лимит запросов. Ожидание {wait_time} сек...")
                time.sleep(wait_time)
            else:
                print(f"  ⚠️  Ошибка получения листов: {error}")
                return []
    
    return []

def read_sheet_data(service, spreadsheet_id: str, sheet_name: str, 
                   start_row: int = 2, batch_size: int = BATCH_SIZE) -> List[List]:
    """
    Читает данные с листа пакетами
    
    Args:
        service: Google Sheets API service
        spreadsheet_id: ID таблицы
        sheet_name: Название листа
        start_row: Начальная строка (2 = пропустить заголовок)
        batch_size: Размер пакета
        
    Returns:
        List[List]: Данные (строки)
    """
    max_retries = 3
    retry_delay = 2  # секунды
    
    for attempt in range(max_retries):
        try:
            # Читаем колонки B-I (номер карты, статус, дата и т.д.)
            range_name = f"'{sheet_name}'!B{start_row}:I{start_row + batch_size - 1}"
            
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            # Задержка после каждого успешного запроса (1.5 сек = ~40 запросов/мин)
            time.sleep(1.5)
            
            return result.get('values', [])
            
        except HttpError as error:
            if error.resp.status == 429:  # Rate limit exceeded
                wait_time = retry_delay * (attempt + 1) * 10  # 20, 40, 60 секунд
                print(f"\n    ⏳ Лимит запросов. Ожидание {wait_time} сек...", end="", flush=True)
                time.sleep(wait_time)
                print(" продолжаем")
            else:
                print(f"    ⚠️  Ошибка чтения {sheet_name}: {error}")
                return []
    
    print(f"    ❌ Не удалось прочитать {sheet_name} после {max_retries} попыток")
    return []

def process_operator(service, operator: Dict[str, str], progress_callback=None) -> Dict[str, Any]:
    """
    Обрабатывает данные одного оператора
    
    Args:
        service: Google Sheets API service
        operator: Данные оператора
        progress_callback: Функция для обновления прогресса
        
    Returns:
        Dict: Статистика оператора
    """
    operator_name = operator["name"]
    spreadsheet_id = operator["spreadsheet_id"]
    
    print(f"\n▶ Обработка: {operator_name}")
    
    # Получаем список листов
    sheets = get_sheet_list(service, spreadsheet_id)
    if not sheets:
        print(f"  ⚠️  Нет листов для обработки")
        return None
    
    print(f"  Найдено листов: {len(sheets)}")
    
    # Статистика
    stats = {
        "operator": operator_name,
        "total_records": 0,
        "unique_cards": set(),
        "by_sheet": {},
        "by_status": defaultdict(int),
        "by_month": defaultdict(int),
        "by_date": defaultdict(int)
    }
    
    # Обрабатываем каждый лист
    for sheet_name in sheets:
        print(f"    📄 {sheet_name}...", end=" ", flush=True)
        
        sheet_records = 0
        start_row = 2
        
        # Читаем лист пакетами
        while True:
            rows = read_sheet_data(service, spreadsheet_id, sheet_name, start_row, BATCH_SIZE)
            
            if not rows:
                break
            
            # Обрабатываем пакет
            for row in rows:
                if len(row) < 1:
                    continue
                
                card_num = row[0].strip() if len(row) > 0 else ""
                status = row[3].strip() if len(row) > 3 else ""  # Колонка E (индекс 3)
                date_value = row[7] if len(row) > 7 else ""      # Колонка I (индекс 7)
                
                if not card_num:
                    continue
                
                # Парсим дату
                date_str = parse_date(date_value)
                if not date_str:
                    continue
                
                # Обновляем статистику
                stats["total_records"] += 1
                stats["unique_cards"].add(card_num)
                stats["by_status"][status] += 1
                
                # Группировка по месяцам и датам
                try:
                    dt = datetime.strptime(date_str, '%d.%m.%Y %H:%M:%S')
                    month_key = dt.strftime('%m.%Y')
                    date_key = dt.strftime('%d.%m.%Y')
                    
                    stats["by_month"][month_key] += 1
                    stats["by_date"][date_key] += 1
                except:
                    pass
                
                sheet_records += 1
            
            # Если пакет неполный - достигли конца листа
            if len(rows) < BATCH_SIZE:
                break
            
            start_row += BATCH_SIZE
            
            # Обновляем прогресс
            if progress_callback:
                progress_callback(sheet_records)
        
        stats["by_sheet"][sheet_name] = sheet_records
        print(f"{sheet_records} записей")
    
    print(f"  ✅ Всего записей: {stats['total_records']}, уникальных карт: {len(stats['unique_cards'])}")
    
    return stats

# =============================================================================
# ПАРСИНГ ДАТЫ
# =============================================================================

def parse_date(value) -> Optional[str]:
    """Парсит дату из различных форматов"""
    if not value:
        return None
    
    # Если уже строка с датой
    if isinstance(value, str):
        # Формат: "01.12.2024 10:30:45"
        if value.count('.') == 2 and value.count(':') == 2:
            return value
        # Формат: "01.12.2024"
        elif value.count('.') == 2:
            return value + " 00:00:00"
    
    # Если объект даты (serial number из Excel)
    try:
        # Google Sheets serial date (days since 30.12.1899)
        if isinstance(value, (int, float)) and value > 0:
            base_date = datetime(1899, 12, 30)
            date_obj = base_date + timedelta(days=value)
            return date_obj.strftime('%d.%m.%Y 00:00:00')
    except:
        pass
    
    return None

# =============================================================================
# СОХРАНЕНИЕ РЕЗУЛЬТАТОВ
# =============================================================================

def save_results_to_sheets(service, all_stats: List[Dict[str, Any]]):
    """
    Сохраняет результаты в Google Sheets
    
    Args:
        service: Google Sheets API service
        all_stats: Список статистик всех операторов
    """
    print(f"\n📝 Запись результатов в {OUTPUT_SPREADSHEET_ID}...")
    
    try:
        # Лист 1: Сводная статистика по операторам
        operator_data = [
            ['ФИО оператора', 'Всего записей', 'Уникальных карт', 'Листов обработано', 'Статусы (топ-3)']
        ]
        
        for stats in sorted(all_stats, key=lambda x: x['total_records'], reverse=True):
            if stats:
                top_statuses = sorted(stats['by_status'].items(), key=lambda x: x[1], reverse=True)[:3]
                status_str = ', '.join([f"{k}: {v}" for k, v in top_statuses])
                
                operator_data.append([
                    stats['operator'],
                    stats['total_records'],
                    len(stats['unique_cards']),
                    len(stats['by_sheet']),
                    status_str
                ])
        
        # Записываем
        service.spreadsheets().values().update(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            range='Сводка по операторам!A1',
            valueInputOption='RAW',
            body={'values': operator_data}
        ).execute()
        
        print(f"  ✅ Записано операторов: {len(operator_data) - 1}")
        
        # Лист 2: Статистика по месяцам
        monthly_data = [['Месяц', 'Всего записей', 'Операторов', 'Уникальных карт']]
        
        monthly_totals = defaultdict(lambda: {'records': 0, 'operators': set(), 'cards': set()})
        
        for stats in all_stats:
            if stats:
                for month, count in stats['by_month'].items():
                    monthly_totals[month]['records'] += count
                    monthly_totals[month]['operators'].add(stats['operator'])
                    monthly_totals[month]['cards'].update(stats['unique_cards'])
        
        for month in sorted(monthly_totals.keys(), reverse=True):
            data = monthly_totals[month]
            monthly_data.append([
                month,
                data['records'],
                len(data['operators']),
                len(data['cards'])
            ])
        
        service.spreadsheets().values().update(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            range='Статистика по месяцам!A1',
            valueInputOption='RAW',
            body={'values': monthly_data}
        ).execute()
        
        print(f"  ✅ Записано месяцев: {len(monthly_data) - 1}")
        
        # Лист 3: Общая информация
        total_records = sum(s['total_records'] for s in all_stats if s)
        total_cards = len(set().union(*[s['unique_cards'] for s in all_stats if s]))
        
        summary_data = [
            ['Параметр', 'Значение'],
            ['Дата обработки', datetime.now().strftime('%d.%m.%Y %H:%M:%S')],
            ['Обработано операторов', len([s for s in all_stats if s])],
            ['Всего записей', total_records],
            ['Уникальных карт', total_cards],
            ['Средняя нагрузка', int(total_records / len([s for s in all_stats if s])) if all_stats else 0]
        ]
        
        service.spreadsheets().values().update(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            range='Общая информация!A1',
            valueInputOption='RAW',
            body={'values': summary_data}
        ).execute()
        
        print("  ✅ Записана общая информация")
        
    except HttpError as error:
        print(f"  ❌ Ошибка записи: {error}")

# =============================================================================
# КЭШИРОВАНИЕ (для возобновления после прерывания)
# =============================================================================

def save_cache(all_stats: List[Dict[str, Any]], processed_operators: List[str]):
    """Сохраняет промежуточные результаты"""
    cache = {
        'stats': all_stats,
        'processed': processed_operators,
        'timestamp': datetime.now().isoformat()
    }
    with open(CACHE_FILE, 'wb') as f:
        pickle.dump(cache, f)

def load_cache() -> Optional[Dict]:
    """Загружает кэш"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'rb') as f:
                return pickle.load(f)
        except:
            pass
    return None

def clear_cache():
    """Удаляет кэш"""
    if os.path.exists(CACHE_FILE):
        os.remove(CACHE_FILE)

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    """Основная функция сборки и обработки"""
    print("=" * 80)
    print("ПРЯМОЙ СБОРЩИК ДАННЫХ - PYTHON → GOOGLE SHEETS API")
    print("=" * 80)
    
    # Проверка конфигурации
    if MASTER_SPREADSHEET_ID == "ВСТАВЬТЕ_ID_ГЛАВНОЙ_ТАБЛИЦЫ":
        print("\n❌ Необходимо настроить MASTER_SPREADSHEET_ID в коде")
        return
    
    if OUTPUT_SPREADSHEET_ID == "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_РЕЗУЛЬТАТОВ":
        print("\n❌ Необходимо настроить OUTPUT_SPREADSHEET_ID в коде")
        return
    
    # Аутентификация
    creds = authenticate()
    if not creds:
        return
    
    # Создаем сервис (прокси уже настроен через переменные окружения)
    try:
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        print("✅ Google Sheets API подключен")
        
        # Тестируем подключение
        print("🔍 Проверка доступа к таблице...")
        test = service.spreadsheets().get(spreadsheetId=MASTER_SPREADSHEET_ID).execute()
        print(f"✅ Доступ получен: {test.get('properties', {}).get('title', 'Без имени')}")
        
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        print("\n💡 Возможные причины:")
        print("   1. Проблема с интернет-соединением через прокси")
        print("   2. Прокси требует авторизацию")
        print("   3. Прокси блокирует googleapis.com")
        print("\n🔧 Попробуйте:")
        print("   - Использовать мобильный интернет")
        print("   - Запустить на домашнем компьютере")
        return
    
    # Проверяем кэш
    cache = load_cache()
    if cache:
        print(f"\n💾 Найден кэш от {cache['timestamp']}")
        print(f"   Обработано: {len(cache['processed'])} операторов")
        response = input("Продолжить с места прерывания? (y/n): ")
        if response.lower() == 'y':
            all_stats = cache['stats']
            processed_operators = set(cache['processed'])
        else:
            all_stats = []
            processed_operators = set()
            clear_cache()
    else:
        all_stats = []
        processed_operators = set()
    
    # Получаем список операторов
    operators = get_operator_list(service)
    if not operators:
        print("\n⚠️  Нет операторов для обработки")
        return
    
    # Фильтруем необработанных
    operators_to_process = [op for op in operators if op['name'] not in processed_operators]
    
    if not operators_to_process:
        print("\n✅ Все операторы уже обработаны")
    else:
        print(f"\n🚀 Начало обработки: {len(operators_to_process)} операторов")
        print(f"   (уже обработано: {len(processed_operators)})")
        
        start_time = time.time()
        
        # Обрабатываем операторов
        iterator = tqdm(operators_to_process, desc="Операторы") if HAS_TQDM else operators_to_process
        
        for operator in iterator:
            try:
                stats = process_operator(service, operator)
                if stats:
                    all_stats.append(stats)
                    processed_operators.add(operator['name'])
                    
                    # Сохраняем кэш каждые 5 операторов
                    if len(processed_operators) % 5 == 0:
                        save_cache(all_stats, list(processed_operators))
                
            except KeyboardInterrupt:
                print("\n\n⚠️  Прерывание пользователем")
                save_cache(all_stats, list(processed_operators))
                print("💾 Прогресс сохранен. Запустите снова для продолжения.")
                return
            
            except Exception as e:
                print(f"\n❌ Ошибка обработки {operator['name']}: {e}")
                continue
        
        duration = time.time() - start_time
        
        print(f"\n⏱️  Время обработки: {int(duration)} сек ({int(duration/60)} мин)")
    
    # Сохраняем результаты
    if all_stats:
        save_results_to_sheets(service, all_stats)
        
        # Статистика
        total_records = sum(s['total_records'] for s in all_stats)
        total_cards = len(set().union(*[s['unique_cards'] for s in all_stats]))
        
        print("\n" + "=" * 80)
        print("✅ ОБРАБОТКА ЗАВЕРШЕНА")
        print(f"   Операторов: {len(all_stats)}")
        print(f"   Записей: {total_records:,}")
        print(f"   Уникальных карт: {total_cards:,}")
        print("=" * 80)
        
        # Очищаем кэш после успешного завершения
        clear_cache()
    else:
        print("\n⚠️  Нет данных для сохранения")

if __name__ == '__main__':
    from datetime import timedelta
    main()
