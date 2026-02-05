#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ИМПОРТ И СОПОСТАВЛЕНИЕ ДАННЫХ ЗА ЯНВАРЬ 2026
=============================================================================
Импортирует данные из Google Sheets и сопоставляет с файлами 112 из папки 123
Учитывает:
- Номер карточки (Карточка рақами)
- Номер инцидента (Ҳодиса рақами)
- Логику жалоб и положительных отзывов
=============================================================================
"""

import os
import pandas as pd
import glob
from datetime import datetime
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import socket

# Прокси (закомментировано для работы в Codespaces)
# os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
# os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
# socket.setdefaulttimeout(120)

BASE_DIR = Path(__file__).parent
TOKEN_FILE = BASE_DIR / 'config' / 'token.json'
CREDENTIALS_FILE = BASE_DIR / 'config' / 'credentials.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Список ID документов для импорта
SPREADSHEET_IDS = [
    "18y_QSol_XIZiaKGdoc64-tqerxYXg1kwmO7mmxo21rQ",
    "1JG9RqC64MZbP63HCSa0avp55zXP8nEqzsiuQbbSi9Do",
    "1Ld37ljkNect6iLV0X0QztTJC7vhgPDkaRXTJGVYOtfg",
    "16AK6_7FcWeksg2KaLfuzxo6sOGxsORw9hU-lycgeGiM",
    "1h-EZltuaagu2dyFj0wf-lch2qIhXKaycyCM4MYwLa4U",
    "18WH4ocx371qXsuPIf00hvKR1Kc7u1cd2ZTOUsWS377U",
    "1iSH66bKucYRU7St8Ch_-NThYay3C-tD_p_VlLpVS7WA",
    "1toIzc0CpyQIditC9KqeMVf3II2CWAW4c96pqp1b1t6s",
    "1nrlXxHexPwEBJCyXUt6ehLSMUoC11A6tTr8B5MK1Tuw",
    "11cK2I6pQ_hrHMrbs0AhKmqLLiCbOV00R7tboG0CdrS0",
    "1-T0dKoRATQ4uWmYYli6qU7Jq9FvblD7AGW2wQZFoUB0",
    "10aFfdXjkLNlt_H0D9e4VJnxFJalGX5R9zLzAf8XpQDw",
    "1Gb0RJDqr-Z34D9dHXfPVcGJguKP46YKgPq1qgT_A4nI",
    "1-BdML7lK0fW3vrcxl8yTpgL9FpeM2NUotigf6JIN4CQ",
    "12EVYbShzbwujbqm42rkXZACrRbblo_2Ls-8XmUnID2I",
    "1YmNqdrkLeQBH5Nnq_gVdeVs3LDHArWXfHjo1BBYfpD4",
    "1mJfVK1dCSIMV4ME2lHXUyT1NXD0qaQgr_1pKqMPgCe4",
    "1MLsLwaimRSR6Gdhcm9fuatNs1B-kTgVDfUfbm6BSCjY",
    "12hSjYYlTj9DrVq4PTkI9IwTD4radEgG-Z0bYpgddXiU",
    "1_Ch5eolJHF5JQdSH7uLweBYvIZ1DomhED7pdWRdn7oI",
    "1uhvZlw1GEbMdDi5sGRJoPkd5tyvKnvob0XU0uKl99Z4",
    "1xohqcQR6vpLcmvWPNgbIPO1UCDQJmPILQ29yUpnH_A0",
    "1yDlr5nqVkoEpzPDdKRFyHTxSYJwnlyzBQ_NaJOev300",
    "1D6EIWhpH-QgjL1HvVqX54cQQl_52GJs0oAXWNQF8FKc",
    "1mP6RJtA918WUi8zq7N2RmPh4jMJfGa41UZT0mu3U4nQ",
    "1LbRFZb3830m77GKIBVipifmW6D0kCxWLP-Etku6sYts",
    "1XmQDC7hk0VYV1TQ9ETf9ZJfll5On41ZD2742H5fRmT0",
    "1j_VMVVb8CkM883y8nw2b1BHbTCCFw_43KkMZWhEK1SM",
    "1QMCAddnW5qn5OG9awAyI7Jqeo5Jzdh3mvEHuwqVBHsU",
    "1LjNHy0nsNqjeHRoAfCbGIRh_0QsLoSWSLPwRt4pik58",
    "1Ii1LlQRHtq8dqyZHCtFkCNti37fFK5ff9qbPxNHhcVw",
    "1S7oJBkx9NjsXramXxYeDq36zvN-3a--9f9KhhilfiNc",
    "1jJ8nz7lzFOgz40bN12kX1cyjkhCcquba6H9QHn8kldE",
    "1UaWDeG1pcNbGvvMrVSaRKGX6nqNdEmpQ92mueII1VRE",
    "1YuRCpm_iZkuK-eVqJ3EN5rCgEBOSgaY3qg5shaszb9A"
]

# ФИО для поиска листов
FIO_LIST = [
    "Narziyeva Gavxar Atxamjanovna",
    "Xoshimov Akromjon Axmadjon o'g'li",
    "Farxodov Xusniddin Murodjon o'g'li",
    "Rahimjonov Kamoliddin Olimjon o'g'li",
    "Xusniddinova Shaxnoza Akramovna",
    "Qosimov Firdavs Nuriddin o'g'li",
    "Abdullayev Dilmurod Xayrulla o'g'li",
    "Sirojiddinov Ismoilbek Shavkat o'g'li",
    "Mavlyanova Dilobar Rustam qizi",
    "Zokirjonova Surayyo Rustam qizi",
    "Payziyeva Shoxista Navro'zjon qizi",
    "Muxamadaliyeva Mufazzal Abduqaxxorovna",
    "Turg'unboyeva Azizaxon Shuxrat qizi",
    "Ruziyeva Dilnoza Xoshimjonovna",
    "Sobirjonova Umidaxon Rustamovna",
    "Karimova Durdona Toir qizi",
    "Xasanova Maftuna Askar qizi",
    "Sagdullayeva Moxinur Asqar qizi",
    "Mirbabayeva Shirin Kaxramonovna"
]

def normalize_phone(phone):
    """Нормализация телефонного номера"""
    if pd.isna(phone):
        return ''
    phone_str = str(phone).replace('.0', '')
    phone_clean = ''.join(filter(str.isdigit, phone_str))
    if phone_clean.startswith('998') and len(phone_clean) == 12:
        phone_clean = phone_clean[3:]
    return phone_clean

def rename_statuses(status):
    """Переименование статусов"""
    if pd.isna(status):
        return status
    
    status_str = str(status).strip()
    
    statuses_to_rename = [
        'НЕТ ОТВЕТА (ЗАНЯТО)',
        'Заявка закрыта (не удалось дозвониться)',
        'Тиббиёт ходими аризаси',
        'Открыть карту',
        'Не удалось дозвониться'
    ]
    
    if status_str in statuses_to_rename:
        return 'Не удалось дозвониться'
    
    return status_str

def get_credentials():
    """Получение учётных данных для Google Sheets API"""
    creds = None
    
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)
        
        TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(creds.to_json())
    
    return creds

def load_sheets_data(use_api=False):
    """
    Загрузка данных из Google Sheets
    
    Args:
        use_api: True - загружать через API, False - использовать локальный файл
    """
    print("="*80)
    print("ИМПОРТ ДАННЫХ ИЗ GOOGLE SHEETS")
    print("="*80)
    
    # Проверяем наличие локального файла
    local_file = Path('data/КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv')
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    
    if not use_api and local_files:
        print("\nИспользуем локальный файл вместо API...")
        latest_file = max(local_files, key=lambda p: p.stat().st_ctime)
        print(f"Файл: {latest_file}")
        
        df_sheets = pd.read_csv(latest_file)
        print(f"Загружено записей: {len(df_sheets)}")
        
        # Переименовываем колонки
        col_mapping = {
            'Колонка_2': 'Номер_карты',
            'Колонка_3': 'Телефон_Sheets',
            'Колонка_4': 'Дата_звонка',
            'Колонка_5': 'Статус_Sheets',
            'Колонка_6': 'Служба_Sheets',
            'Колонка_7': 'Район_Sheets',
            'Колонка_8': 'Оператор_Sheets',
            'Колонка_9': 'Дата_обработки',
            'Колонка_10': 'Положительно',
            'Колонка_11': 'Жалоба',
            'Колонка_12': 'Дополнительно'
        }
        df_sheets = df_sheets.rename(columns=col_mapping)
        
        # Нормализация
        df_sheets['Телефон_нормализованный'] = df_sheets['Телефон_Sheets'].apply(normalize_phone)
        df_sheets['Статус_Sheets'] = df_sheets['Статус_Sheets'].apply(rename_statuses)
        df_sheets['Номер_карты_norm'] = df_sheets['Номер_карты'].astype(str).str.strip()
        df_sheets['Есть_жалоба'] = df_sheets['Жалоба'].notna() & (df_sheets['Жалоба'].astype(str).str.strip() != '')
        
        print(f"Уникальных карт: {df_sheets['Номер_карты_norm'].nunique()}")
        print(f"Записей с жалобами: {df_sheets['Есть_жалоба'].sum()}")
        
        return df_sheets
    
    # Иначе загружаем через API
    print("\nЗагрузка через Google Sheets API...")
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    # Листы для импорта
    sheets_to_import = ["FIKSA"]
    for fio in FIO_LIST:
        sheets_to_import.append(f"{fio} 01.2026")
    
    all_data = []
    
    for spreadsheet_id in SPREADSHEET_IDS:
        print(f"\nОбработка документа: {spreadsheet_id[:20]}...")
        
        try:
            spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            available_sheets = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
            
            for sheet_name in sheets_to_import:
                if sheet_name not in available_sheets:
                    continue
                
                print(f"  Импорт листа: {sheet_name}")
                
                try:
                    # Читаем все данные листа
                    result = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id,
                        range=f"'{sheet_name}'!A:Z"
                    ).execute()
                    
                    values = result.get('values', [])
                    if not values or len(values) < 2:
                        print(f"    Лист пуст, пропускаем")
                        continue
                    
                    # Преобразуем в DataFrame
                    df = pd.DataFrame(values[1:], columns=values[0])
                    
                    # Фильтруем только полностью заполненные строки
                    # Проверяем наличие ключевых колонок
                    required_cols = []
                    if len(df.columns) >= 12:
                        # Берём колонки 2-12 (индексы 1-11)
                        for i in range(1, 12):
                            if i < len(df.columns):
                                required_cols.append(df.columns[i])
                    
                    if required_cols:
                        # Оставляем только строки, где все ключевые колонки заполнены
                        df_filled = df.dropna(subset=required_cols, how='any')
                        df_filled = df_filled[df_filled[required_cols].ne('').all(axis=1)]
                        
                        if len(df_filled) > 0:
                            df_filled['Источник_лист'] = sheet_name
                            df_filled['Источник_документ'] = spreadsheet_id
                            all_data.append(df_filled)
                            print(f"    Импортировано: {len(df_filled)} заполненных строк")
                        else:
                            print(f"    Нет полностью заполненных строк")
                    
                except Exception as e:
                    print(f"    ОШИБКА при чтении листа: {e}")
                    continue
        
        except Exception as e:
            print(f"  ОШИБКА при обработке документа: {e}")
            continue
    
    if not all_data:
        print("\nНЕ НАЙДЕНО ДАННЫХ ДЛЯ ИМПОРТА!")
        return pd.DataFrame()
    
    # Объединяем все данные
    df_sheets = pd.concat(all_data, ignore_index=True)
    
    print(f"\n{'='*80}")
    print(f"ИТОГО ИМПОРТИРОВАНО: {len(df_sheets)} записей")
    print(f"{'='*80}")
    
    # Переименовываем колонки (берём колонки 2-12)
    if len(df_sheets.columns) >= 12:
        col_mapping = {}
        col_names = [
            'Номер_карты',  # Колонка 2
            'Телефон_Sheets',  # Колонка 3
            'Дата_звонка',  # Колонка 4
            'Статус_Sheets',  # Колонка 5
            'Служба_Sheets',  # Колонка 6
            'Район_Sheets',  # Колонка 7
            'Оператор_Sheets',  # Колонка 8
            'Дата_обработки',  # Колонка 9
            'Положительно',  # Колонка 10
            'Жалоба',  # Колонка 11
            'Дополнительно'  # Колонка 12
        ]
        
        for i, name in enumerate(col_names, start=1):
            if i < len(df_sheets.columns):
                col_mapping[df_sheets.columns[i]] = name
        
        df_sheets = df_sheets.rename(columns=col_mapping)
    
    # Нормализация
    df_sheets['Телефон_нормализованный'] = df_sheets['Телефон_Sheets'].apply(normalize_phone)
    df_sheets['Статус_Sheets'] = df_sheets['Статус_Sheets'].apply(rename_statuses)
    df_sheets['Номер_карты_norm'] = df_sheets['Номер_карты'].astype(str).str.strip()
    df_sheets['Есть_жалоба'] = df_sheets['Жалоба'].notna() & (df_sheets['Жалоба'].astype(str).str.strip() != '')
    
    print(f"Уникальных карт: {df_sheets['Номер_карты_norm'].nunique()}")
    print(f"Записей с жалобами: {df_sheets['Есть_жалоба'].sum()}")
    
    return df_sheets

def load_112_data_january():
    """Загрузка данных 112 за январь из папки 123"""
    print("\n" + "="*80)
    print("ЗАГРУЗКА ДАННЫХ 112 ЗА ЯНВАРЬ")
    print("="*80)
    
    files = glob.glob('123/ЧақирувТарихи_112_*.xlsx')
    print(f"Найдено файлов: {len(files)}")
    
    all_data = []
    for file in files:
        print(f"\nЧитаю файл: {file}")
        df = pd.read_excel(file)
        print(f"  Загружено строк: {len(df)}")
        print(f"  Колонок: {len(df.columns)}")
        all_data.append(df)
    
    if not all_data:
        raise ValueError("Не найдено файлов 112!")
    
    df_112 = pd.concat(all_data, ignore_index=True)
    
    print(f"\nВсего строк после объединения: {len(df_112)}")
    
    # Удаляем дубликаты
    print("\nУдаление дубликатов...")
    initial_count = len(df_112)
    df_112 = df_112.drop_duplicates()
    duplicates_removed = initial_count - len(df_112)
    print(f"Удалено полных дубликатов: {duplicates_removed}")
    
    # Переименование колонок (узбекские названия -> русские)
    df_112 = df_112.rename(columns={
        'Карточка рақами': 'Карта_112',
        'Ҳодиса рақами': 'Инцидент_112',
        'Хизмат': 'Служба_112',
        'Мурожаатчи телефон рақами': 'Телефон_112',
        'Ҳолат': 'Статус_112',
        'Туман': 'Район_112',
        'Оператор': 'Оператор_112',
        'Сана': 'Дата_112'
    })
    
    # Нормализация
    df_112['Телефон_нормализованный'] = df_112['Телефон_112'].apply(normalize_phone)
    df_112['Статус_112'] = df_112['Статус_112'].apply(rename_statuses)
    df_112['Карта_112_norm'] = df_112['Карта_112'].astype(str).str.strip()
    df_112['Инцидент_112_norm'] = df_112['Инцидент_112'].astype(str).str.strip()
    df_112['Служба_112'] = df_112['Служба_112'].astype(str)
    
    # Удаляем дубликаты по ключевым полям
    df_112 = df_112.drop_duplicates(
        subset=['Инцидент_112_norm', 'Карта_112_norm', 'Служба_112'],
        keep='first'
    )
    
    print(f"\nИтого записей 112: {len(df_112)}")
    print(f"Уникальных инцидентов: {df_112['Инцидент_112_norm'].nunique()}")
    print(f"Уникальных карт: {df_112['Карта_112_norm'].nunique()}")
    print(f"\nСлужбы: {sorted(df_112['Служба_112'].unique())}")
    
    return df_112

def match_data(df_sheets, df_112):
    """
    Сопоставление данных с учетом логики:
    1. Приоритет - сопоставление по номеру карты
    2. Если есть жалоба - ищем конкретную службу в инциденте
    3. Если нет жалобы и инцидент с несколькими службами - всем положительно
    """
    print("\n" + "="*80)
    print("СОПОСТАВЛЕНИЕ ДАННЫХ")
    print("="*80)
    
    # Считаем количество служб в каждом инциденте
    incident_counts = df_112.groupby('Инцидент_112_norm').agg({
        'Служба_112': 'count'
    }).rename(columns={'Служба_112': 'Количество_служб'}).to_dict()['Количество_служб']
    
    # ОСНОВНОЕ СОПОСТАВЛЕНИЕ: ПО НОМЕРУ КАРТЫ
    print("\nСопоставление по номеру карты...")
    result = pd.merge(
        df_sheets,
        df_112,
        left_on='Номер_карты_norm',
        right_on='Карта_112_norm',
        how='left',
        indicator=True
    )
    
    matched = result[result['_merge'] == 'both']
    unmatched = result[result['_merge'] == 'left_only']
    
    print(f"Найдено совпадений по карте: {len(matched)}")
    print(f"Не найдено совпадений: {len(unmatched)}")
    
    if len(matched) > 0:
        # Добавляем количество служб в инциденте
        matched['Количество_служб_в_инциденте'] = matched['Инцидент_112_norm'].map(incident_counts)
        
        # ПРИМЕНЯЕМ ЛОГИКУ
        print("\nПрименение логики жалоб и положительных...")
        
        # Для записей с жалобами - оставляем как есть
        mask_complaint = matched['Есть_жалоба']
        matched.loc[mask_complaint, 'Тип_совпадения'] = 'Жалоба - по номеру карты'
        
        # Для записей без жалоб - ставим положительно
        mask_no_complaint = ~matched['Есть_жалоба']
        matched.loc[mask_no_complaint, 'Положительно'] = 'Положительно'
        matched.loc[mask_no_complaint & (matched['Количество_служб_в_инциденте'] > 1), 'Тип_совпадения'] = 'Положительно - несколько служб'
        matched.loc[mask_no_complaint & (matched['Количество_служб_в_инциденте'] == 1), 'Тип_совпадения'] = 'Положительно - одна служба'
        
        print("\nРаспределение по типам совпадений:")
        print(matched['Тип_совпадения'].value_counts())
        
        print("\nРаспределение по службам из 112:")
        print(matched['Служба_112'].value_counts())
    
    # Объединяем результаты
    result_final = pd.concat([matched, unmatched], ignore_index=True)
    result_final.loc[result_final['_merge'] == 'left_only', 'Тип_совпадения'] = 'Не найдено в 112'
    
    return result_final

def save_results(df_result):
    """Сохранение результатов"""
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    
    # Создаём папку reports если её нет
    reports_dir = Path('reports')
    reports_dir.mkdir(exist_ok=True)
    
    # CSV файл с полными данными
    csv_file = reports_dir / f'ЯНВАРЬ_2026_ПОЛНЫЙ_ОТЧЕТ_{timestamp}.csv'
    df_result.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"\n{'='*80}")
    print(f"СОХРАНЁН CSV ФАЙЛ: {csv_file}")
    
    # Excel файл для удобства
    excel_file = reports_dir / f'ЯНВАРЬ_2026_ПОЛНЫЙ_ОТЧЕТ_{timestamp}.xlsx'
    df_result.to_excel(excel_file, index=False, engine='openpyxl')
    print(f"СОХРАНЁН EXCEL ФАЙЛ: {excel_file}")
    print(f"{'='*80}")
    
    # Статистика
    print("\n" + "="*80)
    print("СТАТИСТИКА СОПОСТАВЛЕНИЯ")
    print("="*80)
    
    if 'Тип_совпадения' in df_result.columns:
        print("\nПо типам совпадений:")
        print(df_result['Тип_совпадения'].value_counts())
    
    if 'Служба_112' in df_result.columns:
        matched_count = df_result['Служба_112'].notna().sum()
        print(f"\nВсего сопоставлено записей: {matched_count}")
        print(f"Всего записей из Sheets: {len(df_result)}")
        print(f"Процент совпадений: {matched_count/len(df_result)*100:.1f}%")
    
    return csv_file, excel_file

def main(use_api=False):
    """
    Основная функция
    
    Args:
        use_api: True - загружать через API, False - использовать локальный файл
    """
    print("\n" + "="*80)
    print("ИМПОРТ И СОПОСТАВЛЕНИЕ ДАННЫХ ЗА ЯНВАРЬ 2026")
    print("="*80)
    print(f"Дата запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    try:
        # 1. Загружаем данные из Google Sheets
        df_sheets = load_sheets_data(use_api=use_api)
        if df_sheets.empty:
            print("\nНЕТ ДАННЫХ ИЗ GOOGLE SHEETS!")
            return
        
        # 2. Загружаем данные 112 за январь
        df_112 = load_112_data_january()
        
        # 3. Сопоставляем данные
        df_result = match_data(df_sheets, df_112)
        
        # 4. Сохраняем результаты
        csv_file, excel_file = save_results(df_result)
        
        print("\n" + "="*80)
        print("ГОТОВО!")
        print("="*80)
        print(f"\nРезультаты сохранены в:")
        print(f"  - {csv_file}")
        print(f"  - {excel_file}")
        
    except Exception as e:
        print(f"\n{'='*80}")
        print(f"ОШИБКА: {e}")
        print(f"{'='*80}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
