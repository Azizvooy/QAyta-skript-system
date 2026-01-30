"""
Сопоставление данных из Google Sheets с данными 112 по инцидентам и службам
С учетом логики жалоб и переименования статусов
"""
import pandas as pd
import glob
from datetime import datetime
import os

def normalize_phone(phone):
    """Нормализация телефонного номера"""
    if pd.isna(phone):
        return ''
    phone_str = str(phone).replace('.0', '')
    # Убираем все нецифровые символы
    phone_clean = ''.join(filter(str.isdigit, phone_str))
    # Убираем префикс 998 если есть
    if phone_clean.startswith('998') and len(phone_clean) == 12:
        phone_clean = phone_clean[3:]
    return phone_clean

def rename_statuses(status):
    """Переименование статусов в 'Не удалось дозвониться'"""
    if pd.isna(status):
        return status
    
    status_str = str(status).strip()
    
    statuses_to_rename = [
        'НЕТ ОТВЕТА (ЗАНЯТО)',
        'Заявка закрыта (не удалось дозвониться)',
        'Тиббиёт ходими аризаси',
        'Открыть карту'
    ]
    
    if status_str in statuses_to_rename:
        return 'Не удалось дозвониться'
    
    return status_str

def load_112_data():
    """Загрузка всех файлов 112 (только файлы журнала с детальными данными)"""
    print("Загрузка данных из файлов 112...")
    
    files = glob.glob('123/ИсторияЗвонков_112_*.xlsx')
    print(f"Найдено файлов: {len(files)}")
    
    all_data = []
    for file in files:
        print(f"Читаю файл: {file}")
        df = pd.read_excel(file)
        
        # Проверяем тип файла - нужны только файлы ЖУРНАЛА (с колонкой "Служба")
        if 'Служба' not in df.columns:
            print(f"  ПРОПУЩЕН (сводный файл без детализации по службам)")
            continue
        
        print(f"  Загружено строк: {len(df)}")
        all_data.append(df)
    
    if not all_data:
        raise ValueError("Не найдено ни одного файла журнала с детальными данными!")
    
    df_112 = pd.concat(all_data, ignore_index=True)
    
    print(f"\nВсего строк после объединения: {len(df_112)}")
    
    # ВАЖНО: Удаляем дубликаты (если идентичные строки - это одно и то же)
    print("Удаление дубликатов...")
    initial_count = len(df_112)
    
    # Удаляем полные дубликаты строк
    df_112 = df_112.drop_duplicates()
    
    duplicates_removed = initial_count - len(df_112)
    print(f"Удалено полных дубликатов: {duplicates_removed}")
    
    # Дополнительно: удаляем дубликаты по ключевым полям
    # (один и тот же инцидент + карта + служба = дубликат)
    df_112 = df_112.drop_duplicates(
        subset=['Номер инцидента', 'Номер карточки', 'Служба', 'Номер телефона заявителя'],
        keep='first'
    )
    
    duplicates_removed_key = initial_count - duplicates_removed - len(df_112)
    print(f"Удалено дубликатов по ключевым полям: {duplicates_removed_key}")
    
    df_112 = df_112.reset_index(drop=True)
    
    # Переименование колонок для удобства
    df_112 = df_112.rename(columns={
        'Номер инцидента': 'Инцидент_112',
        'Номер карточки': 'Карта_112',
        'Служба': 'Служба_112',
        'Номер телефона заявителя': 'Телефон_112',
        'Статус': 'Статус_112',
        'Район': 'Район_112',
        'Оператор': 'Оператор_112',
        'Дата': 'Дата_112'
    })
    
    # Нормализация телефонов
    df_112['Телефон_нормализованный'] = df_112['Телефон_112'].apply(normalize_phone)
    
    # Переименование статусов
    df_112['Статус_112'] = df_112['Статус_112'].apply(rename_statuses)
    
    # Приводим службу к строковому типу для слияния
    df_112['Служба_112'] = df_112['Служба_112'].astype(str)
    
    # ВАЖНО: Исключаем записи без службы
    df_112 = df_112[df_112['Служба_112'] != 'nan'].copy()
    
    # Нормализуем номер карты для сопоставления
    df_112['Карта_112_norm'] = df_112['Карта_112'].astype(str).str.strip()
    
    # Нормализуем номер инцидента
    df_112['Инцидент_112_norm'] = df_112['Инцидент_112'].astype(str).str.strip()
    
    print(f"Всего записей 112: {len(df_112)}")
    print(f"Уникальных инцидентов: {df_112['Инцидент_112_norm'].nunique()}")
    print(f"Уникальных карт: {df_112['Карта_112_norm'].nunique()}")
    print(f"Служб: {sorted(df_112['Служба_112'].unique())}")
    
    return df_112

def load_sheets_data():
    """Загрузка данных из Google Sheets"""
    print("\nЗагрузка данных из Google Sheets...")
    
    # Ищем последний файл
    files = glob.glob('data/КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv')
    if not files:
        raise FileNotFoundError("Не найден файл с консолидированными данными")
    
    latest_file = max(files, key=os.path.getctime)
    print(f"Читаю файл: {latest_file}")
    
    df_sheets = pd.read_csv(latest_file)
    
    # Определяем какие колонки что содержат (по порядку из скрипта импорта)
    # Колонки 2-12: Номер карты, Телефон, Дата и время звонка, Статус, Служба, 
    #               Район/город, Оператор, Дата обработки, Положительно, Жалоба, Дополнительно
    
    df_sheets = df_sheets.rename(columns={
        'Колонка_2': 'Инцидент_Sheets',  # ЭТО ИНЦИДЕНТ, НЕ КАРТА!
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
    })
    
    # Нормализация телефонов
    df_sheets['Телефон_нормализованный'] = df_sheets['Телефон_Sheets'].apply(normalize_phone)
    
    # Переименование статусов
    df_sheets['Статус_Sheets'] = df_sheets['Статус_Sheets'].apply(rename_statuses)
    
    # Приводим службу к строковому типу для слияния
    df_sheets['Служба_Sheets'] = df_sheets['Служба_Sheets'].astype(str)
    
    # Нормализуем номер инцидента для сопоставления
    df_sheets['Инцидент_Sheets_norm'] = df_sheets['Инцидент_Sheets'].astype(str).str.strip()
    
    print(f"Всего записей из Sheets: {len(df_sheets)}")
    print(f"Уникальных телефонов: {df_sheets['Телефон_нормализованный'].nunique()}")
    print(f"Уникальных инцидентов: {df_sheets['Инцидент_Sheets_norm'].nunique()}")
    
    # Проверяем жалобы
    has_complaint = df_sheets['Жалоба'].notna().sum()
    print(f"Записей с жалобами: {has_complaint}")
    
    return df_sheets

def match_data_by_incident(df_sheets, df_112):
    """
    Сопоставление данных с учетом логики:
    1. Сопоставление по номеру карты (первичный ключ)
    2. Если есть жалоба - ищем конкретную службу в инциденте
    3. Если нет жалобы и инцидент с несколькими службами - всем положительно
    """
    print("\nНачинаем сопоставление по картам и инцидентам...")
    
    # Подготовка: считаем количество служб в каждом инциденте
    print("Подготовка данных 112...")
    incident_counts = df_112.groupby('Инцидент_112').size().to_dict()
    
    # Добавляем флаг жалобы
    print("Обработка жалоб...")
    df_sheets['Есть_жалоба'] = df_sheets['Жалоба'].notna() & (df_sheets['Жалоба'].astype(str).str.strip() != '')
    
    # ОСНОВНОЕ СОПОСТАВЛЕНИЕ: ПО НОМЕРУ ИНЦИДЕНТА
    print("\nСопоставление по номеру инцидента...")
    result = pd.merge(
        df_sheets,
        df_112,
        left_on='Инцидент_Sheets_norm',
        right_on='Инцидент_112_norm',
        how='inner'
    )
    
    print(f"Найдено совпадений по номеру инцидента: {len(result)}")
    
    if len(result) == 0:
        print("НЕТ СОВПАДЕНИЙ! Проверьте формат номеров инцидентов.")
        return pd.DataFrame()
    
    # Добавляем количество служб в инциденте
    result['Количество_служб_в_инциденте'] = result['Инцидент_112_norm'].map(incident_counts)
    
    # ПРИМЕНЯЕМ ЛОГИКУ
    print("\nПрименение логики жалоб и положительных...")
    
    # Для записей с жалобами - оставляем как есть
    mask_complaint = result['Есть_жалоба']
    result.loc[mask_complaint, 'Тип_совпадения'] = 'Жалоба - по номеру инцидента'
    
    # Для записей без жалоб - ставим положительно
    mask_no_complaint = ~result['Есть_жалоба']
    result.loc[mask_no_complaint, 'Положительно'] = 'Положительно'
    result.loc[mask_no_complaint & (result['Количество_служб_в_инциденте'] > 1), 'Тип_совпадения'] = 'Положительно - несколько служб'
    result.loc[mask_no_complaint & (result['Количество_служб_в_инциденте'] == 1), 'Тип_совпадения'] = 'Положительно - одна служба'
    
    df_result = result
    
    print(f"\nИтого сопоставлено: {len(df_result)} записей")
    
    if len(df_result) > 0:
        print("\nРаспределение по типам совпадений:")
        print(df_result['Тип_совпадения'].value_counts())
        
        print("\nРаспределение по службам из 112:")
        print(df_result['Служба_112'].value_counts())
    
    return df_result

def save_results(df_result):
    """Сохранение результатов"""
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    
    # CSV файл с полными данными
    csv_file = f'reports/СОПОСТАВЛЕНИЕ_ПО_ИНЦИДЕНТАМ_{timestamp}.csv'
    df_result.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"\nСохранён CSV файл: {csv_file}")
    
    # Excel файл (ограничено 1млн строк)
    excel_file = f'reports/СОПОСТАВЛЕНИЕ_ПО_ИНЦИДЕНТАМ_{timestamp}.xlsx'
    if len(df_result) <= 1000000:
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Сопоставление')
        print(f"Сохранён Excel файл: {excel_file}")
    else:
        print("Excel файл не создан (слишком много строк)")
    
    # Текстовый отчёт
    txt_file = f'reports/ОТЧЁТ_СОПОСТАВЛЕНИЕ_ИНЦИДЕНТЫ_{timestamp}.txt'
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("ОТЧЁТ: СОПОСТАВЛЕНИЕ ПО ИНЦИДЕНТАМ И СЛУЖБАМ\n")
        f.write("="*80 + "\n")
        f.write(f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Всего записей: {len(df_result)}\n\n")
        
        f.write("РАСПРЕДЕЛЕНИЕ ПО ТИПАМ СОВПАДЕНИЙ:\n")
        f.write("-"*80 + "\n")
        for tipo, count in df_result['Тип_совпадения'].value_counts().items():
            pct = count / len(df_result) * 100
            f.write(f"{tipo:50} {count:10,} ({pct:5.1f}%)\n")
        
        f.write("\n" + "="*80 + "\n")
        f.write("РАСПРЕДЕЛЕНИЕ ПО СЛУЖБАМ:\n")
        f.write("-"*80 + "\n")
        for service, count in df_result['Служба_112'].value_counts().items():
            pct = count / len(df_result) * 100
            f.write(f"Служба {service:3} {count:10,} ({pct:5.1f}%)\n")
        
        f.write("\n" + "="*80 + "\n")
        f.write("СТАТУСЫ ИЗ 112:\n")
        f.write("-"*80 + "\n")
        for status, count in df_result['Статус_112'].value_counts().head(20).items():
            pct = count / len(df_result) * 100
            f.write(f"{str(status):50} {count:10,} ({pct:5.1f}%)\n")
        
        f.write("\n" + "="*80 + "\n")
        f.write("СТАТУСЫ ИЗ SHEETS (после переименования):\n")
        f.write("-"*80 + "\n")
        for status, count in df_result['Статус_Sheets'].value_counts().head(20).items():
            pct = count / len(df_result) * 100
            f.write(f"{str(status):50} {count:10,} ({pct:5.1f}%)\n")
        
        f.write("\n" + "="*80 + "\n")
        f.write("ТОП-20 АГЕНТОВ:\n")
        f.write("-"*80 + "\n")
        for agent, count in df_result['Документ'].value_counts().head(20).items():
            pct = count / len(df_result) * 100
            f.write(f"{agent:20} {count:10,} ({pct:5.1f}%)\n")
    
    print(f"Сохранён текстовый отчёт: {txt_file}")
    
    return csv_file, txt_file

def main():
    """Основная функция"""
    print("="*80)
    print("СОПОСТАВЛЕНИЕ ДАННЫХ ПО ИНЦИДЕНТАМ И СЛУЖБАМ")
    print("="*80)
    
    # Загрузка данных
    df_112 = load_112_data()
    df_sheets = load_sheets_data()
    
    # Сопоставление
    df_result = match_data_by_incident(df_sheets, df_112)
    
    # Сохранение
    if len(df_result) > 0:
        csv_file, txt_file = save_results(df_result)
        
        print("\n" + "="*80)
        print("ГОТОВО!")
        print("="*80)
        print(f"Создано файлов: 3 (CSV + TXT + Excel)")
        print(f"Всего записей: {len(df_result):,}")
    else:
        print("\nНЕТ СОВПАДЕНИЙ!")

if __name__ == '__main__':
    main()
