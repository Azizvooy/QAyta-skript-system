#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ПРОСТАЯ ОБРАБОТКА ЯНВАРСКИХ ДАННЫХ
=============================================================================
Сопоставление данных из папки 123 с локальным файлом Google Sheets
=============================================================================
"""

import pandas as pd
import glob
from datetime import datetime
from pathlib import Path

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

def load_sheets_data():
    """Загрузка данных из локального файла Google Sheets"""
    print("="*80)
    print("ЗАГРУЗКА ДАННЫХ ИЗ GOOGLE SHEETS")
    print("="*80)
    
    # Ищем последний файл
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    
    if not local_files:
        print("\nЛОКАЛЬНЫЙ ФАЙЛ НЕ НАЙДЕН!")
        print("Сначала запустите: python import_sheets_consolidated.py")
        return pd.DataFrame()
    
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
    
    # Заполнение пустых статусов
    print("\nЗаполнение пустых полей...")
    
    # Если статус пустой - ставим "Не удалось дозвониться"
    empty_status = df_sheets['Статус_Sheets'].isna() | (df_sheets['Статус_Sheets'].astype(str).str.strip() == '')
    if empty_status.sum() > 0:
        df_sheets.loc[empty_status, 'Статус_Sheets'] = 'Не удалось дозвониться'
        print(f"  Заполнено пустых статусов: {empty_status.sum()}")
    
    # Если "Положительно" пустое и нет жалобы - ставим "Нет"
    empty_positive = df_sheets['Положительно'].isna() | (df_sheets['Положительно'].astype(str).str.strip() == '')
    no_complaint = ~df_sheets['Есть_жалоба']
    to_fill = empty_positive & no_complaint
    if to_fill.sum() > 0:
        df_sheets.loc[to_fill, 'Положительно'] = 'Нет'
        print(f"  Заполнено пустых 'Положительно' (без жалоб): {to_fill.sum()}")
    
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
    
    if not files:
        print("\nФАЙЛЫ НЕ НАЙДЕНЫ В ПАПКЕ 123!")
        return pd.DataFrame()
    
    all_data = []
    for file in files:
        print(f"\nЧитаю файл: {Path(file).name}")
        df = pd.read_excel(file)
        print(f"  Загружено строк: {len(df)}")
        all_data.append(df)
    
    df_112 = pd.concat(all_data, ignore_index=True)
    
    print(f"\nВсего строк после объединения: {len(df_112)}")
    
    # Удаляем дубликаты
    print("Удаление дубликатов...")
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
    
    if 'Служба_112' in df_112.columns:
        print(f"\nСлужбы ({df_112['Служба_112'].nunique()}):")
        for service in sorted(df_112['Служба_112'].unique()):
            count = (df_112['Служба_112'] == service).sum()
            print(f"  {service}: {count}")
    
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
    
    if df_sheets.empty or df_112.empty:
        print("\nНЕТ ДАННЫХ ДЛЯ СОПОСТАВЛЕНИЯ!")
        return pd.DataFrame()
    
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
    
    matched = result[result['_merge'] == 'both'].copy()
    unmatched = result[result['_merge'] == 'left_only'].copy()
    
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
    csv_file = reports_dir / f'ЯНВАРЬ_2026_ОТЧЕТ_{timestamp}.csv'
    df_result.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"\n{'='*80}")
    print(f"СОХРАНЁН CSV ФАЙЛ: {csv_file}")
    
    # Excel файл для удобства
    excel_file = reports_dir / f'ЯНВАРЬ_2026_ОТЧЕТ_{timestamp}.xlsx'
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
        
        print("\nРаспределение по службам:")
        services = df_result[df_result['Служба_112'].notna()]['Служба_112'].value_counts()
        for service, count in services.items():
            print(f"  {service}: {count}")
    
    return csv_file, excel_file

def main():
    """Основная функция"""
    print("\n" + "="*80)
    print("ОБРАБОТКА ЯНВАРСКИХ ДАННЫХ - УПРОЩЁННАЯ ВЕРСИЯ")
    print("="*80)
    print(f"Дата запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    try:
        # 1. Загружаем данные из Google Sheets (локальный файл)
        df_sheets = load_sheets_data()
        if df_sheets.empty:
            return
        
        # 2. Загружаем данные 112 за январь
        df_112 = load_112_data_january()
        if df_112.empty:
            return
        
        # 3. Сопоставляем данные
        df_result = match_data(df_sheets, df_112)
        if df_result.empty:
            return
        
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
