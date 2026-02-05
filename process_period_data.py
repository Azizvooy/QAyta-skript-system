#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
УНИВЕРСАЛЬНАЯ ОБРАБОТКА ДАННЫХ ПО ПЕРИОДАМ
=============================================================================
Обрабатывает данные из папки 123 за любой период
=============================================================================
"""

import os
import pandas as pd
import glob
from datetime import datetime
from pathlib import Path
import re

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

def scan_123_folder():
    """Сканирование папки 123 и определение доступных периодов"""
    print("="*80)
    print("СКАНИРОВАНИЕ ПАПКИ 123")
    print("="*80)
    
    files = glob.glob('123/*.xlsx')
    
    if not files:
        print("\n❌ Файлы не найдены в папке 123!")
        print("\nПожалуйста, загрузите файлы Excel в папку 123/")
        return []
    
    print(f"\n✓ Найдено файлов: {len(files)}\n")
    
    periods = {}
    for file in files:
        filename = Path(file).name
        print(f"  • {filename}")
        
        # Пытаемся извлечь период из имени файла
        # Формат: ЧақирувТарихи_112_2026_01_04_00_00_00_2026_01_11_23_59_59.xlsx
        match = re.search(r'(\d{4})_(\d{2})_(\d{2}).*?(\d{4})_(\d{2})_(\d{2})', filename)
        if match:
            start_year, start_month, start_day = match.group(1), match.group(2), match.group(3)
            end_year, end_month, end_day = match.group(4), match.group(5), match.group(6)
            period_key = f"{start_year}-{start_month}"
            
            if period_key not in periods:
                periods[period_key] = []
            periods[period_key].append(file)
    
    print(f"\n{'='*80}")
    print(f"Найдено периодов: {len(periods)}")
    for period, files_list in sorted(periods.items()):
        print(f"  • {period}: {len(files_list)} файл(ов)")
    print(f"{'='*80}\n")
    
    return periods

def select_period(periods):
    """Выбор периода для обработки"""
    if not periods:
        return None, []
    
    period_list = sorted(periods.keys())
    
    print("ВЫБЕРИТЕ ПЕРИОД ДЛЯ ОБРАБОТКИ:")
    print("="*80)
    for i, period in enumerate(period_list, 1):
        file_count = len(periods[period])
        print(f"  {i}. {period} ({file_count} файл(ов))")
    print(f"  {len(period_list) + 1}. Обработать ВСЕ файлы")
    print("="*80)
    
    while True:
        try:
            choice = input(f"\nВведите номер (1-{len(period_list) + 1}): ").strip()
            choice_num = int(choice)
            
            if choice_num == len(period_list) + 1:
                # Все файлы
                all_files = []
                for files_list in periods.values():
                    all_files.extend(files_list)
                return "ВСЕ", all_files
            elif 1 <= choice_num <= len(period_list):
                selected_period = period_list[choice_num - 1]
                return selected_period, periods[selected_period]
            else:
                print(f"❌ Введите число от 1 до {len(period_list) + 1}")
        except ValueError:
            print("❌ Введите корректное число")
        except KeyboardInterrupt:
            print("\n\nОтмена операции")
            return None, []

def load_sheets_data():
    """Загрузка данных из локального файла Google Sheets"""
    print("\n" + "="*80)
    print("ЗАГРУЗКА ДАННЫХ ИЗ GOOGLE SHEETS")
    print("="*80)
    
    # Ищем последний файл
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    
    if not local_files:
        print("\n❌ ОШИБКА: Локальный файл не найден!")
        print("\n⚠️  Для работы системы ОБЯЗАТЕЛЬНО нужны данные из Google Sheets!")
        print("\nВыполните на вашем компьютере:")
        print("   1. python import_sheets_consolidated.py")
        print("   2. Загрузите файл КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv в папку data/")
        print("\n" + "="*80)
        return pd.DataFrame()
    
    latest_file = max(local_files, key=lambda p: p.stat().st_ctime)
    print(f"\n✓ Файл: {latest_file.name}")
    
    df_sheets = pd.read_csv(latest_file)
    print(f"✓ Загружено записей: {len(df_sheets)}")
    
    # Переименовываем колонки (структура из Google Sheets)
    # 1 колонка - № (игнорируем)
    # 2 колонка - Номер карты (в файле 112 это НОМЕР ИНЦИДЕНТА)
    # 3 колонка - Телефон обзвона
    # 4 колонка - Дата открытия карты
    # 5 колонка - Статус связи
    # 6 колонка - Служба (может быть несколько)
    # 7 колонка - Комментарий/жалоба
    col_mapping = {
        'Колонка_2': 'Инцидент_Sheets',
        'Колонка_3': 'Телефон_Sheets',
        'Колонка_4': 'Дата_открытия',
        'Колонка_5': 'Статус_связи',
        'Колонка_6': 'Служба_Sheets',
        'Колонка_7': 'Жалоба',
        'Колонка_8': 'Положительно',
        'Колонка_9': 'Примечание',
        'Колонка_10': 'Дополнительно',
        'Колонка_11': 'Дополнительно_2',
        'Колонка_12': 'Дополнительно_3'
    }
    df_sheets = df_sheets.rename(columns=col_mapping)
    
    # Нормализация
    df_sheets['Телефон_нормализованный'] = df_sheets['Телефон_Sheets'].apply(normalize_phone)
    df_sheets['Статус_связи'] = df_sheets['Статус_связи'].apply(rename_statuses)
    df_sheets['Инцидент_Sheets_norm'] = df_sheets['Инцидент_Sheets'].astype(str).str.strip()
    df_sheets['Есть_жалоба'] = df_sheets['Жалоба'].notna() & (df_sheets['Жалоба'].astype(str).str.strip() != '')

    # Разбор служб (может быть несколько)
    def extract_services(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return []
        text = str(value)
        found = re.findall(r"\b(101|102|103|104)\b", text)
        if found:
            return list(dict.fromkeys(found))
        parts = re.split(r"[;,/\\|\s]+", text)
        parts = [p.strip() for p in parts if p.strip()]
        return list(dict.fromkeys(parts))

    df_sheets['Службы_list'] = df_sheets['Служба_Sheets'].apply(extract_services)
    df_sheets['Службы_list'] = df_sheets['Службы_list'].apply(lambda x: x if x else [None])
    df_sheets = df_sheets.explode('Службы_list').reset_index(drop=True)
    df_sheets['Служба_Sheets_norm'] = df_sheets['Службы_list'].astype(str).str.strip()
    
    # Заполнение пустых статусов
    print("\n✓ Заполнение пустых полей...")
    
    # Преобразуем колонки в строковый тип для заполнения
    df_sheets['Статус_связи'] = df_sheets['Статус_связи'].astype(str)
    df_sheets['Положительно'] = df_sheets['Положительно'].astype(str)
    
    # Если статус пустой - ставим "Не удалось дозвониться"
    empty_status = (df_sheets['Статус_связи'].isin(['', 'nan', 'None'])) | df_sheets['Статус_связи'].isna()
    if empty_status.sum() > 0:
        df_sheets.loc[empty_status, 'Статус_связи'] = 'Не удалось дозвониться'
        print(f"  • Заполнено пустых статусов: {empty_status.sum()}")
    
    # Если "Положительно" пустое и нет жалобы - ставим "Нет"
    empty_positive = (df_sheets['Положительно'].isin(['', 'nan', 'None'])) | df_sheets['Положительно'].isna()
    no_complaint = ~df_sheets['Есть_жалоба']
    to_fill = empty_positive & no_complaint
    if to_fill.sum() > 0:
        df_sheets.loc[to_fill, 'Положительно'] = 'Нет'
        print(f"  • Заполнено пустых 'Положительно' (без жалоб): {to_fill.sum()}")
    
    print(f"\n✓ Уникальных инцидентов: {df_sheets['Инцидент_Sheets_norm'].nunique()}")
    print(f"✓ Записей с жалобами: {df_sheets['Есть_жалоба'].sum()}")
    
    return df_sheets

def load_112_data(files_list):
    """Загрузка данных 112 из выбранных файлов"""
    print("\n" + "="*80)
    print("ЗАГРУЗКА ДАННЫХ 112")
    print("="*80)
    
    all_data = []
    for file in files_list:
        print(f"\n✓ Читаю: {Path(file).name}")
        df = pd.read_excel(file)
        print(f"  Строк: {len(df)}")
        all_data.append(df)
    
    df_112 = pd.concat(all_data, ignore_index=True)
    
    print(f"\n✓ Всего строк после объединения: {len(df_112)}")
    
    # Удаляем дубликаты
    initial_count = len(df_112)
    df_112 = df_112.drop_duplicates()
    duplicates_removed = initial_count - len(df_112)
    print(f"✓ Удалено полных дубликатов: {duplicates_removed}")
    
    # Переименование колонок (узбекские названия -> русские)
    df_112 = df_112.rename(columns={
        'Карточка рақами': 'Карта_112',
        'Ҳодиса рақами': 'Инцидент_112',
        'Хизмат': 'Служба_112',
        'Мурожаатчи телефон рақами': 'Телефон_112',
        'Ҳолат': 'Статус_112',
        'Вилоят': 'Регион_112',
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
    
    print(f"\n✓ Итого записей 112: {len(df_112)}")
    print(f"✓ Уникальных инцидентов: {df_112['Инцидент_112_norm'].nunique()}")
    print(f"✓ Уникальных карт: {df_112['Карта_112_norm'].nunique()}")
    
    if 'Служба_112' in df_112.columns:
        print(f"\n✓ Служб: {df_112['Служба_112'].nunique()}")
        for service in sorted(df_112['Служба_112'].unique())[:10]:  # Показываем первые 10
            count = (df_112['Служба_112'] == service).sum()
            print(f"  • {service}: {count}")
    
    return df_112

def match_data(df_sheets, df_112, period_name):
    """Сопоставление данных"""
    print("\n" + "="*80)
    print("СОПОСТАВЛЕНИЕ ДАННЫХ")
    print("="*80)
    
    if df_112.empty:
        print("\n❌ Нет данных 112 для сопоставления!")
        return pd.DataFrame()
    
    if df_sheets.empty:
        print("\n❌ Нет данных Google Sheets для сопоставления!")
        return pd.DataFrame()
    
    # Фильтрация Sheets по периоду и инцидентам 112
    if 'Инцидент_Sheets_norm' in df_sheets.columns:
        incident_set = set(df_112['Инцидент_112_norm'].dropna().unique())
        before_count = len(df_sheets)
        df_sheets = df_sheets[df_sheets['Инцидент_Sheets_norm'].isin(incident_set)]
        after_count = len(df_sheets)
        print(f"\n✓ Отфильтровано по инцидентам 112: {before_count} -> {after_count}")

    # Фильтрация по периоду (например 2026-01)
    if 'Дата_открытия' in df_sheets.columns:
        if period_name and '-' in period_name:
            year, month = period_name.split('-')[0], period_name.split('-')[1]
            period_token = f"{month}.{year}"
            mask_period = df_sheets['Дата_открытия'].astype(str).str.contains(period_token, na=False)
            if mask_period.any():
                df_sheets = df_sheets[mask_period]
                print(f"✓ Отфильтровано по периоду {period_name}: {mask_period.sum()}")

    # Если служба не указана в Sheets — распределяем по службам из 112 по инциденту
    if 'Служба_Sheets_norm' in df_sheets.columns:
        missing_mask = df_sheets['Служба_Sheets_norm'].isin(['', 'None', 'nan']) | df_sheets['Служба_Sheets_norm'].isna()
        if missing_mask.any():
            print(f"✓ Найдены строки без службы: {missing_mask.sum()} — распределяем по службам 112")
            services_map = (
                df_112.groupby('Инцидент_112_norm')['Служба_112']
                .apply(lambda s: sorted(set(s.astype(str))))
                .reset_index(name='Службы_112_list')
            )
            missing_df = df_sheets[missing_mask].merge(
                services_map,
                left_on='Инцидент_Sheets_norm',
                right_on='Инцидент_112_norm',
                how='left'
            )
            missing_df = missing_df[missing_df['Службы_112_list'].notna()].copy()
            missing_df = missing_df.explode('Службы_112_list').reset_index(drop=True)
            missing_df['Служба_Sheets_norm'] = missing_df['Службы_112_list'].astype(str).str.strip()
            missing_df = missing_df.drop(columns=['Службы_112_list', 'Инцидент_112_norm'])
            df_sheets = pd.concat([df_sheets[~missing_mask], missing_df], ignore_index=True)

    # Считаем количество служб в каждом инциденте
    incident_counts = df_112.groupby('Инцидент_112_norm').agg({
        'Служба_112': 'count'
    }).rename(columns={'Служба_112': 'Количество_служб'}).to_dict()['Количество_служб']

    # Уникальные записи по ключу (инцидент + служба) чтобы избежать раздувания merge
    df_112_key = df_112.drop_duplicates(subset=['Инцидент_112_norm', 'Служба_112']).copy()
    df_112_key = df_112_key[
        ['Инцидент_112_norm', 'Служба_112', 'Карта_112', 'Телефон_112', 'Статус_112', 'Регион_112', 'Район_112', 'Оператор_112', 'Дата_112']
    ]
    
    # ОСНОВНОЕ СОПОСТАВЛЕНИЕ: ПО НОМЕРУ ИНЦИДЕНТА + СЛУЖБЕ
    print("\n✓ Сопоставление по номеру инцидента + службе...")
    
    # Создаём ключи для сопоставления
    df_sheets['match_key'] = df_sheets['Инцидент_Sheets_norm'] + '_' + df_sheets['Служба_Sheets_norm']
    df_112_key['match_key'] = df_112_key['Инцидент_112_norm'] + '_' + df_112_key['Служба_112'].astype(str)
    
    result = pd.merge(
        df_sheets,
        df_112_key,
        left_on='match_key',
        right_on='match_key',
        how='left',
        indicator=True
    )
    
    matched = result[result['_merge'] == 'both'].copy()
    unmatched = result[result['_merge'] == 'left_only'].copy()
    
    print(f"  ✓ Найдено совпадений: {len(matched)}")
    print(f"  ⚠ Не найдено: {len(unmatched)}")
    
    if len(matched) > 0:
        # Добавляем количество служб в инциденте
        matched['Количество_служб_в_инциденте'] = matched['Инцидент_112_norm'].map(incident_counts)
        
        # ПРИМЕНЯЕМ ЛОГИКУ
        print("\n✓ Применение логики жалоб и положительных...")
        
        # Для записей с жалобами - оставляем как есть
        mask_complaint = matched['Есть_жалоба']
        matched.loc[mask_complaint, 'Тип_совпадения'] = 'Жалоба - инцидент+служба'
        
        # Для записей без жалоб - ставим положительно
        mask_no_complaint = ~matched['Есть_жалоба']
        matched.loc[mask_no_complaint, 'Положительно'] = 'Положительно'
        matched.loc[mask_no_complaint & (matched['Количество_служб_в_инциденте'] > 1), 'Тип_совпадения'] = 'Положительно - несколько служб'
        matched.loc[mask_no_complaint & (matched['Количество_служб_в_инциденте'] == 1), 'Тип_совпадения'] = 'Положительно - одна служба'

        # Если есть жалоба по одной службе, добавляем положительно для остальных служб в этом инциденте
        complaint_incidents = matched.loc[mask_complaint, 'Инцидент_112_norm'].dropna().unique().tolist()
        if complaint_incidents:
            extras = []
            for incident in complaint_incidents:
                services_in_112 = df_112[df_112['Инцидент_112_norm'] == incident]['Служба_112'].astype(str).unique().tolist()
                complained_services = matched[(matched['Инцидент_112_norm'] == incident) & (matched['Есть_жалоба'])]['Служба_112'].astype(str).unique().tolist()
                other_services = [s for s in services_in_112 if s not in complained_services]
                if not other_services:
                    continue
                base_row = matched[(matched['Инцидент_112_norm'] == incident) & (matched['Есть_жалоба'])].iloc[0].copy()
                for service in other_services:
                    new_row = base_row.copy()
                    new_row['Служба_112'] = service
                    new_row['Служба_Sheets_norm'] = service
                    new_row['match_key'] = f"{incident}_{service}"
                    new_row['Жалоба'] = ''
                    new_row['Есть_жалоба'] = False
                    new_row['Положительно'] = 'Положительно'
                    new_row['Тип_совпадения'] = 'Положительно - другие службы'
                    extras.append(new_row)
            if extras:
                matched = pd.concat([matched, pd.DataFrame(extras)], ignore_index=True)
        
        print("\n  Распределение по типам:")
        for type_name, count in matched['Тип_совпадения'].value_counts().items():
            print(f"    • {type_name}: {count}")
    
    # Объединяем результаты
    result_final = pd.concat([matched, unmatched], ignore_index=True)
    result_final.loc[result_final['_merge'] == 'left_only', 'Тип_совпадения'] = 'Не найдено в 112'
    result_final['Период'] = period_name
    
    return result_final

def build_summary_tables(df_result):
    """Формирование сводных таблиц для отчета"""
    # Регион
    region_col = None
    if 'Регион_112' in df_result.columns:
        region_col = 'Регион_112'
    elif 'Регион_Sheets' in df_result.columns:
        region_col = 'Регион_Sheets'

    # Жалобы по регионам
    if region_col and 'Есть_жалоба' in df_result.columns:
        complaints_region = (
            df_result[(df_result['Есть_жалоба'] == True) & (df_result['Жалоба'].astype(str).str.strip() != '')]
            .groupby(region_col)
            .size()
            .reset_index(name='Количество_жалоб')
            .sort_values('Количество_жалоб', ascending=False)
        )
    else:
        complaints_region = pd.DataFrame([
            {"Примечание": "Нет данных о жалобах или регионах"}
        ])

    # Жалобы по регионам и типам
    if region_col and 'Жалоба' in df_result.columns:
        complaints_region_type = (
            df_result[(df_result['Жалоба'].astype(str).str.strip() != '')]
            .groupby([region_col, 'Жалоба'])
            .size()
            .reset_index(name='Количество')
            .sort_values(['Количество'], ascending=False)
        )
    else:
        complaints_region_type = pd.DataFrame([
            {"Примечание": "Нет данных о жалобах"}
        ])

    # Отрицательные и жалобы
    status_col = None
    if 'Статус_связи' in df_result.columns:
        status_col = 'Статус_связи'
    elif 'Статус_Sheets' in df_result.columns:
        status_col = 'Статус_Sheets'

    negative_mask = pd.Series(False, index=df_result.index)
    if status_col:
        negative_mask = df_result[status_col].astype(str).str.contains(
            r"отриц|не удалось|недозвон|не дозвон|не ответ|занят|занято|сброс",
            case=False,
            na=False
        )
    complaints_mask = df_result['Жалоба'].astype(str).str.strip() != '' if 'Жалоба' in df_result.columns else pd.Series(False, index=df_result.index)
    negative_df = df_result[negative_mask | complaints_mask].copy()

    return complaints_region, complaints_region_type, negative_df

def save_service_files(df_result, period_name, period_dir):
    """Сохранение отдельных файлов по службам"""
    if 'Служба_112' not in df_result.columns:
        return []

    safe_period = period_name.replace(':', '-').replace('/', '-')
    created_files = []
    for service_code in ['101', '102', '103', '104']:
        df_service = df_result[df_result['Служба_112'].astype(str) == service_code]
        if df_service.empty:
            continue
        file_path = period_dir / f'ОТЧЁТ_{safe_period}_СЛУЖБА_{service_code}.xlsx'
        with pd.ExcelWriter(file_path, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            complaints_region, complaints_region_type, negative_df = build_summary_tables(df_service)
            complaints_region.to_excel(writer, index=False, sheet_name='Жалобы_по_регионам')
            complaints_region_type.to_excel(writer, index=False, sheet_name='Регионы_и_жалобы')
            df_service.to_excel(writer, index=False, sheet_name='Детальные')
            negative_df.to_excel(writer, index=False, sheet_name='Отрицательные_и_жалобы')
        created_files.append(file_path)
    return created_files

def save_results(df_result, period_name):
    """Сохранение результатов"""
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    
    # Создаём папку reports если её нет
    reports_dir = Path('reports')
    reports_dir.mkdir(exist_ok=True)
    
    # Создаём подпапку для периода
    safe_period = period_name.replace(':', '-').replace('/', '-')
    period_dir = reports_dir / safe_period
    period_dir.mkdir(exist_ok=True)
    
    # CSV файл
    csv_file = period_dir / f'ОТЧЁТ_{safe_period}_{timestamp}.csv'
    print(f"\n{'='*80}")
    print("💾 СОХРАНЕНИЕ РЕЗУЛЬТАТОВ")
    print(f"{'='*80}")
    print(f"\n  Сохранение CSV...")
    df_result.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"  ✓ CSV сохранён: {csv_file}")
    
    # Excel файл - с обработкой больших данных
    excel_file = period_dir / f'ОТЧЁТ_{safe_period}_{timestamp}.xlsx'
    print(f"\n  Сохранение Excel (это может занять время для больших файлов)...")
    
    try:
        complaints_region, complaints_region_type, negative_df = build_summary_tables(df_result)
        # Для больших файлов используем xlsxwriter
        with pd.ExcelWriter(excel_file, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            complaints_region.to_excel(writer, index=False, sheet_name='Жалобы_по_регионам')
            complaints_region_type.to_excel(writer, index=False, sheet_name='Регионы_и_жалобы')
            df_result.to_excel(writer, index=False, sheet_name='Детальные')
            negative_df.to_excel(writer, index=False, sheet_name='Отрицательные_и_жалобы')
        print(f"  ✓ Excel сохранён: {excel_file}")
    except Exception as e:
        print(f"  ⚠️  Ошибка при сохранении Excel: {e}")
        print(f"  💡 Используйте CSV файл вместо Excel")
    
    print(f"\n{'='*80}")
    print("✓ РЕЗУЛЬТАТЫ СОХРАНЕНЫ")
    print(f"{'='*80}")
    print(f"\n  📄 CSV:   {csv_file}")
    print(f"  📊 Excel: {excel_file}")
    
    # Статистика
    print(f"\n{'='*80}")
    print("📊 СТАТИСТИКА")
    print(f"{'='*80}")
    
    print(f"\n  Всего записей: {len(df_result)}")
    
    if 'Тип_совпадения' in df_result.columns:
        print("\n  По типам совпадений:")
        for type_name, count in df_result['Тип_совпадения'].value_counts().items():
            print(f"    • {type_name}: {count}")
    
    if 'Служба_112' in df_result.columns:
        matched_count = df_result['Служба_112'].notna().sum()
        if matched_count > 0:
            print(f"\n  Сопоставлено с 112: {matched_count}")
            print(f"  Процент совпадений: {matched_count/len(df_result)*100:.1f}%")
            
            print("\n  По службам:")
            services = df_result[df_result['Служба_112'].notna()]['Служба_112'].value_counts()
            for service, count in list(services.items())[:10]:  # Топ 10
                print(f"    • {service}: {count}")
    
    service_files = save_service_files(df_result, period_name, period_dir)
    if service_files:
        print("\n  📌 Отдельные файлы по службам:")
        for path in service_files:
            print(f"    • {path}")

    return csv_file, excel_file

def main():
    """Основная функция"""
    print("\n" + "="*80)
    print("🔄 УНИВЕРСАЛЬНАЯ ОБРАБОТКА ДАННЫХ ПО ПЕРИОДАМ")
    print("="*80)
    print(f"📅 Дата запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    try:
        # 1. Сканируем папку 123
        periods = scan_123_folder()
        if not periods:
            return
        
        # 2. Выбираем период
        period_name, files_list = select_period(periods)
        if not files_list:
            print("\n❌ Отмена операции")
            return
        
        print(f"\n✓ Выбран период: {period_name}")
        print(f"✓ Файлов к обработке: {len(files_list)}")
        
        # 3. Загружаем данные из Google Sheets (ОБЯЗАТЕЛЬНО)
        df_sheets = load_sheets_data()
        if df_sheets.empty:
            print("\n❌ Невозможно продолжить без данных Google Sheets!")
            print("Импортируйте данные и запустите снова.")
            return
        
        # 4. Загружаем данные 112
        df_112 = load_112_data(files_list)
        if df_112.empty:
            print("\n❌ Нет данных для обработки!")
            return
        
        # 5. Сопоставляем данные
        df_result = match_data(df_sheets, df_112, period_name)
        if df_result.empty:
            print("\n❌ Нет результатов!")
            return
        
        # 6. Сохраняем результаты
        csv_file, excel_file = save_results(df_result, period_name)
        
        print(f"\n{'='*80}")
        print("✅ ГОТОВО!")
        print(f"{'='*80}")
        print("\n💡 Совет: Откройте Excel файл для просмотра результатов")
        
    except KeyboardInterrupt:
        print("\n\n❌ Операция прервана пользователем")
    except Exception as e:
        print(f"\n{'='*80}")
        print(f"❌ ОШИБКА: {e}")
        print(f"{'='*80}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
