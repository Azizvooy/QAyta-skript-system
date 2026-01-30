#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Отчет по адресам: регионы, районы, статусы, жалобы
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import re

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def parse_address(address):
    """Разбить адрес на область и район"""
    if not address or pd.isna(address):
        return 'Не указано', 'Не указано'
    
    parts = [p.strip() for p in str(address).split(',')]
    
    if len(parts) >= 2:
        region = parts[0]
        district = parts[1]
    elif len(parts) == 1:
        region = parts[0]
        district = 'Не указано'
    else:
        region = 'Не указано'
        district = 'Не указано'
    
    return region, district

def get_available_regions():
    """Получить список доступных регионов"""
    conn = get_db_connection()
    query = "SELECT DISTINCT address FROM applications WHERE address IS NOT NULL AND address != ''"
    addresses = pd.read_sql_query(query, conn)
    conn.close()
    
    regions = set()
    for addr in addresses['address']:
        region, _ = parse_address(addr)
        if region != 'Не указано':
            regions.add(region)
    
    return sorted(list(regions))

def generate_address_report(start_date=None, end_date=None, selected_regions=None):
    """Создать отчет по адресам"""
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('ОТЧЕТ ПО АДРЕСАМ: РЕГИОНЫ, РАЙОНЫ, СТАТУСЫ, ЖАЛОБЫ')
    print('=' * 80)
    
    # Построить WHERE условия
    conditions = ["a.address IS NOT NULL AND a.address != ''"]
    
    if start_date:
        conditions.append(f"DATE(a.import_date) >= '{start_date}'")
        print(f'📅 Период с: {start_date}')
    
    if end_date:
        conditions.append(f"DATE(a.import_date) <= '{end_date}'")
        print(f'📅 Период по: {end_date}')
    
    where_clause = " AND ".join(conditions)
    
    # Получить все заявки с их статусами из FIKSA
    query = f'''
        SELECT 
            a.application_number,
            a.phone,
            a.address,
            a.notes as complaint,
            a.import_date,
            f.operator_name,
            f.status,
            f.call_date
        FROM applications a
        LEFT JOIN fiksa_records f ON (
            f.full_name = a.application_number 
            OR f.phone LIKE '%' || REPLACE(REPLACE(a.phone, '+998', ''), '+', '') || '%'
        )
        WHERE {where_clause}
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        print('❌ Нет данных')
        return
    
    print(f'📊 Всего записей: {len(df)}')
    
    # Разобрать адреса
    df[['Область', 'Район']] = df['address'].apply(
        lambda x: pd.Series(parse_address(x))
    )
    
    # Фильтр по регионам
    if selected_regions:
        df = df[df['Область'].isin(selected_regions)]
        print(f'🌍 Фильтр по регионам: {", ".join(selected_regions)}')
    
    if df.empty:
        print('❌ Нет данных по выбранным критериям')
        return
    
    # 1. СВОДКА ПО РЕГИОНАМ
    print('\n📍 Группировка по регионам...')
    region_stats = df.groupby('Область').agg({
        'application_number': 'count',
        'status': lambda x: x.value_counts().to_dict() if x.notna().any() else {}
    }).reset_index()
    region_stats.columns = ['Область', 'Всего заявок', 'Статусы']
    
    # 2. ДЕТАЛИЗАЦИЯ ПО РАЙОНАМ
    print('📍 Группировка по районам...')
    district_stats = df.groupby(['Область', 'Район']).agg({
        'application_number': 'count',
        'status': lambda x: x.value_counts().to_dict() if x.notna().any() else {}
    }).reset_index()
    district_stats.columns = ['Область', 'Район', 'Всего заявок', 'Статусы']
    
    # 3. ПОЛНЫЙ СПИСОК С ЖАЛОБАМИ
    print('📝 Формирование детального списка...')
    detailed = df[[
        'Область', 'Район', 'application_number', 'phone', 
        'complaint', 'status', 'operator_name', 'call_date'
    ]].copy()
    detailed.columns = [
        'Область', 'Район', 'Номер заявки', 'Телефон',
        'Описание жалобы', 'Статус', 'Оператор', 'Дата звонка'
    ]
    
    # 4. СТАТИСТИКА ПО СТАТУСАМ
    print('📊 Статистика по статусам...')
    status_by_region = df[df['status'].notna()].groupby(['Область', 'status']).size().reset_index()
    status_by_region.columns = ['Область', 'Статус', 'Количество']
    status_pivot = status_by_region.pivot(index='Область', columns='Статус', values='Количество').fillna(0)
    status_pivot.reset_index(inplace=True)
    
    # Сохранить в Excel
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Добавить дату в имя файла если выбран период
    filename_parts = ['address_report']
    if start_date and end_date:
        filename_parts.append(f'{start_date}_to_{end_date}')
    elif start_date:
        filename_parts.append(f'from_{start_date}')
    elif end_date:
        filename_parts.append(f'to_{end_date}')
    if selected_regions and len(selected_regions) <= 3:
        region_str = '_'.join([r[:10] for r in selected_regions])
        filename_parts.append(region_str)
    filename_parts.append(timestamp)
    
    output_file = OUTPUT_DIR / f'{"_".join(filename_parts)}.xlsx'
    
    print(f'\n💾 Сохранение отчета: {output_file}')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Лист 1: Сводка по регионам
        region_summary = region_stats.copy()
        region_summary['Статусы'] = region_summary['Статусы'].apply(
            lambda x: ', '.join([f'{k}: {v}' for k, v in x.items()]) if isinstance(x, dict) and x else 'Нет данных'
        )
        region_summary.to_excel(writer, sheet_name='По регионам', index=False)
        
        # Лист 2: Детализация по районам
        district_summary = district_stats.copy()
        district_summary['Статусы'] = district_summary['Статусы'].apply(
            lambda x: ', '.join([f'{k}: {v}' for k, v in x.items()]) if isinstance(x, dict) and x else 'Нет данных'
        )
        district_summary.to_excel(writer, sheet_name='По районам', index=False)
        
        # Лист 3: Статусы по регионам (таблица)
        status_pivot.to_excel(writer, sheet_name='Статусы по регионам', index=False)
        
        # Лист 4: Детальный список
        detailed.to_excel(writer, sheet_name='Детальный список', index=False)
        
        # Форматирование
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print('\n✅ ОТЧЕТ ГОТОВ!\n')
    print(f'📂 Файл: {output_file}')
    print(f'📊 Регионов: {len(region_stats)}')
    print(f'📍 Районов: {len(district_stats)}')
    print(f'📝 Записей: {len(detailed)}')
    
    # Показать топ регионов
    print('\n🏆 Топ-5 регионов по количеству заявок:')
    top_regions = region_summary.nlargest(5, 'Всего заявок')
    for idx, row in top_regions.iterrows():
        print(f'   {row["Область"]}: {row["Всего заявок"]} заявок')

def interactive_menu():
    """Интерактивное меню"""
    print('\n' + '=' * 80)
    print('ГЕНЕРАТОР ОТЧЕТОВ ПО АДРЕСАМ')
    print('=' * 80)
    
    # Выбор периода
    print('\n📅 ВЫБОР ПЕРИОДА:')
    print('1. Все даты')
    print('2. За последний день')
    print('3. За последнюю неделю')
    print('4. За последний месяц')
    print('5. Указать свой период')
    
    choice = input('\nВыберите (1-5): ').strip()
    
    start_date = None
    end_date = None
    
    if choice == '2':
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = end_date
    elif choice == '3':
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - pd.Timedelta(days=7)).strftime('%Y-%m-%d')
    elif choice == '4':
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - pd.Timedelta(days=30)).strftime('%Y-%m-%d')
    elif choice == '5':
        start_date = input('Дата начала (YYYY-MM-DD, Enter для пропуска): ').strip() or None
        end_date = input('Дата окончания (YYYY-MM-DD, Enter для пропуска): ').strip() or None
    
    # Выбор регионов
    print('\n🌍 ВЫБОР РЕГИОНОВ:')
    print('1. Все регионы')
    print('2. Выбрать конкретные регионы')
    
    region_choice = input('\nВыберите (1-2): ').strip()
    
    selected_regions = None
    
    if region_choice == '2':
        print('\nДоступные регионы:')
        regions = get_available_regions()
        for i, region in enumerate(regions, 1):
            print(f'{i:2d}. {region}')
        
        region_input = input('\nВведите номера регионов через запятую (например: 1,3,5): ').strip()
        if region_input:
            try:
                indices = [int(x.strip()) - 1 for x in region_input.split(',')]
                selected_regions = [regions[i] for i in indices if 0 <= i < len(regions)]
            except:
                print('⚠️ Некорректный ввод, выбраны все регионы')
    
    # Генерация отчета
    generate_address_report(start_date, end_date, selected_regions)

if __name__ == '__main__':
    interactive_menu()
