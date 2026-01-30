#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Расширенный отчет по адресам с агрегацией и отправкой в Telegram
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import re

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'
CONFIG_FILE = BASE_DIR / 'telegram_config.txt'

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

def aggregate_data_to_db():
    """Агрегировать данные для быстрого доступа"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    print('📊 Агрегация данных...')
    
    # Создать таблицу агрегированных данных
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS aggregated_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            region TEXT,
            district TEXT,
            status TEXT,
            complaint_type TEXT,
            count INTEGER,
            last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Очистить старые данные
    cursor.execute('DELETE FROM aggregated_reports')
    
    # Получить данные с адресами
    query = '''
        SELECT 
            a.address,
            a.notes as complaint,
            f.status
        FROM applications a
        LEFT JOIN fiksa_records f ON (
            f.full_name = a.application_number 
            OR f.phone LIKE '%' || REPLACE(REPLACE(a.phone, '+998', ''), '+', '') || '%'
        )
        WHERE a.address IS NOT NULL AND a.address != ''
    '''
    
    df = pd.read_sql_query(query, conn)
    
    # Разобрать адреса
    df[['region', 'district']] = df['address'].apply(
        lambda x: pd.Series(parse_address(x))
    )
    
    # Упростить жалобы (первые 50 символов как тип)
    df['complaint_type'] = df['complaint'].fillna('Не указано').str[:50]
    df['status'] = df['status'].fillna('Нет данных')
    
    # Агрегировать
    agg_data = df.groupby(['region', 'district', 'status', 'complaint_type']).size().reset_index(name='count')
    
    # Записать в БД
    agg_data.to_sql('aggregated_reports', conn, if_exists='append', index=False)
    
    conn.commit()
    conn.close()
    
    print(f'✅ Агрегировано {len(agg_data)} записей')

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

def generate_advanced_report(start_date=None, end_date=None, selected_regions=None):
    """Создать расширенный отчет по адресам"""
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('РАСШИРЕННЫЙ ОТЧЕТ ПО АДРЕСАМ')
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
    
    # Получить данные
    query = f'''
        SELECT 
            a.application_number,
            a.phone,
            a.address,
            a.notes as complaint,
            a.import_date,
            f.status,
            f.call_date
        FROM applications a
        LEFT JOIN fiksa_records f ON (
            f.full_name = a.application_number 
            OR f.phone LIKE '%' || REPLACE(REPLACE(a.phone, '+998', ''), '+', '') || '%'
        )
        WHERE {where_clause}
        ORDER BY a.import_date DESC, a.application_number
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        print('❌ Нет данных')
        return None
    
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
        return None
    
    # Заполнить пустые статусы
    df['status'] = df['status'].fillna('Нет данных')
    df['complaint'] = df['complaint'].fillna('Не указано')
    
    # 1. ПО РЕГИОНАМ С РАЗВЕРНУТЫМИ СТАТУСАМИ
    print('\n📍 Статистика по регионам...')
    
    # Получить все уникальные статусы
    all_statuses = df['status'].unique()
    
    # Создать сводную таблицу: регион × статус
    region_status_pivot = pd.crosstab(df['Область'], df['status'], margins=True, margins_name='ИТОГО')
    region_status_pivot.reset_index(inplace=True)
    
    # 2. ПО РАЙОНАМ С РАЗВЕРНУТЫМИ СТАТУСАМИ
    print('📍 Статистика по районам...')
    district_status_pivot = pd.crosstab(
        [df['Область'], df['Район']], 
        df['status'], 
        margins=True, 
        margins_name='ИТОГО'
    )
    district_status_pivot.reset_index(inplace=True)
    
    # 3. ТОП ЖАЛОБ ПО РЕГИОНАМ
    print('📝 Топ жалоб...')
    complaint_stats = df.groupby(['Область', 'complaint']).size().reset_index(name='Количество')
    complaint_stats = complaint_stats.sort_values(['Область', 'Количество'], ascending=[True, False])
    
    # 4. ДЕТАЛЬНЫЙ СПИСОК (БЕЗ ОПЕРАТОРА)
    print('📋 Детальный список...')
    detailed = df[[
        'import_date', 'Область', 'Район', 'application_number', 
        'phone', 'complaint', 'status', 'call_date'
    ]].copy()
    detailed.columns = [
        'Дата импорта', 'Область', 'Район', 'Номер заявки', 
        'Телефон', 'Описание жалобы', 'Статус', 'Дата звонка'
    ]
    detailed = detailed.sort_values('Дата импорта', ascending=False)
    
    # Сохранить в Excel
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    filename_parts = ['address_advanced']
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
    
    print(f'\n💾 Сохранение отчета: {output_file.name}')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Лист 1: Статусы по регионам (развернуто)
        region_status_pivot.to_excel(writer, sheet_name='Статусы по регионам', index=False)
        
        # Лист 2: Статусы по районам (развернуто)
        district_status_pivot.to_excel(writer, sheet_name='Статусы по районам', index=False)
        
        # Лист 3: Топ жалоб по регионам
        complaint_stats.to_excel(writer, sheet_name='Топ жалоб', index=False)
        
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
    print(f'📂 Файл: {output_file.name}')
    print(f'📊 Регионов: {len(region_status_pivot) - 1}')  # -1 для строки ИТОГО
    print(f'📍 Районов: {len(district_status_pivot) - 1}')
    print(f'📝 Записей: {len(detailed)}')
    
    return output_file

def send_to_telegram(file_path, message=None):
    """Отправить файл в Telegram"""
    try:
        # Проверить конфиг
        if not CONFIG_FILE.exists():
            print('\n⚠️ Telegram не настроен')
            print(f'Создайте файл {CONFIG_FILE} с двумя строками:')
            print('BOT_TOKEN=your_bot_token')
            print('CHAT_ID=your_chat_id')
            return False
        
        # Прочитать конфиг
        config = {}
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    config[key] = value
        
        if 'BOT_TOKEN' not in config or 'CHAT_ID' not in config:
            print('❌ Неверный формат конфига')
            return False
        
        print('\n📤 Отправка в Telegram...')
        
        import asyncio
        from telegram import Bot
        
        async def send():
            bot = Bot(token=config['BOT_TOKEN'])
            
            # Отправить сообщение
            if message:
                await bot.send_message(chat_id=config['CHAT_ID'], text=message)
            
            # Отправить файл
            with open(file_path, 'rb') as f:
                await bot.send_document(
                    chat_id=config['CHAT_ID'], 
                    document=f,
                    filename=file_path.name,
                    caption=f'📊 Отчет готов: {file_path.name}'
                )
        
        asyncio.run(send())
        print('✅ Отправлено в Telegram')
        return True
        
    except Exception as e:
        print(f'❌ Ошибка отправки: {e}')
        return False

def interactive_menu():
    """Интерактивное меню"""
    print('\n' + '=' * 80)
    print('ГЕНЕРАТОР РАСШИРЕННЫХ ОТЧЕТОВ ПО АДРЕСАМ')
    print('=' * 80)
    
    # Агрегация данных
    print('\n🔄 Обновление агрегированных данных...')
    aggregate_data_to_db()
    
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
    output_file = generate_advanced_report(start_date, end_date, selected_regions)
    
    if output_file:
        # Отправка в Telegram
        print('\n📱 ОТПРАВКА В TELEGRAM:')
        print('1. Да')
        print('2. Нет')
        
        telegram_choice = input('\nОтправить в Telegram? (1-2): ').strip()
        
        if telegram_choice == '1':
            message = f'📊 Отчет по адресам\n'
            if start_date or end_date:
                message += f'📅 Период: {start_date or "..."} - {end_date or "..."}\n'
            if selected_regions:
                message += f'🌍 Регионы: {", ".join(selected_regions)}\n'
            
            send_to_telegram(output_file, message)

if __name__ == '__main__':
    interactive_menu()
