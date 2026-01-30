#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Автоматическая генерация отчета без интерактивного меню
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import sys

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

def parse_complaint(complaint_text):
    """Разобрать описание жалобы на категорию и описание"""
    if not complaint_text or pd.isna(complaint_text):
        return 'Не указано', 'Не указано', ''
    
    text = str(complaint_text)
    
    # Если есть двоеточие - первая часть категория
    if ':' in text:
        parts = text.split(':', 1)
        category = parts[0].strip()
        description = parts[1].strip() if len(parts) > 1 else ''
    else:
        category = 'Общая'
        description = text
    
    # Дополнительная информация (после "Дата:")
    additional_info = ''
    if 'Дата:' in description:
        parts = description.split('Дата:', 1)
        description = parts[0].strip()
        additional_info = 'Дата: ' + parts[1].strip()
    
    return category, description, additional_info

def extract_service_from_incident(incident_number):
    """Извлечь код службы из номера инцидента"""
    if not incident_number or pd.isna(incident_number):
        return 'Неизвестно'
    
    # Формат: 01.AAC4685/26
    # Код службы - это часть между точкой и цифрами
    match = str(incident_number)
    if '.' in match and '/' in match:
        parts = match.split('.')
        if len(parts) > 1:
            service_code = parts[1][:3]  # AAC, например
            return service_code
    
    return 'Неизвестно'

def generate_auto_report():
    """Автоматическая генерация отчета за все время"""
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('АВТОМАТИЧЕСКАЯ ГЕНЕРАЦИЯ ОТЧЕТА')
    print('=' * 80)
    
    # Получить все данные
    query = '''
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
        WHERE a.address IS NOT NULL AND a.address != ''
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
    
    # Разобрать жалобы
    df[['Категория жалобы', 'Описание', 'Дополнительная информация']] = df['complaint'].apply(
        lambda x: pd.Series(parse_complaint(x))
    )
    
    # Извлечь службу
    df['Служба'] = df['application_number'].apply(extract_service_from_incident)
    
    # Заполнить пустые значения
    df['status'] = df['status'].fillna('Нет данных')
    df['call_date'] = df['call_date'].fillna('Не указано')
    
    # Нормализовать статусы
    def normalize_status(status):
        status_lower = str(status).lower()
        if 'положительн' in status_lower:
            return 'Положительный'
        elif 'отрицательн' in status_lower:
            return 'Отрицательный'
        elif 'не удалось дозвониться' in status_lower or 'недозвон' in status_lower:
            return 'Недозвонились'
        elif 'нет ответа' in status_lower or 'занято' in status_lower:
            return 'Нет ответа'
        else:
            return status
    
    df['Статус_норм'] = df['status'].apply(normalize_status)
    
    # 1. ОБЩАЯ СТАТИСТИКА ПО РЕГИОНАМ
    print('📊 Общая статистика по регионам...')
    region_stats = df.groupby(['Область', 'Статус_норм']).size().unstack(fill_value=0)
    region_stats['ВСЕГО'] = region_stats.sum(axis=1)
    region_stats.reset_index(inplace=True)
    
    # 2. СТАТИСТИКА ПО СЛУЖБАМ
    print('🚑 Статистика по службам...')
    service_stats = df.groupby(['Служба', 'Статус_норм']).size().unstack(fill_value=0)
    service_stats['ВСЕГО'] = service_stats.sum(axis=1)
    service_stats.reset_index(inplace=True)
    
    # 3. ПО СЛУЖБАМ И РЕГИОНАМ
    print('🌍 По службам и регионам...')
    service_region = df.groupby(['Служба', 'Область', 'Статус_норм']).size().unstack(fill_value=0)
    service_region['ВСЕГО'] = service_region.sum(axis=1)
    service_region.reset_index(inplace=True)
    
    # 4. ОТРИЦАТЕЛЬНЫЕ ЖАЛОБЫ ПО РЕГИОНАМ
    print('❌ Отрицательные жалобы...')
    negative = df[df['Статус_норм'] == 'Отрицательный']
    negative_complaints = negative.groupby(['Область', 'Категория жалобы']).size().reset_index(name='Количество')
    negative_complaints = negative_complaints.sort_values(['Область', 'Количество'], ascending=[True, False])
    
    # 5. ДЕТАЛЬНЫЙ СПИСОК
    print('📋 Детальный список...')
    detailed = df[[
        'call_date', 'Область', 'Район', 'application_number', 
        'phone', 'Категория жалобы', 'Описание', 'Дополнительная информация', 'status'
    ]].copy()
    detailed.columns = [
        'Дата звонка', 'Область', 'Район', 'Номер заявки', 
        'Телефон', 'Категория жалобы', 'Описание жалобы', 'Дополнительная информация', 'Статус'
    ]
    # Сортировка по номеру заявки вместо даты
    detailed = detailed.sort_values('Номер заявки', ascending=False)
    
    # Сохранить отчет
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    date_str = datetime.now().strftime('%Y-%m-%d')
    
    # Общий отчет
    output_file = OUTPUT_DIR / f'ОБЩИЙ_ОТЧЕТ_{date_str}_{timestamp}.xlsx'
    
    print(f'\n💾 Сохранение общего отчета: {output_file.name}')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        region_stats.to_excel(writer, sheet_name='Статистика по регионам', index=False)
        service_stats.to_excel(writer, sheet_name='По службам', index=False)
        service_region.to_excel(writer, sheet_name='Службы по регионам', index=False)
        negative_complaints.to_excel(writer, sheet_name='Отрицательные жалобы', index=False)
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
    
    print('\n✅ ОБЩИЙ ОТЧЕТ ГОТОВ!')
    print(f'📂 {output_file}')
    print(f'📊 Регионов: {len(region_stats)}')
    print(f'📝 Записей: {len(detailed)}')
    
    # Создать отчеты по каждому региону
    print('\n📍 Создание отчетов по регионам...')
    
    regions = df['Область'].unique()
    region_files = []
    
    for region in regions:
        if region == 'Не указано':
            continue
        
        region_df = df[df['Область'] == region].copy()
        
        # Статистика по статусам
        region_status = region_df.groupby('Статус_норм').size().reset_index(name='Количество')
        
        # По службам
        region_service = region_df.groupby(['Служба', 'Статус_норм']).size().unstack(fill_value=0)
        region_service['ВСЕГО'] = region_service.sum(axis=1)
        region_service.reset_index(inplace=True)
        
        # Отрицательные
        region_negative = region_df[region_df['Статус_норм'] == 'Отрицательный']
        region_neg_complaints = region_negative.groupby('Категория жалобы').size().reset_index(name='Количество')
        region_neg_complaints = region_neg_complaints.sort_values('Количество', ascending=False)
        
        # Детальный список
        region_detailed = region_df[[
            'call_date', 'Район', 'application_number', 
            'phone', 'Категория жалобы', 'Описание', 'Дополнительная информация', 'status'
        ]].copy()
        region_detailed.columns = [
            'Дата звонка', 'Район', 'Номер заявки', 
            'Телефон', 'Категория жалобы', 'Описание жалобы', 'Дополнительная информация', 'Статус'
        ]
        region_detailed = region_detailed.sort_values('Номер заявки', ascending=False)
        
        # Сохранить
        region_filename = OUTPUT_DIR / f'{region}_{date_str}_{timestamp}.xlsx'
        
        with pd.ExcelWriter(region_filename, engine='openpyxl') as writer:
            region_status.to_excel(writer, sheet_name='Статистика', index=False)
            region_service.to_excel(writer, sheet_name='По службам', index=False)
            region_neg_complaints.to_excel(writer, sheet_name='Отрицательные жалобы', index=False)
            region_detailed.to_excel(writer, sheet_name='Детальный список', index=False)
            
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
        
        region_files.append(region_filename)
        print(f'  ✅ {region}: {len(region_detailed)} записей')
    
    return output_file, region_files

def send_to_telegram(file_path):
    """Отправить файл в Telegram"""
    try:
        if not CONFIG_FILE.exists():
            print('\n⚠️ Telegram не настроен')
            return False
        
        config = {}
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    config[key] = value
        
        if 'BOT_TOKEN' not in config or 'CHAT_ID' not in config:
            print('❌ Неверный конфиг Telegram')
            return False
        
        if config['BOT_TOKEN'] == 'your_bot_token_here':
            print('⚠️ Telegram не настроен (токен не указан)')
            return False
        
        print('\n📤 Отправка в Telegram...')
        
        import asyncio
        from telegram import Bot
        
        async def send():
            bot = Bot(token=config['BOT_TOKEN'])
            
            with open(file_path, 'rb') as f:
                await bot.send_document(
                    chat_id=config['CHAT_ID'], 
                    document=f,
                    filename=file_path.name,
                    caption=f'📊 Автоматический отчет\n🕐 {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
                )
        
        asyncio.run(send())
        print('✅ Отправлено в Telegram')
        return True
        
    except Exception as e:
        print(f'❌ Ошибка Telegram: {e}')
        return False

if __name__ == '__main__':
    # Генерация отчета
    result = generate_auto_report()
    
    if result:
        output_file, region_files = result
        
        # Автоматическая отправка общего отчета в Telegram
        print('\n📱 Отправка общего отчета в Telegram...')
        send_to_telegram(output_file)
        
        print(f'\n✅ ГОТОВО! Создано {len(region_files) + 1} файлов')
        print(f'   - 1 общий отчет')
        print(f'   - {len(region_files)} отчетов по регионам')
