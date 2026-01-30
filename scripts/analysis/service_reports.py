#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Расширенный отчет с данными по службам из истории звонков 112
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import asyncio
from telegram import Bot

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'
CONFIG_FILE = BASE_DIR / 'telegram_config.txt'

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def generate_service_reports():
    """Генерация отчетов по службам"""
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('ГЕНЕРАЦИЯ ОТЧЕТОВ ПО СЛУЖБАМ (102, 103, 104)')
    print('=' * 80)
    
    # Используем представление v_applications_full для объединенных данных
    query = '''
        SELECT 
            call_date,
            service_name,
            incident_number,
            card_number,
            caller_phone,
            region_name as region,
            district,
            address,
            call_reason as reason,
            application_notes as description,
            operator_name as call_operator,
            application_status as call_status,
            operator_name as fiksa_operator,
            fixation_status as fiksa_status,
            call_date as fiksa_call_date,
            citizen_phone as fiksa_phone,
            card_number as fiksa_card_number,
            notes as fiksa_notes,
            application_notes as complaint
        FROM v_applications_full
        WHERE service_name IS NOT NULL
        ORDER BY call_date DESC, incident_number
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        print('[ОШИБКА] Нет данных')
        return []
    
    print(f'[ДАННЫЕ] Всего записей: {len(df):,}')
    
    # Нормализовать статусы
    def normalize_status(status):
        if pd.isna(status) or status == '' or str(status).strip() == '':
            return 'Нет данных'
        status_lower = str(status).lower().strip()
        if 'положительн' in status_lower:
            return 'Положительный'
        elif 'отрицательн' in status_lower:
            return 'Отрицательный'
        elif 'не удалось дозвониться' in status_lower or 'недозвон' in status_lower:
            return 'Недозвонились'
        elif 'нет ответа' in status_lower or 'занято' in status_lower:
            return 'Нет ответа'
        elif 'закрыт' in status_lower or 'закрыт' in status_lower:
            return 'Закрыто'
        elif 'прерван' in status_lower or 'разорван' in status_lower:
            return 'Прервано'
        elif 'тишина' in status_lower:
            return 'Тишина'
        elif 'медработник' in status_lower or 'ходими' in status_lower or 'аризаси' in status_lower:
            return 'Медработник'
        elif 'открыт' in status_lower or 'карту' in status_lower:
            return 'Открыто'
        else:
            # Если статус неизвестный, помещаем в "Другое"
            return 'Другое'
    
    df['fiksa_status_norm'] = df['fiksa_status'].apply(normalize_status)
    df['region'] = df['region'].fillna('Не указано')
    df['service_name'] = df['service_name'].fillna('Неизвестно')
    
    # Дата для имени файла
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    date_str = datetime.now().strftime('%Y-%m-%d')
    
    report_files = []
    
    # 1. ОБЩИЙ ОТЧЕТ ПО ВСЕМ СЛУЖБАМ
    print('\n[ОТЧЕТ] Общий отчет...')
    general_file = create_general_report(df, date_str, timestamp)
    if general_file:
        report_files.append(general_file)
    
    # 2. ОТЧЕТЫ ПО КАЖДОЙ СЛУЖБЕ
    print('\n[ОТЧЕТ] Отчеты по службам...')
    services = df['service_code'].unique()
    
    for service in sorted(services):
        if service == 'Неизвестно':
            continue
        
        service_df = df[df['service_code'] == service].copy()
        service_file = create_service_report(service_df, service, date_str, timestamp)
        if service_file:
            report_files.append(service_file)
    
    # 3. ОТЧЕТЫ ПО РЕГИОНАМ ДЛЯ КАЖДОЙ СЛУЖБЫ
    print('\n[ОТЧЕТ] Отчеты по регионам и службам...')
    
    for service in sorted(services):
        if service == 'Неизвестно':
            continue
        
        service_df = df[df['service_code'] == service].copy()
        regions = service_df['region'].unique()
        
        for region in sorted(regions):
            if region == 'Не указано':
                continue
            
            region_service_df = service_df[service_df['region'] == region].copy()
            if len(region_service_df) > 0:
                filename = create_region_service_report(
                    region_service_df, region, service, date_str, timestamp
                )
                if filename:
                    report_files.append(filename)
    
    return report_files

def create_general_report(df, date_str, timestamp):
    """Создать общий отчет по всем службам"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # Статистика по службам
    service_stats = df.groupby(['service_code', 'fiksa_status_norm']).size().unstack(fill_value=0)
    service_stats['ВСЕГО'] = service_stats.sum(axis=1)
    service_stats.reset_index(inplace=True)
    
    # Статистика по регионам
    region_stats = df.groupby(['region', 'fiksa_status_norm']).size().unstack(fill_value=0)
    region_stats['ВСЕГО'] = region_stats.sum(axis=1)
    region_stats.reset_index(inplace=True)
    
    # Службы по регионам
    service_region = df.groupby(['service_code', 'region']).size().reset_index(name='Количество')
    service_region_pivot = service_region.pivot(index='region', columns='service_code', values='Количество').fillna(0)
    service_region_pivot.reset_index(inplace=True)
    
    # Детальный список с данными FIKSA
    detailed = df[[
        'call_date', 'service_code', 'region', 'district', 'incident_number',
        'caller_phone', 'reason', 'description', 'fiksa_operator', 'fiksa_call_date',
        'fiksa_phone', 'fiksa_notes', 'fiksa_status_norm'
    ]].copy()
    detailed.columns = [
        'Дата звонка', 'Служба', 'Область', 'Район', 'Номер инцидента',
        'Телефон заявителя', 'Повод', 'Описание', 'Оператор FIKSA', 'Дата фиксации',
        'Телефон FIKSA', 'Примечания FIKSA', 'Статус'
    ]
    
    # Сохранить
    output_file = OUTPUT_DIR / f'СЛУЖБЫ_ОБЩИЙ_{date_str}_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        service_stats.to_excel(writer, sheet_name='По службам', index=False)
        region_stats.to_excel(writer, sheet_name='По регионам', index=False)
        service_region_pivot.to_excel(writer, sheet_name='Службы по регионам', index=False)
        detailed.to_excel(writer, sheet_name='Детальный список', index=False)
        
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
    
    print(f'   [OK] Общий отчет: {len(detailed)} записей')
    return output_file

def create_service_report(service_df, service_code, date_str, timestamp):
    """Создать отчет по конкретной службе"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # Статистика по статусам
    status_stats = service_df.groupby('fiksa_status_norm').size().reset_index(name='Количество')
    
    # По регионам
    region_stats = service_df.groupby(['region', 'fiksa_status_norm']).size().unstack(fill_value=0)
    region_stats['ВСЕГО'] = region_stats.sum(axis=1)
    region_stats.reset_index(inplace=True)
    
    # Отрицательные жалобы
    negative = service_df[service_df['fiksa_status_norm'] == 'Отрицательный']
    if len(negative) > 0:
        neg_reasons = negative.groupby(['region', 'reason']).size().reset_index(name='Количество')
        neg_reasons = neg_reasons.sort_values(['region', 'Количество'], ascending=[True, False])
    else:
        neg_reasons = pd.DataFrame(columns=['region', 'reason', 'Количество'])
    
    # Детальный список с данными FIKSA
    detailed = service_df[[
        'call_date', 'region', 'district', 'incident_number',
        'caller_phone', 'reason', 'description', 'fiksa_operator', 'fiksa_call_date', 'fiksa_status_norm'
    ]].copy()
    detailed.columns = [
        'Дата звонка', 'Область', 'Район', 'Номер инцидента',
        'Телефон заявителя', 'Повод', 'Описание', 'Оператор FIKSA', 'Дата фиксации', 'Статус'
    ]
    
    # Сохранить
    service_name = f'Служба_{service_code}'
    output_file = OUTPUT_DIR / f'{service_name}_{date_str}_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        status_stats.to_excel(writer, sheet_name='Статистика', index=False)
        region_stats.to_excel(writer, sheet_name='По регионам', index=False)
        neg_reasons.to_excel(writer, sheet_name='Отрицательные', index=False)
        detailed.to_excel(writer, sheet_name='Детальный список', index=False)
        
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
    
    print(f'   [OK] Служба {service_code}: {len(detailed)} записей')
    return output_file

def create_region_service_report(region_service_df, region, service_code, date_str, timestamp):
    """Создать отчет по региону и службе"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # Статистика
    status_stats = region_service_df.groupby('fiksa_status_norm').size().reset_index(name='Количество')
    
    # Отрицательные
    negative = region_service_df[region_service_df['fiksa_status_norm'] == 'Отрицательный']
    if len(negative) > 0:
        neg_reasons = negative.groupby('reason').size().reset_index(name='Количество')
        neg_reasons = neg_reasons.sort_values('Количество', ascending=False)
    else:
        neg_reasons = pd.DataFrame(columns=['reason', 'Количество'])
    
    # Детальный список с данными FIKSA
    detailed = region_service_df[[
        'call_date', 'district', 'incident_number',
        'caller_phone', 'reason', 'description', 'fiksa_operator', 'fiksa_call_date', 'fiksa_status_norm'
    ]].copy()
    detailed.columns = [
        'Дата звонка', 'Район', 'Номер инцидента',
        'Телефон заявителя', 'Повод', 'Описание', 'Оператор FIKSA', 'Дата фиксации', 'Статус'
    ]
    
    # Сохранить
    output_file = OUTPUT_DIR / f'{region}_Служба_{service_code}_{date_str}_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        status_stats.to_excel(writer, sheet_name='Статистика', index=False)
        neg_reasons.to_excel(writer, sheet_name='Отрицательные', index=False)
        detailed.to_excel(writer, sheet_name='Детальный список', index=False)
        
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
    
    print(f'      [OK] {region} - Служба {service_code}: {len(detailed)} записей')
    return output_file

def send_to_telegram(file_path):
    """Отправить файл в Telegram"""
    try:
        if not CONFIG_FILE.exists():
            return False
        
        config = {}
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    config[key] = value
        
        if 'BOT_TOKEN' not in config or 'CHAT_ID' not in config:
            return False
        
        if config['BOT_TOKEN'] == 'your_bot_token_here':
            return False
        
        async def send():
            bot = Bot(token=config['BOT_TOKEN'])
            with open(file_path, 'rb') as f:
                await bot.send_document(
                    chat_id=config['CHAT_ID'], 
                    document=f,
                    filename=file_path.name,
                    caption=f'[ОТЧЕТ] Отчет по службам\n[ВРЕМЯ] {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
                )
        
        asyncio.run(send())
        return True
        
    except Exception as e:
        print(f'[ОШИБКА] Telegram: {e}')
        return False

if __name__ == '__main__':
    report_files = generate_service_reports()
    
    if report_files:
        print(f'\n[OK] СОЗДАНО {len(report_files)} ФАЙЛОВ')
        
        # Отправить общий отчет в Telegram
        general_report = [f for f in report_files if 'ОБЩИЙ' in str(f)]
        if general_report:
            print('\n[TELEGRAM] Отправка в Telegram...')
            if send_to_telegram(general_report[0]):
                print('[OK] Отправлено')
