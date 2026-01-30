#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
АНАЛИЗ СТАТИСТИКИ ОПЕРАТОРОВ И ФИДБЭКОВ ОТ СЛУЖБ
=============================================================================
Генерирует детальные отчеты по:
- Работе операторов
- Фидбэкам от служб (102, 103, 104)
- Ответам граждан
- Качеству обслуживания
=============================================================================
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

# Паттерны для исключения из отчетов (имена операторов, которые нужно пропускать)
EXCLUDE_PATTERNS = [
    'Тренды',
    'Текущий месяц - Сводка',
    'Предыдущий месяц - Сводка',
    'СВОДКА СОТРУДНИКИ',
    'Ноябрь 2025',
    'Декабрь 2025',
    'сводка',  # любые сводки
    'итого',   # любые итоги
]

# =============================================================================
# СТАТИСТИКА ОПЕРАТОРОВ
# =============================================================================

def operator_performance_report():
    """Детальный отчет по работе операторов"""
    conn = sqlite3.connect(DB_PATH)
    
    query = '''
        SELECT 
            operator_name as "Оператор",
            COUNT(*) as "Всего звонков",
            COUNT(CASE WHEN fixation_status LIKE '%Положительн%' THEN 1 END) as "Положительных",
            COUNT(CASE WHEN fixation_status LIKE '%Отрицательн%' THEN 1 END) as "Отрицательных",
            COUNT(CASE WHEN fixation_status LIKE '%Недозвон%' THEN 1 END) as "Недозвонились",
            COUNT(CASE WHEN fixation_status LIKE '%Нет ответа%' OR fixation_status LIKE '%занято%' THEN 1 END) as "Нет ответа",
            ROUND(COUNT(CASE WHEN fixation_status LIKE '%Положительн%' THEN 1 END) * 100.0 / COUNT(*), 1) as "% Положительных",
            ROUND(COUNT(CASE WHEN fixation_status LIKE '%Отрицательн%' THEN 1 END) * 100.0 / COUNT(*), 1) as "% Отрицательных"
        FROM v_fixations_full
        WHERE call_date = date('now') AND fixation_status IS NOT NULL AND fixation_status != ''
        GROUP BY operator_name
        ORDER BY COUNT(*) DESC
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    # Фильтруем исключаемые строки
    if not df.empty and 'Оператор' in df.columns:
        for pattern in EXCLUDE_PATTERNS:
            df = df[~df['Оператор'].str.contains(pattern, case=False, na=False)]
    
    # Сохранить в Excel
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'СТАТИСТИКА_ОПЕРАТОРОВ_{timestamp}.xlsx'
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Статистика операторов', index=False)
        
        # Форматирование
        worksheet = writer.sheets['Статистика операторов']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    cell_value = str(cell.value) if cell.value is not None else ''
                    if len(cell_value) > max_length:
                        max_length = len(cell_value)
                except Exception:
                    pass
            adjusted_width = min(max(max_length + 2, 10), 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f'[OK] Отчет по операторам: {len(df)} операторов')
    return output_file

# =============================================================================
# АНАЛИЗ ФИДБЭКОВ ОТ СЛУЖБ
# =============================================================================

def service_feedback_report():
    """Детальный отчет по фидбэкам от служб"""
    conn = sqlite3.connect(DB_PATH)
    
    # Общая статистика по службам из представления
    query_general = '''
        SELECT 
            service_name as "Служба",
            COUNT(*) as "Всего обращений",
            COUNT(CASE WHEN fixation_status IS NOT NULL AND fixation_status != '' THEN 1 END) as "С фидбэком",
            COUNT(CASE WHEN fixation_status LIKE '%Положительн%' THEN 1 END) as "Положительных",
            COUNT(CASE WHEN fixation_status LIKE '%Отрицательн%' THEN 1 END) as "Отрицательных",
            ROUND(COUNT(CASE WHEN fixation_status LIKE '%Положительн%' THEN 1 END) * 100.0 / 
                  NULLIF(COUNT(CASE WHEN fixation_status IS NOT NULL AND fixation_status != '' THEN 1 END), 0), 1) as "% Положительных"
        FROM v_applications_full
        WHERE service_name IS NOT NULL
        GROUP BY service_name
        ORDER BY COUNT(*) DESC
    '''
    
    df_general = pd.read_sql_query(query_general, conn)
    
    # Детальная разбивка по типам фидбэков
    query_detailed = '''
        SELECT 
            service_name as "Служба",
            fixation_status as "Тип фидбэка",
            COUNT(*) as "Количество",
            ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY service_name), 1) as "% от службы"
        FROM v_applications_full
        WHERE service_name IS NOT NULL AND fixation_status IS NOT NULL AND fixation_status != ''
        GROUP BY service_name, fixation_status
        ORDER BY service_name, COUNT(*) DESC
    '''
    
    df_detailed = pd.read_sql_query(query_detailed, conn)
    
    # Проблемные случаи (отрицательные фидбэки)
    query_problems = '''
        SELECT 
            service_name as "Служба",
            region_name as "Регион",
            incident_number as "Номер инцидента",
            call_reason as "Повод",
            operator_name as "Оператор",
            fixation_status as "Фидбэк",
            notes as "Примечания"
        FROM v_applications_full
        WHERE fixation_status LIKE '%Отрицательн%'
        ORDER BY service_name, region_name
        LIMIT 500
    '''
    
    df_problems = pd.read_sql_query(query_problems, conn)
    conn.close()
    
    # Сохранить в Excel
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'ФИДБЭКИ_ОТ_СЛУЖБ_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_general.to_excel(writer, sheet_name='Общая статистика', index=False)
        df_detailed.to_excel(writer, sheet_name='Детальная разбивка', index=False)
        df_problems.to_excel(writer, sheet_name='Проблемные случаи', index=False)
        
        # Форматирование всех листов
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        cell_value = str(cell.value) if cell.value is not None else ''
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except Exception:
                        pass
                adjusted_width = min(max(max_length + 2, 10), 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f'[OK] Отчет по фидбэкам: {len(df_general)} служб, {len(df_problems)} проблемных случаев')
    return output_file

# =============================================================================
# АНАЛИЗ ОТВЕТОВ ГРАЖДАН
# =============================================================================

def citizen_response_analysis():
    """Анализ ответов граждан на обращения"""
    conn = sqlite3.connect(DB_PATH)
    
    # Статистика по типам ответов из представления
    query = '''
        SELECT 
            operator_name as "Оператор",
            fixation_status as "Статус",
            COUNT(*) as "Количество",
            ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (PARTITION BY operator_name), 1) as "% от оператора"
        FROM v_fixations_full
        WHERE call_date = date('now') AND fixation_status IS NOT NULL AND fixation_status != ''
        GROUP BY operator_name, fixation_status
        ORDER BY operator_name, COUNT(*) DESC
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    # Фильтруем исключаемые строки
    if not df.empty and 'Оператор' in df.columns:
        for pattern in EXCLUDE_PATTERNS:
            df = df[~df['Оператор'].str.contains(pattern, case=False, na=False)]
    
    # Сохранить
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'ОТВЕТЫ_ГРАЖДАН_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Ответы граждан', index=False)
        
        worksheet = writer.sheets['Ответы граждан']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    cell_value = str(cell.value) if cell.value is not None else ''
                    if len(cell_value) > max_length:
                        max_length = len(cell_value)
                except Exception:
                    pass
            adjusted_width = min(max(max_length + 2, 10), 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f'[OK] Анализ ответов граждан: {len(df)} записей')
    return output_file

# =============================================================================
# ОТПРАВКА В TELEGRAM
# =============================================================================

def send_to_telegram(file_path, caption):
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
        
        async def send():
            bot = Bot(token=config['BOT_TOKEN'])
            with open(file_path, 'rb') as f:
                await bot.send_document(
                    chat_id=config['CHAT_ID'],
                    document=f,
                    filename=file_path.name,
                    caption=caption
                )
        
        asyncio.run(send())
        return True
        
    except Exception as e:
        print(f'[ОШИБКА] Telegram: {e}')
        return False

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    """Генерация всех аналитических отчетов"""
    print('\n' + '=' * 80)
    print('ГЕНЕРАЦИЯ АНАЛИТИЧЕСКИХ ОТЧЕТОВ')
    print('=' * 80)
    
    reports = []
    
    # 1. Статистика операторов
    print('\n[1/3] Статистика операторов...')
    operator_file = operator_performance_report()
    reports.append(('Статистика операторов', operator_file))
    
    # 2. Фидбэки от служб
    print('\n[2/3] Фидбэки от служб...')
    feedback_file = service_feedback_report()
    reports.append(('Фидбэки от служб', feedback_file))
    
    # 3. Ответы граждан
    print('\n[3/3] Ответы граждан...')
    citizen_file = citizen_response_analysis()
    reports.append(('Ответы граждан', citizen_file))
    
    # Отправка в Telegram
    print('\n[TELEGRAM] Отправка отчетов...')
    for caption, file in reports:
        if send_to_telegram(file, f'[АНАЛИТИКА] {caption}'):
            print(f'   [OK] {caption}')
        else:
            print(f'   [ПРОПУЩЕНО] {caption}')
    
    print('\n' + '=' * 80)
    print(f'[OK] СОЗДАНО {len(reports)} АНАЛИТИЧЕСКИХ ОТЧЕТОВ')
    print('=' * 80)

if __name__ == '__main__':
    main()
