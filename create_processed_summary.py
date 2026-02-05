#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
СОЗДАНИЕ СВОДНОГО ОТЧЕТА ПО ОБРАБОТАННЫМ ДАННЫМ
=============================================================================
Создает отчет по данным из БД с детальной статистикой
=============================================================================
"""

import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'processed_data' / datetime.now().strftime('%Y-%m-%d') / 'reports'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print('\n' + '='*80)
print('📊 СОЗДАНИЕ СВОДНОГО ОТЧЕТА ПО ОБРАБОТАННЫМ ДАННЫМ')
print('='*80)

def create_summary_report():
    """Создание сводного отчета"""
    conn = sqlite3.connect(DB_PATH)
    
    # Общая статистика
    total_query = "SELECT COUNT(*) as count FROM fiksa_records"
    total = pd.read_sql_query(total_query, conn).iloc[0]['count']
    
    print(f'\n📌 Всего записей в БД: {total:,}')
    
    # Статистика по статусам
    print('\n[1/5] Анализ по статусам...')
    status_query = """
        SELECT 
            status,
            COUNT(*) as count,
            ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM fiksa_records), 2) as percentage
        FROM fiksa_records
        WHERE status IS NOT NULL AND status != ''
        GROUP BY status
        ORDER BY count DESC
    """
    df_status = pd.read_sql_query(status_query, conn)
    
    # Категоризация
    df_status['category'] = 'Другое'
    df_status.loc[df_status['status'].str.contains('Положительн|qanoatlantir', case=False, na=False), 'category'] = 'Положительно'
    df_status.loc[df_status['status'].str.contains('Отрицательн|qanoatlantirilmadi|НЕТ ОТВЕТА|не отвечает', case=False, na=False), 'category'] = 'Отрицательно'
    
    # Сводка по категориям
    category_stats = df_status.groupby('category')['count'].sum().reset_index()
    category_stats['percentage'] = (category_stats['count'] / total * 100).round(2)
    
    print(f'  ✅ Найдено статусов: {len(df_status)}')
    print('\\n  📊 По категориям:')
    for _, row in category_stats.iterrows():
        print(f'     {row["category"]}: {row["count"]:,} ({row["percentage"]}%)')
    
    # Статистика по операторам  
    print('\\n[2/5] Анализ по операторам...')
    operator_query = """
        SELECT 
            operator_name,
            COUNT(*) as total_fixations,
            SUM(CASE WHEN status LIKE '%Положительн%' OR status LIKE '%qanoatlantir%' THEN 1 ELSE 0 END) as positive,
            SUM(CASE WHEN status LIKE '%Отрицательн%' OR status LIKE '%НЕТ ОТВЕТА%' THEN 1 ELSE 0 END) as negative
        FROM fiksa_records
        WHERE operator_name IS NOT NULL AND operator_name != ''
        GROUP BY operator_name
        ORDER BY total_fixations DESC
    """
    df_operators = pd.read_sql_query(operator_query, conn)
    df_operators['positive_pct'] = (df_operators['positive'] / df_operators['total_fixations'] * 100).round(2)
    df_operators['negative_pct'] = (df_operators['negative'] / df_operators['total_fixations'] * 100).round(2)
    
    print(f'  ✅ Найдено операторов: {len(df_operators)}')
    
    # Статистика по датам
    print('\\n[3/5] Анализ по датам...')
    date_query = """
        SELECT 
            substr(call_date, 1, 10) as date,
            COUNT(*) as count
        FROM fiksa_records
        WHERE call_date IS NOT NULL AND call_date != ''
        GROUP BY substr(call_date, 1, 10)
        ORDER BY date DESC
        LIMIT 30
    """
    df_dates = pd.read_sql_query(date_query, conn)
    
    print(f'  ✅ Найдено дат: {len(df_dates)}')
    
    # Создание Excel отчета
    print('\\n[4/5] Создание Excel отчета...')
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    excel_file = OUTPUT_DIR / f'СВОДНЫЙ_ОТЧЕТ_{timestamp}.xlsx'
    
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        # Лист 1: Общая сводка
        summary_data = {
            'Показатель': [
                'Всего записей',
                'Положительных',
                'Отрицательных',
                'Других',
                'Операторов',
                'Уникальных дат'
            ],
            'Значение': [
                total,
                category_stats[category_stats['category'] == 'Положительно']['count'].sum() if 'Положительно' in category_stats['category'].values else 0,
                category_stats[category_stats['category'] == 'Отрицательно']['count'].sum() if 'Отрицательно' in category_stats['category'].values else 0,
                category_stats[category_stats['category'] == 'Другое']['count'].sum() if 'Другое' in category_stats['category'].values else 0,
                len(df_operators),
                len(df_dates)
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='СВОДКА', index=False)
        
        # Лист 2: По статусам
        df_status.to_excel(writer, sheet_name='ПО_СТАТУСАМ', index=False)
        
        # Лист 3: По категориям
        category_stats.to_excel(writer, sheet_name='ПО_КАТЕГОРИЯМ', index=False)
        
        # Лист 4: По операторам
        df_operators.to_excel(writer, sheet_name='ПО_ОПЕРАТОРАМ', index=False)
        
        # Лист 5: По датам (последние 30 дней)
        df_dates.to_excel(writer, sheet_name='ПО_ДАТАМ', index=False)
    
    print(f'  ✅ Создан: {excel_file.name}')
    
    # Создание текстового отчета
    print('\\n[5/5] Создание текстового отчета...')
    txt_file = OUTPUT_DIR / f'СВОДНЫЙ_ОТЧЕТ_{timestamp}.txt'
    
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write('='*80 + '\\n')
        f.write('СВОДНЫЙ ОТЧЕТ ПО ОБРАБОТАННЫМ ДАННЫМ\\n')
        f.write('='*80 + '\\n')
        f.write(f'Дата: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\\n')
        f.write(f'База данных: {DB_PATH}\\n')
        f.write('='*80 + '\\n\\n')
        
        f.write('📊 ОБЩАЯ СТАТИСТИКА\\n')
        f.write('-'*80 + '\\n')
        f.write(f'Всего записей: {total:,}\\n\\n')
        
        f.write('📋 ПО КАТЕГОРИЯМ\\n')
        f.write('-'*80 + '\\n')
        for _, row in category_stats.iterrows():
            f.write(f'{row["category"]:<30} - {row["count"]:>10,} ({row["percentage"]:>5.1f}%)\\n')
        f.write('\\n')
        
        f.write('👥 ТОП-20 ОПЕРАТОРОВ\\n')
        f.write('-'*80 + '\\n')
        for idx, row in df_operators.head(20).iterrows():
            f.write(f'{idx+1:2}. {row["operator_name"]:<40} - {row["total_fixations"]:>7,} фиксаций\\n')
            f.write(f'     Положительных: {row["positive"]:>7,} ({row["positive_pct"]:>5.1f}%)  ')
            f.write(f'Отрицательных: {row["negative"]:>7,} ({row["negative_pct"]:>5.1f}%)\\n')
        f.write('\\n')
        
        f.write('📅 ПОСЛЕДНИЕ 10 ДНЕЙ\\n')
        f.write('-'*80 + '\\n')
        for idx, row in df_dates.head(10).iterrows():
            f.write(f'{row["date"]}: {row["count"]:,} фиксаций\\n')
        f.write('\\n')
        
        f.write('='*80 + '\\n')
        f.write('✅ ОТЧЕТ ЗАВЕРШЕН\\n')
        f.write('='*80 + '\\n')
    
    print(f'  ✅ Создан: {txt_file.name}')
    
    conn.close()
    
    return excel_file, txt_file

def main():
    """Главная функция"""
    excel_file, txt_file = create_summary_report()
    
    print('\\n' + '='*80)
    print('✅ СВОДНЫЙ ОТЧЕТ СОЗДАН УСПЕШНО!')
    print('='*80)
    print(f'📄 Excel: {excel_file}')
    print(f'📄 Текст: {txt_file}')
    print('='*80 + '\\n')

if __name__ == '__main__':
    main()
