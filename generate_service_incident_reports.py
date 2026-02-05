#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ГЕНЕРАЦИЯ ДЕТАЛЬНЫХ ОТЧЕТОВ ПО СЛУЖБАМ С ЖАЛОБАМИ
=============================================================================
Создает отчеты в формате Codespaces с разбивкой по регионам и типам жалоб
=============================================================================
"""

import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'reports' / 'службы_детально'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

print('\n' + '='*80)
print('📊 ГЕНЕРАЦИЯ ДЕТАЛЬНЫХ ОТЧЕТОВ ПО СЛУЖБАМ')
print('='*80)

def load_data_from_db():
    """Загрузка данных из БД"""
    print('\n[1/6] Загрузка данных из БД...')
    
    conn = sqlite3.connect(DB_PATH)
    
    # Загружаем фиксации
    query = """
        SELECT 
            f.*,
            o.operator_name,
            s.service_code as Служба_112
        FROM fixations f
        LEFT JOIN operators o ON f.operator_id = o.operator_id  
        LEFT JOIN services s ON f.service_id = s.service_id
        WHERE f.call_date IS NOT NULL
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        print(f'  ✅ Загружено записей: {len(df):,}')
    except:
        # Если таблица fixations не существует, используем fiksa_records
        print('  ⚠️ Таблица fixations не найдена, используем fiksa_records')
        df = pd.read_sql_query("SELECT * FROM fiksa_records", conn)
        # Добавляем колонку службы (нужно извлечь из других источников)
        df['Служба_112'] = None
    
    conn.close()
    
    return df

def categorize_status(status):
    """Категоризация статуса"""
    if pd.isna(status):
        return 'Прочее'
    
    status_str = str(status).lower()
    
    # Положительные
    if any(word in status_str for word in ['положительн', 'qanoatlantir', 'қаноатлантир']):
        return 'Положительно'
    
    # Отрицательные/жалобы
    if any(word in status_str for word in ['отрицательн', 'qanoatlantirilmadi', 'нет ответа', 'не отвечает', 'жалоб']):
        return 'Отрицательно'
    
    # Не дозвонились
    if any(word in status_str for word in ['занято', 'не дозвон', 'нет связи']):
        return 'Не дозвонились'
    
    return 'Прочее'

def create_service_report(df_service, service_code, timestamp):
    """Создание отчета по одной службе"""
    
    service_num = str(service_code).replace('.0', '')
    print(f'\n{"="*80}')
    print(f'СЛУЖБА {service_num}')
    print('='*80)
    
    excel_file = OUTPUT_DIR / f'СЛУЖБА_{service_num}_ИНЦИДЕНТЫ_{timestamp}.xlsx'
    
    # Добавляем категорию
    df_service['Категория'] = df_service['status'].apply(categorize_status)
    
    # Создаем Excel с листами
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        
        # ЛИСТ 1: Жалобы по регионам
        if 'region' in df_service.columns or 'Регион_112' in df_service.columns:
            region_col = 'Регион_112' if 'Регион_112' in df_service.columns else 'region'
            
            region_summary = df_service.groupby(region_col).agg({
                'card_number': 'count',
                'Категория': lambda x: (x == 'Отрицательно').sum()
            }).reset_index()
            
            region_summary.columns = [region_col, 'Количество_жалоб', 'Отрицательные']
            region_summary['Положительные'] = df_service.groupby(region_col).apply(
                lambda x: (x['Категория'] == 'Положительно').sum()
            ).values
            region_summary['Не дозвонились'] = df_service.groupby(region_col).apply(
                lambda x: (x['Категория'] == 'Не дозвонились').sum()
            ).values
            region_summary['Прочее'] = df_service.groupby(region_col).apply(
                lambda x: (x['Категория'] == 'Прочее').sum()
            ).values
            
            region_summary.insert(0, '№', range(1, len(region_summary) + 1))
            region_summary.to_excel(writer, sheet_name='Жалобы_по_регионам', index=False)
            print(f'  ✓ Лист "Жалобы_по_регионам": {len(region_summary)} регионов')
        
        # ЛИСТ 2: Регионы и жалобы (матрица по типам)
        # Упрощенная версия - можно расширить с конкретными типами жалоб
        if 'region' in df_service.columns or 'Регион_112' in df_service.columns:
            region_matrix = region_summary.copy()
            region_matrix.to_excel(writer, sheet_name='Регионы_и_жалобы', index=False)
            print(f'  ✓ Лист "Регионы_и_жалобы"')
        
        # ЛИСТ 3: Детальные (все записи)
        df_details = df_service.copy()
        df_details.insert(0, '№', range(1, len(df_details) + 1))
        
        # Ограничиваем до 100k строк для Excel
        if len(df_details) > 100000:
            df_details = df_details.head(100000)
            print(f'  ⚠️ Детальные ограничены до 100,000 строк из {len(df_service):,}')
        
        df_details.to_excel(writer, sheet_name='Детальные', index=False)
        print(f'  ✓ Лист "Детальные": {len(df_details):,} записей')
        
        # ЛИСТ 4: Отрицательные и жалобы
        df_negative = df_service[df_service['Категория'] == 'Отрицательно'].copy()
        if len(df_negative) > 0:
            df_negative.insert(0, '№', range(1, len(df_negative) + 1))
            df_negative.to_excel(writer, sheet_name='Отрицательные_и_жалобы', index=False)
            print(f'  ✓ Лист "Отрицательные_и_жалобы": {len(df_negative):,} записей')
        
        # ЛИСТ 5: Не найденные заявки (статусы "Прочее")
        df_not_found = df_service[df_service['Категория'] == 'Прочее'].copy()
        if len(df_not_found) > 0:
            df_not_found.insert(0, '№', range(1, len(df_not_found) + 1))
            if len(df_not_found) > 50000:
                df_not_found = df_not_found.head(50000)
            df_not_found.to_excel(writer, sheet_name='Не_найденные_заявки', index=False)
            print(f'  ✓ Лист "Не_найденные_заявки": {len(df_not_found):,} записей')
    
    print(f'  ✅ Создан: {excel_file.name}')
    return excel_file

def create_general_report(df, timestamp):
    """Создание общего отчета по всем службам"""
    print(f'\n{"="*80}')
    print('ОБЩИЙ ОТЧЕТ ПО ВСЕМ СЛУЖБАМ')
    print('='*80)
    
    excel_file = OUTPUT_DIR / f'ОБЩИЙ_ОТЧЕТ_ВСЕ_СЛУЖБЫ_{timestamp}.xlsx'
    
    # Добавляем категорию
    df['Категория'] = df['status'].apply(categorize_status)
    
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        
        # ЛИСТ 1: Сводка количество
        summary_data = {
            'Показатель': [
                'Всего записей',
                'Положительных',
                'Отрицательных',  
                'Не дозвонились',
                'Прочее'
            ],
            'Количество': [
                len(df),
                (df['Категория'] == 'Положительно').sum(),
                (df['Категория'] == 'Отрицательно').sum(),
                (df['Категория'] == 'Не дозвонились').sum(),
                (df['Категория'] == 'Прочее').sum()
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Сводка_количество', index=False)
        print(f'  ✓ Лист "Сводка_количество"')
        
        # ЛИСТ 2: Матрица жалоб
        if 'region' in df.columns or 'Регион_112' in df.columns:
            region_col = 'Регион_112' if 'Регион_112' in df.columns else 'region'
            
            matrix = df.groupby([region_col, 'Категория']).size().unstack(fill_value=0)
            matrix.to_excel(writer, sheet_name='Матрица_жалоб')
            print(f'  ✓ Лист "Матрица_жалоб"')
        
        # ЛИСТ 3: Детальные (ограничено)
        df_details = df.copy()
        df_details.insert(0, '№', range(1, len(df_details) + 1))
        if len(df_details) > 100000:
            df_details = df_details.head(100000)
        df_details.to_excel(writer, sheet_name='Детальные', index=False)
        print(f'  ✓ Лист "Детальные": {len(df_details):,} записей')
        
        # ЛИСТ 4: Статусы по регионам
        if 'region' in df.columns or 'Регион_112' in df.columns:
            region_col = 'Регион_112' if 'Регион_112' in df.columns else 'region'
            
            status_by_region = df.groupby([region_col, 'Категория']).size().unstack(fill_value=0)
            status_by_region.to_excel(writer, sheet_name='Статусы_по_регионам')
            print(f'  ✓ Лист "Статусы_по_регионам"')
        
        # ЛИСТ 5: Отрицательные и жалобы
        df_negative = df[df['Категория'] == 'Отрицательно'].copy()
        if len(df_negative) > 0:
            df_negative.insert(0, '№', range(1, len(df_negative) + 1))
            if len(df_negative) > 100000:
                df_negative = df_negative.head(100000)
            df_negative.to_excel(writer, sheet_name='Отрицательные_и_жалобы', index=False)
            print(f'  ✓ Лист "Отрицательные_и_жалобы": {len(df_negative):,} записей')
        
        # ЛИСТ 6: Жалобы по регионам
        if 'region' in df.columns or 'Регион_112' in df.columns:
            region_summary = df.groupby(region_col).agg({
                'card_number': 'count',
                'Категория': lambda x: (x == 'Отрицательно').sum()
            }).reset_index()
            region_summary.columns = [region_col, 'Количество_жалоб', 'Отрицательные']
            region_summary.insert(0, '№', range(1, len(region_summary) + 1))
            region_summary.to_excel(writer, sheet_name='Жалобы_по_регионам', index=False)
            print(f'  ✓ Лист "Жалобы_по_регионам"')
        
        # ЛИСТ 7: Не найденные заявки
        df_not_found = df[df['Категория'] == 'Прочее'].copy()
        if len(df_not_found) > 0:
            df_not_found.insert(0, '№', range(1, len(df_not_found) + 1))
            if len(df_not_found) > 100000:
                df_not_found = df_not_found.head(100000)
            df_not_found.to_excel(writer, sheet_name='Не_найденные_заявки', index=False)
            print(f'  ✓ Лист "Не_найденные_заявки": {len(df_not_found):,} записей')
    
    print(f'  ✅ Создан: {excel_file.name}')
    return excel_file

def main():
    """Главная функция"""
    
    # Загрузка данных
    df = load_data_from_db()
    
    if df.empty:
        print('\n❌ Нет данных для обработки!')
        return
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    
    # Проверяем наличие колонки службы
    if 'Служба_112' not in df.columns or df['Служба_112'].isna().all():
        print('\n⚠️ ВНИМАНИЕ: Колонка "Служба_112" не найдена или пуста')
        print('Создается только общий отчет без разбивки по службам')
        
        # Создаем только общий отчет
        create_general_report(df, timestamp)
    else:
        # Создаем отчеты по каждой службе
        print('\n[2/6] Создание отчетов по службам...')
        
        services = df['Служба_112'].dropna().unique()
        print(f'  Найдено служб: {sorted(services)}')
        
        for service in sorted(services):
            df_service = df[df['Служба_112'] == service].copy()
            create_service_report(df_service, service, timestamp)
        
        # Создаем общий отчет
        print('\n[3/6] Создание общего отчета...')
        create_general_report(df, timestamp)
    
    print('\n' + '='*80)
    print('✅ ГЕНЕРАЦИЯ ОТЧЕТОВ ЗАВЕРШЕНА УСПЕШНО!')
    print('='*80)
    print(f'📁 Папка: {OUTPUT_DIR}')
    print('='*80 + '\n')

if __name__ == '__main__':
    main()
