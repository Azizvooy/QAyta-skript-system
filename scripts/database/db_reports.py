#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
СОЗДАНИЕ ОТЧЕТОВ ИЗ УЛУЧШЕННОЙ БД
=============================================================================
Генерирует все отчеты напрямую из БД без пустых значений
=============================================================================
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import sys

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'
LOG_DIR = BASE_DIR / 'logs' / 'database'

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

def log_report(message, level='INFO'):
    """Логирование создания отчетов"""
    log_file = LOG_DIR / f'reports_log_{datetime.now().strftime("%Y%m%d")}.log'
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'[{timestamp}] [{level}] {message}\n')
    
    print(f'[{level}] {message}')

def get_db_connection():
    """Подключение к БД"""
    return sqlite3.connect(DB_PATH)

def create_operator_report():
    """Отчет по операторам из БД"""
    log_report('Создание отчета по операторам')
    
    conn = get_db_connection()
    
    # Используем представление для статистики
    query = '''
        SELECT 
            operator_name as "Оператор",
            total_fixations as "Всего звонков",
            positive as "Положительных",
            negative as "Отрицательных",
            no_answer as "Нет ответа",
            positive_percent as "% Положительных"
        FROM v_operator_stats
        WHERE total_fixations > 0
        ORDER BY total_fixations DESC
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        log_report('Нет данных для отчета по операторам', 'WARNING')
        return None
    
    # Сохранение
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'ОТЧЕТ_ОПЕРАТОРЫ_{timestamp}.xlsx'
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Статистика операторов', index=False)
            
            # Автонастройка колонок
            worksheet = writer.sheets['Статистика операторов']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        cell_value = str(cell.value) if cell.value is not None else ''
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except:
                        pass
                adjusted_width = min(max(max_length + 2, 10), 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        log_report(f'Отчет по операторам сохранен: {output_file}', 'SUCCESS')
        return output_file
        
    except Exception as e:
        log_report(f'Ошибка при создании отчета: {str(e)}', 'ERROR')
        return None

def create_service_report():
    """Отчет по службам из БД"""
    log_report('Создание отчета по службам')
    
    conn = get_db_connection()
    
    query = '''
        SELECT 
            s.service_code as "Код службы",
            s.service_name as "Название службы",
            COUNT(DISTINCT a.application_id) as "Всего заявок",
            COUNT(DISTINCT f.fixation_id) as "Зафиксировано",
            SUM(CASE WHEN f.status LIKE '%Положительн%' THEN 1 ELSE 0 END) as "Положительных",
            SUM(CASE WHEN f.status LIKE '%Отрицательн%' THEN 1 ELSE 0 END) as "Отрицательных",
            ROUND(
                CAST(SUM(CASE WHEN f.status LIKE '%Положительн%' THEN 1 ELSE 0 END) AS REAL) * 100.0 /
                NULLIF(COUNT(DISTINCT f.fixation_id), 0),
                2
            ) as "% Положительных"
        FROM services s
        LEFT JOIN applications a ON s.service_id = a.service_id
        LEFT JOIN fixations f ON a.application_id = f.application_id
        GROUP BY s.service_id, s.service_code, s.service_name
        HAVING COUNT(DISTINCT a.application_id) > 0
        ORDER BY s.service_code
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        log_report('Нет данных для отчета по службам', 'WARNING')
        return None
    
    # Сохранение
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'ОТЧЕТ_СЛУЖБЫ_{timestamp}.xlsx'
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Статистика служб', index=False)
            
            # Автонастройка колонок
            worksheet = writer.sheets['Статистика служб']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        cell_value = str(cell.value) if cell.value is not None else ''
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except:
                        pass
                adjusted_width = min(max(max_length + 2, 10), 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        log_report(f'Отчет по службам сохранен: {output_file}', 'SUCCESS')
        return output_file
        
    except Exception as e:
        log_report(f'Ошибка при создании отчета: {str(e)}', 'ERROR')
        return None

def create_detailed_report():
    """Детальный отчет со всеми данными"""
    log_report('Создание детального отчета')
    
    conn = get_db_connection()
    
    # Используем представление для полных данных
    query = '''
        SELECT 
            collection_date as "Дата фиксации",
            operator_name as "Оператор",
            application_number as "Номер заявки",
            card_number as "Номер карты",
            service_code as "Служба",
            service_name as "Название службы",
            region_name as "Регион",
            phone_called as "Телефон",
            status as "Статус",
            feedback_type as "Тип фидбэка",
            notes as "Примечания",
            created_at as "Создано"
        FROM v_fixations_full
        ORDER BY collection_date DESC, operator_name
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        log_report('Нет данных для детального отчета', 'WARNING')
        return None
    
    # Сохранение
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'ОТЧЕТ_ДЕТАЛЬНЫЙ_{timestamp}.xlsx'
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Детальные данные', index=False)
            
            # Автонастройка колонок
            worksheet = writer.sheets['Детальные данные']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        cell_value = str(cell.value) if cell.value is not None else ''
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except:
                        pass
                adjusted_width = min(max(max_length + 2, 10), 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        log_report(f'Детальный отчет сохранен: {output_file}', 'SUCCESS')
        return output_file
        
    except Exception as e:
        log_report(f'Ошибка при создании отчета: {str(e)}', 'ERROR')
        return None

def create_summary_report():
    """Сводный отчет с общей статистикой"""
    log_report('Создание сводного отчета')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Общая статистика
    cursor.execute('''
        SELECT 
            COUNT(DISTINCT operator_id) as operators_count,
            COUNT(DISTINCT application_id) as applications_count,
            COUNT(*) as fixations_count
        FROM fixations
    ''')
    general = cursor.fetchone()
    
    # Статистика по статусам
    cursor.execute('''
        SELECT 
            status,
            COUNT(*) as count,
            ROUND(CAST(COUNT(*) AS REAL) * 100.0 / 
                  (SELECT COUNT(*) FROM fixations), 2) as percent
        FROM fixations
        GROUP BY status
        ORDER BY count DESC
    ''')
    statuses = cursor.fetchall()
    
    # Статистика по службам
    cursor.execute('''
        SELECT 
            s.service_name,
            COUNT(f.fixation_id) as count
        FROM services s
        LEFT JOIN applications a ON s.service_id = a.service_id
        LEFT JOIN fixations f ON a.application_id = f.application_id
        WHERE f.fixation_id IS NOT NULL
        GROUP BY s.service_id, s.service_name
        ORDER BY count DESC
    ''')
    services = cursor.fetchall()
    
    conn.close()
    
    # Формируем отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = OUTPUT_DIR / f'ОТЧЕТ_СВОДНЫЙ_{timestamp}.xlsx'
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Лист 1: Общая статистика
            df_general = pd.DataFrame([{
                'Показатель': 'Количество операторов',
                'Значение': general[0]
            }, {
                'Показатель': 'Количество заявок',
                'Значение': general[1]
            }, {
                'Показатель': 'Всего фиксаций',
                'Значение': general[2]
            }])
            df_general.to_excel(writer, sheet_name='Общая статистика', index=False)
            
            # Лист 2: По статусам
            df_statuses = pd.DataFrame(statuses, columns=['Статус', 'Количество', 'Процент'])
            df_statuses.to_excel(writer, sheet_name='По статусам', index=False)
            
            # Лист 3: По службам
            df_services = pd.DataFrame(services, columns=['Служба', 'Количество'])
            df_services.to_excel(writer, sheet_name='По службам', index=False)
            
            # Автонастройка всех листов
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
                        except:
                            pass
                    adjusted_width = min(max(max_length + 2, 10), 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        log_report(f'Сводный отчет сохранен: {output_file}', 'SUCCESS')
        return output_file
        
    except Exception as e:
        log_report(f'Ошибка при создании отчета: {str(e)}', 'ERROR')
        return None

def create_all_reports():
    """Создать все отчеты"""
    print('=' * 80)
    print('СОЗДАНИЕ ОТЧЕТОВ ИЗ БД')
    print('=' * 80)
    
    reports = []
    
    # 1. Отчет по операторам
    print('\n1️⃣  Отчет по операторам...')
    report = create_operator_report()
    if report:
        reports.append(report)
        print(f'   ✅ {report.name}')
    
    # 2. Отчет по службам
    print('\n2️⃣  Отчет по службам...')
    report = create_service_report()
    if report:
        reports.append(report)
        print(f'   ✅ {report.name}')
    
    # 3. Детальный отчет
    print('\n3️⃣  Детальный отчет...')
    report = create_detailed_report()
    if report:
        reports.append(report)
        print(f'   ✅ {report.name}')
    
    # 4. Сводный отчет
    print('\n4️⃣  Сводный отчет...')
    report = create_summary_report()
    if report:
        reports.append(report)
        print(f'   ✅ {report.name}')
    
    print('\n' + '=' * 80)
    print(f'СОЗДАНО ОТЧЕТОВ: {len(reports)}')
    print('=' * 80)
    
    return reports

if __name__ == '__main__':
    create_all_reports()
