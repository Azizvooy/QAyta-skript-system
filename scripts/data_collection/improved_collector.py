#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
УЛУЧШЕННЫЙ СБОРЩИК ДАННЫХ FIKSA С ФИЛЬТРАЦИЕЙ ПО СТАТУСУ
=============================================================================
- Собирает только записи где колонка E (статус) заполнена
- Автоматический импорт из Google Sheets
- Расчет статистики операторов
- Анализ фидбэков от служб
=============================================================================
"""

import os
import sqlite3
from datetime import datetime
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import socket
import pandas as pd

# Прокси
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
TOKEN_FILE = BASE_DIR / 'config' / 'token.json'
CREDENTIALS_FILE = BASE_DIR / 'config' / 'credentials.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

# =============================================================================
# GOOGLE API
# =============================================================================

def authenticate():
    """Аутентификация в Google API"""
    creds = None
    
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return creds

def get_operator_sheets():
    """Получить список всех листов операторов"""
    creds = authenticate()
    service = build('sheets', 'v4', credentials=creds)
    
    spreadsheet = service.spreadsheets().get(spreadsheetId=MASTER_SPREADSHEET_ID).execute()
    sheets = spreadsheet.get('sheets', [])
    
    operator_sheets = []
    for sheet in sheets:
        title = sheet['properties']['title']
        if title not in ['Настройки', 'Статистика', 'Сводка']:
            operator_sheets.append(title)
    
    return operator_sheets, service

def collect_fiksa_data(operator_name, service):
    """Собрать данные с листа оператора (только со статусом в колонке E)"""
    try:
        # Читаем диапазон A2:Z10000
        range_name = f"'{operator_name}'!A2:Z10000"
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return []
        
        records = []
        today = datetime.now().strftime('%Y-%m-%d')
        
        for row in values:
            # Проверяем что строка не пустая
            if not row or len(row) < 5:
                continue
            
            # КРИТИЧНО: Проверяем колонку E (индекс 4) - статус должен быть заполнен
            status = row[4] if len(row) > 4 else ''
            if not status or status.strip() == '':
                continue  # Пропускаем строки без статуса
            
            # Проверяем что это именно FIKSA запись
            card_number = row[0] if len(row) > 0 else ''
            if 'FIKSA' not in operator_name and not card_number:
                continue
            
            record = {
                'collection_date': today,
                'operator_name': operator_name,
                'card_number': row[0] if len(row) > 0 else None,
                'full_name': row[1] if len(row) > 1 else None,
                'phone': row[2] if len(row) > 2 else None,
                'address': row[3] if len(row) > 3 else None,
                'status': status,
                'call_date': row[5] if len(row) > 5 else None,
                'notes': row[6] if len(row) > 6 else None,
            }
            
            records.append(record)
        
        return records
        
    except Exception as e:
        print(f'   [ОШИБКА] {operator_name}: {e}')
        return []

# =============================================================================
# БАЗА ДАННЫХ
# =============================================================================

def save_to_database(records):
    """Сохранить записи в БД"""
    if not records:
        return 0
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Создаем таблицу если не существует
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fiksa_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            collection_date DATE NOT NULL,
            operator_name TEXT NOT NULL,
            card_number TEXT,
            full_name TEXT,
            phone TEXT,
            address TEXT,
            status TEXT,
            call_date DATETIME,
            notes TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Очищаем старые записи этого оператора за сегодня
    today = datetime.now().strftime('%Y-%m-%d')
    cursor.execute('''
        DELETE FROM fiksa_records 
        WHERE collection_date = ? AND operator_name = ?
    ''', (today, records[0]['operator_name']))
    
    # Вставляем новые записи
    for record in records:
        cursor.execute('''
            INSERT INTO fiksa_records (
                collection_date, operator_name, card_number, full_name,
                phone, address, status, call_date, notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            record['collection_date'],
            record['operator_name'],
            record['card_number'],
            record['full_name'],
            record['phone'],
            record['address'],
            record['status'],
            record['call_date'],
            record['notes']
        ))
    
    conn.commit()
    conn.close()
    
    return len(records)

# =============================================================================
# СТАТИСТИКА ОПЕРАТОРОВ
# =============================================================================

def calculate_operator_stats():
    """Расчет статистики по операторам"""
    conn = sqlite3.connect(DB_PATH)
    
    query = '''
        SELECT 
            operator_name,
            COUNT(*) as total,
            COUNT(CASE WHEN status LIKE '%Положительн%' THEN 1 END) as positive,
            COUNT(CASE WHEN status LIKE '%Отрицательн%' THEN 1 END) as negative,
            COUNT(CASE WHEN status LIKE '%Недозвон%' THEN 1 END) as no_answer,
            COUNT(CASE WHEN status LIKE '%Нет ответа%' OR status LIKE '%занято%' THEN 1 END) as busy
        FROM fiksa_records
        WHERE collection_date = date('now')
        GROUP BY operator_name
        ORDER BY total DESC
    '''
    
    df = pd.read_sql_query(query, conn)
    
    # Сохраняем статистику
    today = datetime.now().strftime('%Y-%m-%d')
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS operator_stats_daily (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            stat_date DATE NOT NULL,
            operator_name TEXT NOT NULL,
            total INTEGER,
            positive INTEGER,
            negative INTEGER,
            no_answer INTEGER,
            busy INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Удаляем старую статистику за сегодня
    cursor.execute('DELETE FROM operator_stats_daily WHERE stat_date = ?', (today,))
    
    # Вставляем новую
    for _, row in df.iterrows():
        cursor.execute('''
            INSERT INTO operator_stats_daily (
                stat_date, operator_name, total, positive, negative, no_answer, busy
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            today,
            row['operator_name'],
            row['total'],
            row['positive'],
            row['negative'],
            row['no_answer'],
            row['busy']
        ))
    
    conn.commit()
    conn.close()
    
    return df

# =============================================================================
# АНАЛИЗ ФИДБЭКОВ ОТ СЛУЖБ
# =============================================================================

def analyze_service_feedback():
    """Анализ фидбэков от служб по заявкам"""
    conn = sqlite3.connect(DB_PATH)
    
    query = '''
        SELECT 
            ch.service_code,
            ch.service_name,
            f.status as feedback_status,
            COUNT(*) as count,
            COUNT(CASE WHEN f.status LIKE '%Положительн%' THEN 1 END) as positive_feedback,
            COUNT(CASE WHEN f.status LIKE '%Отрицательн%' THEN 1 END) as negative_feedback,
            ROUND(COUNT(CASE WHEN f.status LIKE '%Положительн%' THEN 1 END) * 100.0 / COUNT(*), 1) as positive_rate
        FROM call_history_112 ch
        LEFT JOIN fiksa_records f ON f.full_name = ch.incident_number
        WHERE f.status IS NOT NULL AND f.status != ''
        GROUP BY ch.service_code, ch.service_name, f.status
        ORDER BY ch.service_code, count DESC
    '''
    
    df = pd.read_sql_query(query, conn)
    
    # Сохраняем анализ фидбэков
    today = datetime.now().strftime('%Y-%m-%d')
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS service_feedback_stats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            stat_date DATE NOT NULL,
            service_code TEXT,
            service_name TEXT,
            feedback_status TEXT,
            count INTEGER,
            positive_count INTEGER,
            negative_count INTEGER,
            positive_rate REAL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('DELETE FROM service_feedback_stats WHERE stat_date = ?', (today,))
    
    for _, row in df.iterrows():
        cursor.execute('''
            INSERT INTO service_feedback_stats (
                stat_date, service_code, service_name, feedback_status,
                count, positive_count, negative_count, positive_rate
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            today,
            row['service_code'],
            row['service_name'],
            row['feedback_status'],
            row['count'],
            row['positive_feedback'],
            row['negative_feedback'],
            row['positive_rate']
        ))
    
    conn.commit()
    conn.close()
    
    return df

# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================

def main():
    """Главная функция сбора данных"""
    print('\n' + '=' * 80)
    print('АВТОМАТИЧЕСКИЙ СБОР ДАННЫХ FIKSA (ТОЛЬКО СО СТАТУСОМ)')
    print('=' * 80)
    print(f'Время запуска: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    
    # 1. Получаем список операторов
    print('\n[1/4] Получение списка операторов...')
    operators, service = get_operator_sheets()
    print(f'   Найдено операторов: {len(operators)}')
    
    # 2. Собираем данные
    print('\n[2/4] Сбор данных из Google Sheets...')
    total_collected = 0
    
    for operator in operators:
        print(f'   Обработка: {operator}...')
        records = collect_fiksa_data(operator, service)
        
        if records:
            saved = save_to_database(records)
            total_collected += saved
            print(f'      [OK] Сохранено: {saved} записей со статусом')
        else:
            print(f'      [ПУСТО] Нет записей со статусом')
    
    print(f'\n   ИТОГО собрано: {total_collected} записей')
    
    # 3. Расчет статистики операторов
    print('\n[3/4] Расчет статистики операторов...')
    stats = calculate_operator_stats()
    print(f'   Статистика по {len(stats)} операторам')
    
    # 4. Анализ фидбэков от служб
    print('\n[4/4] Анализ фидбэков от служб...')
    feedback = analyze_service_feedback()
    print(f'   Проанализировано {len(feedback)} категорий фидбэков')
    
    print('\n' + '=' * 80)
    print('[OK] СБОР ДАННЫХ ЗАВЕРШЕН')
    print('=' * 80)
    
    return total_collected

if __name__ == '__main__':
    main()
