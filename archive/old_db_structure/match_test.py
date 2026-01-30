import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def match_applications_with_fiksa():
    conn = get_db_connection()
    
    print('\n' + '=' * 80)
    print('СОПОСТАВЛЕНИЕ ЗАЯВОК С ДАННЫМИ ФИКСАЦИИ (ПО НОМЕРУ ТЕЛЕФОНА)')
    print('=' * 80)
    
    applications = pd.read_sql_query('''
        SELECT a.id, a.application_number, a.phone, a.address, a.import_date, a.notes
        FROM applications a
        WHERE a.phone IS NOT NULL AND a.phone != ''
        ORDER BY a.import_date DESC
    ''', conn)
    
    if applications.empty:
        print('НЕТ ЗАЯВОК')
        return
    
    print(f'Найдено заявок: {len(applications)}')
    
    results = []
    matched_count = 0
    
    for idx, app in applications.iterrows():
        if idx % 1000 == 0:
            print(f'  Обработано: {idx}/{len(applications)}...')
            
        phone = app['phone']
        phone_clean = phone.replace('+998', '').replace('+', '').replace(' ', '').replace('-', '')
        
        fiksa_data = pd.read_sql_query(f'''
            SELECT operator_name, card_number, full_name, status, call_date, notes
            FROM fiksa_records
            WHERE phone LIKE '%{phone_clean}%' OR phone LIKE '%{phone}%'
            LIMIT 1
        ''', conn)
        
        if not fiksa_data.empty:
            matched_count += 1
        
        results.append({'phone': phone, 'found': not fiksa_data.empty})
    
    print(f'\nНайдено: {matched_count} из {len(results)}')
    print(f'Процент: {matched_count/len(results)*100:.1f}%')
    conn.close()

if __name__ == '__main__':
    match_applications_with_fiksa()
