import pandas as pd
import re
from pathlib import Path

def extract_services(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    text = str(value)
    found = re.findall(r'\b(101|102|103|104)\b', text)
    if found:
        return list(dict.fromkeys(found))
    parts = re.split(r'[;,/\\|\\s]+', text)
    parts = [p.strip() for p in parts if p.strip()]
    return list(dict.fromkeys(parts))

incident = '01.AAD4284/26'
sheets_file = sorted(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))[-1]

for chunk in pd.read_csv(sheets_file, usecols=['Колонка_2', 'Колонка_4', 'Колонка_5', 'Колонка_6', 'Колонка_7', 'Колонка_8'], chunksize=200000, low_memory=False):
    chunk = chunk.rename(columns={
        'Колонка_2': 'Инцидент_Sheets',
        'Колонка_4': 'Дата_открытия',
        'Колонка_5': 'Статус_связи',
        'Колонка_6': 'Служба_Sheets',
        'Колонка_7': 'Жалоба',
        'Колонка_8': 'Положительно'
    })
    
    # Фильтр по датам
    date_series = pd.to_datetime(chunk['Дата_открытия'], errors='coerce', dayfirst=True)
    start_date = pd.Timestamp('2026-01-04')
    end_date = pd.Timestamp('2026-01-31 23:59:59')
    chunk = chunk[(date_series >= start_date) & (date_series <= end_date)].copy()
    
    chunk['Инцидент_Sheets_norm'] = chunk['Инцидент_Sheets'].astype(str).str.strip()
    
    match = chunk[chunk['Инцидент_Sheets_norm'] == incident]
    if not match.empty:
        row = match.iloc[0]
        print('Найдено в Sheets после фильтрации:')
        print(f'  Статус: {row["Статус_связи"]}')
        print(f'  Служба: {row["Служба_Sheets"]}')
        print(f'  Жалоба: {row["Жалоба"]}')
        print(f'  Положительно: {row["Положительно"]}')
        
        services = extract_services(row['Служба_Sheets'])
        
        complaint = str(row.get('Жалоба', '')).strip()
        if complaint in ('nan', 'None', 'NaN'):
            complaint = ''
        status = str(row.get('Статус_связи', '')).strip()
        if status in ('nan', 'None', 'NaN'):
            status = ''
        positive = str(row.get('Положительно', '')).strip()
        if positive in ('nan', 'None', 'NaN'):
            positive = ''
            
        print(f'\nОбработано:')
        print(f'  complaint="{complaint}"')
        print(f'  status="{status}"')
        print(f'  positive="{positive}"')
        print(f'  services={services}')
        
        if not services:
            print(f'\nСлужбы НЕ указаны -> должно попасть в incident_all_complaints')
            print(f'НО! В коде load_sheets_maps() есть проверка:')
            print(f'    if complaint: incident_all_complaints[incident].add(complaint)')
            print(f'Так как complaint пустой, он НЕ попадет в incident_all_complaints!')
            print(f'\nНО статус и положительно должны попасть в incident_status/incident_positive')
        break
