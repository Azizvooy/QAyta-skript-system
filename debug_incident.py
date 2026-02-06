import pandas as pd
import re
from pathlib import Path
from process_period_data import rename_statuses

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

incident_to_find = '01.AAD4284/26'
sheets_file = sorted(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))[-1]

print(f"Ищем инцидент: '{incident_to_find}'")
print(f"Файл: {sheets_file.name}\n")

for chunk in pd.read_csv(sheets_file, usecols=['Колонка_2', 'Колонка_4', 'Колонка_5', 'Колонка_6', 'Колонка_7', 'Колонка_8'], chunksize=200000, low_memory=False):
    chunk = chunk.rename(columns={
        'Колонка_2': 'Инцидент_Sheets',
        'Колонка_4': 'Дата_открытия',
        'Колонка_5': 'Статус_связи',
        'Колонка_6': 'Служба_Sheets',
        'Колонка_7': 'Жалоба',
        'Колонка_8': 'Положительно'
    })
    
    # Нормализация ДО фильтра
    chunk['Инцидент_Sheets_norm'] = chunk['Инцидент_Sheets'].astype(str).str.strip()
    
    # Проверяем до фильтрации
    match_before = chunk[chunk['Инцидент_Sheets_norm'] == incident_to_find]
    if not match_before.empty:
        print("НАЙДЕНО ДО ФИЛЬТРА ПО ДАТАМ:")
        row = match_before.iloc[0]
        print(f"  Инцидент: '{row['Инцидент_Sheets_norm']}'")
        print(f"  Дата: {row['Дата_открытия']}")
        
        # Применяем фильтр
        date_series = pd.to_datetime(chunk['Дата_открытия'], errors='coerce', dayfirst=True)
        start_date = pd.Timestamp('2026-01-04')
        end_date = pd.Timestamp('2026-01-31 23:59:59')
        chunk_filtered = chunk[(date_series >= start_date) & (date_series <= end_date)].copy()
        
        match_after = chunk_filtered[chunk_filtered['Инцидент_Sheets_norm'] == incident_to_find]
        print(f"\nПОСЛЕ ФИЛЬТРА ПО ДАТАМ:")
        if not match_after.empty:
            print("  ✓ ПРОШЕЛ ФИЛЬТР")
            
            # Применяем rename_statuses
            chunk_filtered['Статус_связи'] = chunk_filtered['Статус_связи'].apply(rename_statuses)
            
            row_after = chunk_filtered[chunk_filtered['Инцидент_Sheets_norm'] == incident_to_find].iloc[0]
            
            incident = str(row_after['Инцидент_Sheets_norm']).strip()
            complaint = str(row_after.get('Жалоба', '')).strip()
            if complaint in ('nan', 'None'):
                complaint = ''
            status = str(row_after.get('Статус_связи', '')).strip()
            if status in ('nan', 'None'):
                status = ''
            positive = str(row_after.get('Положительно', '')).strip()
            if positive in ('nan', 'None'):
                positive = ''
            
            print(f"\nПОСЛЕ ОБРАБОТКИ:")
            print(f"  incident = '{incident}'")
            print(f"  complaint = '{complaint}'")
            print(f"  status = '{status}'")
            print(f"  positive = '{positive}'")
            
            services = extract_services(row_after.get('Служба_Sheets'))
            print(f"  services = {services}")
            
            print(f"\nПРОВЕРКИ:")
            print(f"  incident in ('', 'nan', 'None'): {incident in ('', 'nan', 'None')}")
            print(f"  status пустой: {not status}")
            print(f"  positive пустой: {not positive}")
            
            if incident in ('', 'nan', 'None'):
                print("\n❌ ПРОПУЩЕН: incident пустой!")
            else:
                if not status:
                    print("\n❌ НЕ ПОПАДЕТ в incident_status (status пустой)")
                else:
                    print(f"\n✓ ПОПАДЕТ в incident_status['{incident}'] = '{status}'")
                
                if not positive:
                    print("❌ НЕ ПОПАДЕТ в incident_positive (positive пустой)")
                else:
                    print(f"✓ ПОПАДЕТ в incident_positive['{incident}'] = '{positive}'")
        else:
            print("  ❌ НЕ ПРОШЕЛ ФИЛЬТР ПО ДАТАМ")
        
        break
