import pandas as pd
from pathlib import Path

incident_to_find = '01.AAD4284/26'
sheets_file = sorted(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))[-1]

print(f"Проверка фильтра дат для '{incident_to_find}'\n")

for chunk in pd.read_csv(sheets_file, usecols=['Колонка_2', 'Колонка_4'], chunksize=200000, low_memory=False):
    chunk = chunk.rename(columns={'Колонка_2': 'Инцидент', 'Колонка_4': 'Дата'})
    
    chunk_norm = chunk.copy()
    chunk_norm['Инцидент'] = chunk_norm['Инцидент'].astype(str).str.strip()
    
    match_before = chunk_norm[chunk_norm['Инцидент'] == incident_to_find]
    if not match_before.empty:
        print("До фильтрации:")
        print(f"  Найдено записей: {len(match_before)}")
        print(f"  Дата: {match_before.iloc[0]['Дата']}")
        
        # Применяем фильтр ТОЧНО ТАК ЖЕ как в коде
        date_series = pd.to_datetime(chunk['Дата'], errors='coerce', dayfirst=True)
        start_date = pd.Timestamp('2026-01-04')
        end_date = pd.Timestamp('2026-01-31 23:59:59')
        
        print(f"\nПарсинг даты для записи:")
        parsed_date = pd.to_datetime(match_before.iloc[0]['Дата'], errors='coerce', dayfirst=True)
        print(f"  Исходная: {match_before.iloc[0]['Дата']}")
        print(f"  Парсированная: {parsed_date}")
        print(f"  Начало диапазона: {start_date}")
        print(f"  Конец диапазона: {end_date}")
        print(f"  >= start: {parsed_date >= start_date}")
        print(f"  <= end: {parsed_date <= end_date}")
        
        # Фильтруем
        chunk_filtered = chunk[(date_series >= start_date) & (date_series <= end_date)].copy()
        
        print(f"\nПосле фильтрации:")
        print(f"  Всего записей в chunk: {len(chunk)} -> {len(chunk_filtered)}")
        
        chunk_filtered['Инцидент'] = chunk_filtered['Инцидент'].astype(str).str.strip()
        match_after = chunk_filtered[chunk_filtered['Инцидент'] == incident_to_find]
        print(f"  Найдено '{incident_to_find}': {len(match_after)}")
        
        if match_after.empty:
            print("\n❌ ИНЦИДЕНТ ПОТЕРЯН ПРИ ФИЛЬТРАЦИИ!")
            print("\nПроверяем индексы:")
            print(f"  Индекс до фильтра: {match_before.index.tolist()}")
            print(f"  Индексы после фильтра (первые 10): {chunk_filtered.index.tolist()[:10]}")
            
            # Проверяем, есть ли индекс записи в отфильтрованном chunk
            original_idx = match_before.index[0]
            if original_idx in chunk_filtered.index:
                print(f"  ✓ Индекс {original_idx} присутствует в отфильтрованном chunk")
                print(f"  Значение инцидента: {chunk_filtered.loc[original_idx, 'Инцидент']}")
            else:
                print(f"  ❌ Индекс {original_idx} ОТСУТСТВУЕТ в отфильтрованном chunk!")
                print(f"  Это значит, что строка была отброшена фильтром дат")
        
        break
