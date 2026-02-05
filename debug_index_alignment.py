import pandas as pd

# Читаем Sheets данные
sheets_file = 'data/КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_20260202_130742.csv'

print("Проверка выравнивания индексов при фильтрации")
print("=" * 60)

for chunk in pd.read_csv(sheets_file, chunksize=200000, encoding='utf-8-sig'):
    # Переименование
    chunk = chunk.rename(columns={
        'Колонка_2': 'Инцидент_Sheets',
        'Колонка_4': 'Дата_открытия',
        'Колонка_5': 'Статус_связи',
        'Колонка_6': 'Служба',
        'Колонка_7': 'Жалоба',
        'Колонка_8': 'Положительно'
    })
    
    # Находим нашу строку
    target_row = chunk[chunk['Инцидент_Sheets'] == '01.AAD4284/26']
    if not target_row.empty:
        print(f"\n✓ Нашли строку в chunk")
        print(f"  Индекс chunk: {target_row.index.tolist()}")
        print(f"  Дата: {target_row['Дата_открытия'].values[0]}")
        
        # Создаем date_series
        date_series = pd.to_datetime(chunk['Дата_открытия'], errors='coerce', dayfirst=True)
        print(f"\n  date_series.index (первые 5): {date_series.index[:5].tolist()}")
        print(f"  chunk.index (первые 5): {chunk.index[:5].tolist()}")
        print(f"  Индексы совпадают: {date_series.index.equals(chunk.index)}")
        
        # Создаем маску
        start_date = pd.Timestamp('2026-01-04')
        end_date = pd.Timestamp('2026-01-31 23:59:59')
        mask = (date_series >= start_date) & (date_series <= end_date)
        
        print(f"\n  Маска для нашей строки:")
        target_idx = target_row.index[0]
        print(f"    mask[{target_idx}] = {mask.loc[target_idx]}")
        print(f"    date_series[{target_idx}] = {date_series.loc[target_idx]}")
        
        # Применяем фильтр (как в коде)
        filtered = chunk[mask]
        
        print(f"\n  До фильтра: {len(chunk)} записей")
        print(f"  После фильтра: {len(filtered)} записей")
        print(f"  Наша строка в filtered: {target_idx in filtered.index}")
        
        # Проверяем маску детально
        print(f"\n  Проверка маски:")
        print(f"    True в маске: {mask.sum()}")
        print(f"    False в маске: {(~mask).sum()}")
        print(f"    NaN в date_series: {date_series.isna().sum()}")
        
        break
