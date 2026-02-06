import pandas as pd

sheets_file = 'data/КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_20260202_130742.csv'

print("Проверка парсинга всего столбца дат")
print("=" * 60)

for chunk in pd.read_csv(sheets_file, chunksize=200000, encoding='utf-8-sig'):
    chunk = chunk.rename(columns={
        'Колонка_2': 'Инцидент_Sheets',
        'Колонка_4': 'Дата_открытия',
    })
    
    target = chunk[chunk['Инцидент_Sheets'] == '01.AAD4284/26']
    if not target.empty:
        print(f"\n✓ Нашли инцидент в chunk")
        target_idx = target.index[0]
        
        # Парсим весь столбец
        print(f"\nПарсинг всего столбца date_series...")
        date_series = pd.to_datetime(chunk['Дата_открытия'], errors='coerce', dayfirst=True)
        
        print(f"  Общий результат:")
        print(f"    Всего записей: {len(date_series)}")
        print(f"    NaT записей: {date_series.isna().sum()}")
        print(f"    Валидных дат: {date_series.notna().sum()}")
        
        # Проверяем нашу конкретную запись
        print(f"\n  Для инцидента '01.AAD4284/26' (индекс {target_idx}):")
        print(f"    Исходная дата: '{chunk.loc[target_idx, 'Дата_открытия']}'")
        print(f"    Парсированная: {date_series.loc[target_idx]}")
        print(f"    Это NaT: {pd.isna(date_series.loc[target_idx])}")
        
        # Проверяем что не так с датами
        print(f"\n  Примеры проблемных дат (NaT):")
        nat_mask = date_series.isna()
        problematic = chunk.loc[nat_mask, 'Дата_открытия'].head(10)
        for idx, val in problematic.items():
            print(f"    [{idx}] '{val}' (тип: {type(val)})")
        
        # Проверяем диапазон дат
        print(f"\n  Диапазон валидных дат в chunk:")
        valid_dates = date_series[date_series.notna()]
        if len(valid_dates) > 0:
            print(f"    Минимум: {valid_dates.min()}")
            print(f"    Максимум: {valid_dates.max()}")
        
        # Проверяем фильтр
        start_date = pd.Timestamp('2026-01-04')
        end_date = pd.Timestamp('2026-01-31 23:59:59')
        mask = (date_series >= start_date) & (date_series <= end_date)
        
        print(f"\n  Фильтр по датам (04-31.01.2026):")
        print(f"    Записей прошедших фильтр: {mask.sum()}")
        print(f"    Наша запись прошла: {mask.loc[target_idx] if target_idx in mask.index else 'НЕТ ИНДЕКСА'}")
        
        break
