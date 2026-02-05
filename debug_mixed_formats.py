import pandas as pd

sheets_file = 'data/КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_20260202_130742.csv'

print("Детальная проверка парсинга")
print("=" * 60)

for chunk in pd.read_csv(sheets_file, chunksize=200000, encoding='utf-8-sig'):
    chunk = chunk.rename(columns={'Колонка_4': 'Дата_открытия'})
    
    # Берём проблемную дату
    test_dates = [
        '05.01.2026 20:38:17',
        '04.01.2026 17:02:36',
        '10.6.2025 22:04',
    ]
    
    print("\nТест парсинга отдельных строк:")
    for date_str in test_dates:
        result_dayfirst = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
        result_no_dayfirst = pd.to_datetime(date_str, dayfirst=False, errors='coerce')
        print(f"  '{date_str}'")
        print(f"    dayfirst=True:  {result_dayfirst}")
        print(f"    dayfirst=False: {result_no_dayfirst}")
    
    # Создаём небольшой DataFrame с этими датами
    print("\nТест парсинга DataFrame:")
    test_df = pd.DataFrame({'Дата': test_dates})
    test_df['Parsed_dayfirst'] = pd.to_datetime(test_df['Дата'], dayfirst=True, errors='coerce')
    test_df['Parsed_no_dayfirst'] = pd.to_datetime(test_df['Дата'], dayfirst=False, errors='coerce')
    print(test_df.to_string())
    
    # Проверяем первые 20 записей настоящего chunk
    print("\nПервые 20 записей chunk:")
    sample = chunk.head(20).copy()
    sample['Parsed'] = pd.to_datetime(sample['Дата_открытия'], dayfirst=True, errors='coerce')
    print(sample[['Дата_открытия', 'Parsed']].to_string())
    
    # Проверяем, что находится в индексе 282631
    if 282631 in chunk.index:
        print(f"\nЗапись с индексом 282631:")
        print(f"  Дата: '{chunk.loc[282631, 'Дата_открытия']}'")
        single_parse = pd.to_datetime(chunk.loc[282631, 'Дата_открытия'], dayfirst=True, errors='coerce')
        print(f"  Парсинг одной строки: {single_parse}")
        
        date_series = pd.to_datetime(chunk['Дата_открытия'], dayfirst=True, errors='coerce')
        print(f"  В date_series: {date_series.loc[282631]}")
    
    break
