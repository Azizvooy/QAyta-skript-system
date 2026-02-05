import pandas as pd

sheets_file = 'data/КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_20260202_130742.csv'

print("Проверка формата дат в Sheets")
print("=" * 60)

for chunk in pd.read_csv(sheets_file, chunksize=200000, encoding='utf-8-sig'):
    chunk = chunk.rename(columns={
        'Колонка_2': 'Инцидент_Sheets',
        'Колонка_4': 'Дата_открытия',
    })
    
    target = chunk[chunk['Инцидент_Sheets'] == '01.AAD4284/26']
    if not target.empty:
        date_value = target['Дата_открытия'].values[0]
        print(f"\nДата из Sheets:")
        print(f"  Значение: '{date_value}'")
        print(f"  Тип: {type(date_value)}")
        print(f"  Repr: {repr(date_value)}")
        
        # Пробуем разные способы парсинга
        print(f"\nПопытки парсинга:")
        
        # Текущий способ
        try:
            parsed1 = pd.to_datetime(date_value, dayfirst=True)
            print(f"  pd.to_datetime(dayfirst=True): {parsed1}")
        except Exception as e:
            print(f"  pd.to_datetime(dayfirst=True): ОШИБКА - {e}")
        
        # Без dayfirst
        try:
            parsed2 = pd.to_datetime(date_value, dayfirst=False)
            print(f"  pd.to_datetime(dayfirst=False): {parsed2}")
        except Exception as e:
            print(f"  pd.to_datetime(dayfirst=False): ОШИБКА - {e}")
        
        # С форматом
        try:
            parsed3 = pd.to_datetime(date_value, format='%d.%m.%Y %H:%M:%S')
            print(f"  pd.to_datetime(format='%d.%m.%Y %H:%M:%S'): {parsed3}")
        except Exception as e:
            print(f"  pd.to_datetime(format='%d.%m.%Y %H:%M:%S'): ОШИБКА - {e}")
        
        # Проверим несколько других записей
        print(f"\nПервые 5 дат в этом chunk:")
        for idx, date_str in chunk['Дата_открытия'].head().items():
            try:
                parsed = pd.to_datetime(date_str, dayfirst=True)
                status = "✓"
            except:
                parsed = "NaT"
                status = "✗"
            print(f"  {status} '{date_str}' -> {parsed}")
        
        break
