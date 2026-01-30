import pandas as pd

# Путь к файлу
file_path = r'c:\Users\a.djurayev\Desktop\QAyta skript\ALL_DATA_CLEANED.csv'

print("Загрузка данных...")
df = pd.read_csv(file_path, encoding='utf-8-sig', low_memory=False)

print(f"Всего строк: {len(df):,}")

# Конвертируем даты
df['Дата фиксации'] = pd.to_datetime(df['Дата фиксации'], format='%d.%m.%Y %H:%M:%S', errors='coerce')

# Проверяем какие годы есть
print("\n" + "="*80)
print("РАСПРЕДЕЛЕНИЕ ПО ГОДАМ:")
print("="*80)
df['Год'] = df['Дата фиксации'].dt.year
print(df['Год'].value_counts().sort_index())

# Проверяем данные за 2025 год по месяцам
df_2025 = df[df['Год'] == 2025].copy()
print("\n" + "="*80)
print("РАСПРЕДЕЛЕНИЕ ПО МЕСЯЦАМ 2025 ГОДА:")
print("="*80)

if len(df_2025) > 0:
    df_2025['Месяц'] = df_2025['Дата фиксации'].dt.month
    months_dict = {1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь',
                   7: 'Июль', 8: 'Август', 9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'}
    
    month_counts = df_2025['Месяц'].value_counts().sort_index()
    
    print("\nВсе месяцы 2025 года:")
    for month in range(1, 13):
        count = month_counts.get(month, 0)
        if count > 0:
            print(f"{months_dict[month]:12} ({month:2}): {count:>10,} записей")
        else:
            print(f"{months_dict[month]:12} ({month:2}): {'0':>10} записей")
    
    print(f"\nВсего за 2025 год: {len(df_2025):,} записей")
    
    # Показываем диапазон дат
    print("\n" + "="*80)
    print("ДИАПАЗОН ДАТ В ДАННЫХ:")
    print("="*80)
    print(f"Самая ранняя дата: {df['Дата фиксации'].min()}")
    print(f"Самая поздняя дата: {df['Дата фиксации'].max()}")
    
    print(f"\nЗа 2025 год:")
    print(f"Самая ранняя дата: {df_2025['Дата фиксации'].min()}")
    print(f"Самая поздняя дата: {df_2025['Дата фиксации'].max()}")
else:
    print("Нет данных за 2025 год")
