"""
Фильтрация данных за 2025 год
"""

import pandas as pd
from datetime import datetime

print("="*80)
print("ФИЛЬТРАЦИЯ ДАННЫХ ЗА 2025 ГОД")
print("="*80)

# Загружаем собранные данные
input_file = "ALL_DATA_COLLECTED.csv"
output_file = "ALL_DATA_2025.csv"

print(f"\n📂 Загрузка файла: {input_file}")

try:
    df = pd.read_csv(input_file, encoding='utf-8-sig', low_memory=False)
    print(f"✅ Загружено записей: {len(df):,}")
    
    print(f"\n📋 Колонки в данных:")
    for i, col in enumerate(df.columns, 1):
        print(f"  {i}. {col}")
    
    # Используем ДАТУ ОТКРЫТИЯ КАРТЫ для фильтрации
    if 'Дата открытия карты' in df.columns:
        date_col = 'Дата открытия карты'
        print(f"\n🔍 Используем колонку: {date_col} (дата открытия)")
    else:
        print("\n❌ Не найдена колонка 'Дата открытия карты'!")
        date_columns = [col for col in df.columns if 'дата' in col.lower() or 'date' in col.lower()]
        print(f"Найдены колонки с датами: {date_columns}")
        if date_columns:
            date_col = date_columns[0]
            print(f"Используем: {date_col}")
        else:
            print("Показываю первые строки:")
            print(df.head(3))
            exit(1)
        
        # Конвертируем даты
        print("🔄 Конвертация дат...")
        
        # Пробуем разные форматы
        df[date_col] = pd.to_datetime(df[date_col], format='%d.%m.%Y %H:%M:%S', errors='coerce')
        
        # Если не получилось, пробуем другие форматы
        if df[date_col].isna().all():
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        # Проверяем сколько валидных дат
        valid_dates = df[date_col].notna().sum()
        print(f"✅ Валидных дат: {valid_dates:,} из {len(df):,}")
        
        # Фильтруем за 2025 год
        print("\n🔍 Фильтрация за 2025 год...")
        df_2025 = df[df[date_col].dt.year == 2025].copy()
        
        print(f"✅ Записей за 2025 год: {len(df_2025):,}")
        
        if len(df_2025) > 0:
            # Сохраняем
            print(f"\n💾 Сохранение в: {output_file}")
            df_2025.to_csv(output_file, index=False, encoding='utf-8-sig')
            
            import os
            file_size = os.path.getsize(output_file) / (1024*1024)
            
            print(f"✅ ГОТОВО!")
            print(f"📊 Записей за 2025: {len(df_2025):,}")
            print(f"💾 Размер файла: {file_size:.2f} МБ")
            
            # Статистика
            print(f"\n📈 СТАТИСТИКА ЗА 2025 ГОД:")
            
            if 'Номер карты' in df_2025.columns:
                print(f"  🎫 Уникальных карт: {df_2025['Номер карты'].nunique():,}")
            
            if 'Оператор' in df_2025.columns:
                print(f"  👥 Операторов: {df_2025['Оператор'].nunique()}")
            
            if 'Статус' in df_2025.columns:
                print(f"\n  📊 Топ-5 статусов:")
                for status, count in df_2025['Статус'].value_counts().head(5).items():
                    print(f"    {status}: {count:,}")
            
            # Распределение по месяцам
            print(f"\n  📅 Распределение по месяцам:")
            df_2025['Месяц'] = df_2025[date_col].dt.month
            months_dict = {1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 
                          5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август', 
                          9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'}
            
            for month_num in sorted(df_2025['Месяц'].unique()):
                count = len(df_2025[df_2025['Месяц'] == month_num])
                month_name = months_dict.get(month_num, f'Месяц {month_num}')
                print(f"    {month_name}: {count:,}")
            
            # Диапазон дат
            print(f"\n  📆 Диапазон дат:")
            print(f"    С: {df_2025[date_col].min()}")
            print(f"    По: {df_2025[date_col].max()}")
        else:
            print("\n❌ Нет данных за 2025 год!")
            print("Проверим какие года есть в данных:")
            df['Год'] = df[date_col].dt.year
            print(df['Год'].value_counts().sort_index())

except FileNotFoundError:
    print(f"❌ Файл {input_file} не найден!")
    print("Сначала запустите export_all_sheets.py для сбора данных")
except Exception as e:
    print(f"❌ Ошибка: {e}")
    import traceback
    traceback.print_exc()
