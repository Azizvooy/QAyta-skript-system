"""
Конвертер CSV в Excel с разбивкой на несколько листов
(так как CSV слишком большой для одного листа Excel)
"""
import pandas as pd
import os

print("📊 Конвертация ALL_DATA.csv → Excel...\n")

# Читаем CSV
print("📖 Чтение CSV файла...")
df = pd.read_csv('ALL_DATA.csv', encoding='utf-8-sig')
print(f"✅ Загружено {len(df):,} строк\n")

# Максимум строк в одном листе Excel (оставляем запас)
MAX_ROWS = 1000000

# Если данных меньше лимита - сохраняем в один лист
if len(df) <= MAX_ROWS:
    print("💾 Сохранение в один лист...")
    with pd.ExcelWriter('ALL_DATA_FULL.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Все данные', index=False)
    print("✅ Готово: ALL_DATA_FULL.xlsx")
else:
    # Разбиваем на несколько листов
    num_parts = (len(df) // MAX_ROWS) + 1
    print(f"📦 Разбиваем на {num_parts} листов...\n")
    
    with pd.ExcelWriter('ALL_DATA_MULTI.xlsx', engine='openpyxl') as writer:
        for i in range(num_parts):
            start_idx = i * MAX_ROWS
            end_idx = min((i + 1) * MAX_ROWS, len(df))
            
            df_part = df.iloc[start_idx:end_idx]
            sheet_name = f'Часть {i+1}'
            
            print(f"  💾 {sheet_name}: строки {start_idx+1:,} - {end_idx:,}")
            df_part.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\n✅ Готово: ALL_DATA_MULTI.xlsx ({num_parts} листов)")

print(f"\n📈 Статистика:")
print(f"   Всего строк: {len(df):,}")
print(f"   Операторов: {df['Оператор'].nunique()}")
print(f"   Уникальных карт: {df['Номер карты'].nunique():,}")
