"""
Проверка качества собранных данных
"""
import pandas as pd

print("=" * 80)
print("ПРОВЕРКА КАЧЕСТВА СОБРАННЫХ ДАННЫХ")
print("=" * 80)

file_path = r'c:\Users\a.djurayev\Desktop\QAyta skript\ALL_DATA_COLLECTED.csv'
df = pd.read_csv(file_path, encoding='utf-8-sig', low_memory=False)

print(f"\n📊 ОБЩАЯ ИНФОРМАЦИЯ:")
print(f"   Всего записей: {len(df):,}")
print(f"   Размер файла: {(file_path.__sizeof__() / 1024 / 1024):.2f} МБ")
print(f"   Колонок: {len(df.columns)}")

print(f"\n📋 КОЛОНКИ В ДАННЫХ:")
for i, col in enumerate(df.columns, 1):
    print(f"   {i}. {col}")

# Проверка ключевых колонок
номер_карты = 'Номер карты'
статус = 'Статус связи'
оператор = 'Оператор'
дата = 'Дата фиксации факта звонка'

print(f"\n🔍 ПРОВЕРКА КЛЮЧЕВЫХ ДАННЫХ:")

if номер_карты in df.columns:
    df_clean = df[df[номер_карты].notna()]
    print(f"   Записей с номером карты: {len(df_clean):,}")
    print(f"   Записей БЕЗ номера карты: {len(df) - len(df_clean):,}")
    
    unique_cards = df_clean[номер_карты].nunique()
    print(f"   Уникальных карт: {unique_cards:,}")
    
    # Примеры номеров карт
    print(f"\n   Примеры номеров карт:")
    for i, card in enumerate(df_clean[номер_карты].dropna().head(10), 1):
        print(f"      {i}. {card}")
else:
    print(f"   ❌ Колонка '{номер_карты}' НЕ НАЙДЕНА!")

if статус in df.columns:
    print(f"\n📞 СТАТУСЫ СВЯЗИ (ТОП-15):")
    status_counts = df[статус].value_counts().head(15)
    for i, (st, count) in enumerate(status_counts.items(), 1):
        print(f"   {i}. {st}: {count:,} ({count/len(df)*100:.2f}%)")
else:
    print(f"   ❌ Колонка '{статус}' НЕ НАЙДЕНА!")

if оператор in df.columns:
    df_ops = df[df[оператор].notna() & (df[оператор] != '-')]
    print(f"\n👥 ОПЕРАТОРЫ (ТОП-20):")
    ops_counts = df_ops[оператор].value_counts().head(20)
    for i, (op, count) in enumerate(ops_counts.items(), 1):
        print(f"   {i}. {op}: {count:,}")
    
    total_ops = df_ops[оператор].nunique()
    print(f"\n   Всего уникальных операторов: {total_ops}")
else:
    print(f"   ❌ Колонка '{оператор}' НЕ НАЙДЕНА!")

if дата in df.columns:
    df[дата] = pd.to_datetime(df[дата], errors='coerce', dayfirst=True)
    df_dated = df[df[дата].notna()]
    print(f"\n📅 ДАТЫ:")
    print(f"   Записей с датой: {len(df_dated):,}")
    print(f"   Записей БЕЗ даты: {len(df) - len(df_dated):,}")
    
    if len(df_dated) > 0:
        print(f"   Самая ранняя дата: {df_dated[дата].min()}")
        print(f"   Самая поздняя дата: {df_dated[дата].max()}")
        
        # Распределение по годам
        print(f"\n   Распределение по годам:")
        year_dist = df_dated[дата].dt.year.value_counts().sort_index()
        for year, count in year_dist.items():
            print(f"      {year}: {count:,}")
else:
    print(f"   ❌ Колонка '{дата}' НЕ НАЙДЕНА!")

print("\n" + "=" * 80)
print("✅ ПРОВЕРКА ЗАВЕРШЕНА")
print("=" * 80)
