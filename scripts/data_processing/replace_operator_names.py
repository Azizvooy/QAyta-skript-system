"""
Замена номеров операторов (Оператор 12, 13...) на реальные ФИО из колонки "Оператор фиксировавший"
"""
import pandas as pd

print("=" * 80)
print("ЗАМЕНА НОМЕРОВ ОПЕРАТОРОВ НА РЕАЛЬНЫЕ ФИО")
print("=" * 80)

# Загружаем данные
print("\n📂 Загрузка данных...")
df = pd.read_csv('ALL_DATA_FIXED.csv', encoding='utf-8-sig')
print(f"✅ Загружено строк: {len(df):,}")

# Создаем маппинг номеров операторов к ФИО
print("\n🔍 Анализ операторов...")

# Находим все строки где в "Оператор" стоит номер, а в "Оператор фиксировавший" есть ФИО
operator_mapping = {}

for idx, row in df.iterrows():
    operator_col = str(row['Оператор']).strip()
    fix_operator_col = str(row['Оператор фиксировавший']).strip()
    
    # Если в первой колонке номер (например "Оператор 12"), а во второй ФИО
    if operator_col.startswith('Оператор ') and fix_operator_col and fix_operator_col != operator_col:
        # Проверяем что fix_operator_col это ФИО (содержит буквы, не только "Оператор")
        if not fix_operator_col.startswith('Оператор ') and len(fix_operator_col) > 5:
            if operator_col not in operator_mapping:
                operator_mapping[operator_col] = {}
            
            # Считаем сколько раз этот оператор использовал это ФИО
            if fix_operator_col not in operator_mapping[operator_col]:
                operator_mapping[operator_col][fix_operator_col] = 0
            operator_mapping[operator_col][fix_operator_col] += 1

# Выбираем самое частое ФИО для каждого номера
final_mapping = {}
for operator_num, fio_counts in operator_mapping.items():
    if fio_counts:
        # Берем ФИО которое встречается чаще всего
        most_common_fio = max(fio_counts.items(), key=lambda x: x[1])
        final_mapping[operator_num] = most_common_fio[0]
        print(f"  {operator_num} → {most_common_fio[0]} ({most_common_fio[1]} раз)")

if not final_mapping:
    print("\n⚠️  Не найдено соответствий для замены!")
    print("   Все операторы уже имеют ФИО")
else:
    print(f"\n📊 Найдено соответствий: {len(final_mapping)}")
    
    # Применяем замены
    print("\n🔄 Применение замен...")
    replaced_count = 0
    
    for old_name, new_name in final_mapping.items():
        mask = df['Оператор'] == old_name
        count = mask.sum()
        if count > 0:
            df.loc[mask, 'Оператор'] = new_name
            replaced_count += count
            print(f"  ✓ {old_name} → {new_name} ({count:,} строк)")
    
    # Также обновляем колонку "Оператор фиксировавший" если там тоже номера
    print("\n🔄 Обновление колонки 'Оператор фиксировавший'...")
    fix_replaced = 0
    for old_name, new_name in final_mapping.items():
        mask = df['Оператор фиксировавший'] == old_name
        count = mask.sum()
        if count > 0:
            df.loc[mask, 'Оператор фиксировавший'] = new_name
            fix_replaced += count
    
    print(f"  ✓ Обновлено: {fix_replaced:,} строк")
    
    # Сохраняем
    print("\n💾 Сохранение результата...")
    df.to_csv('ALL_DATA_FIXED.csv', index=False, encoding='utf-8-sig')
    print("✅ Сохранено в: ALL_DATA_FIXED.csv")
    
    # Итоговая статистика
    print("\n" + "=" * 80)
    print("📊 ИТОГОВАЯ СТАТИСТИКА")
    print("=" * 80)
    print(f"Всего операторов: {df['Оператор'].nunique()}")
    print(f"Заменено строк в 'Оператор': {replaced_count:,}")
    print(f"Заменено строк в 'Оператор фиксировавший': {fix_replaced:,}")
    
    print("\n👥 Список операторов после замены:")
    for i, op in enumerate(sorted(df['Оператор'].unique()), 1):
        count = (df['Оператор'] == op).sum()
        print(f"{i:3}. {op} ({count:,} заявок)")

print("\n" + "=" * 80)
print("✅ ГОТОВО!")
print("=" * 80)
