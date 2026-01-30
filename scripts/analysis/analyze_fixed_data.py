"""
Финальный анализ правильно обработанных данных
"""
import pandas as pd
from datetime import datetime

print("=" * 80)
print("ФИНАЛЬНЫЙ АНАЛИЗ - УНИКАЛЬНЫЕ ЗАЯВКИ")
print("=" * 80)

# Загружаем исправленные данные
file_path = r'C:\Users\a.djurayev\Desktop\QAyta skript\ALL_DATA_FIXED.csv'
print(f"\n📂 Загрузка данных...")
df = pd.read_csv(file_path, encoding='utf-8-sig', low_memory=False)

print(f"📊 Всего записей: {len(df):,}")

# Фильтруем только записи с номером карты
df_clean = df[df['Номер карты'].notna()].copy()
print(f"📋 Записей с номером карты: {len(df_clean):,}")

# ДЕДУПЛИКАЦИЯ: берем последнюю запись для каждой уникальной карты
print(f"\n🔄 Дедупликация по номеру карты...")

# Пробуем преобразовать дату если она есть
if 'Дата' in df_clean.columns:
    df_clean['Дата_parsed'] = pd.to_datetime(df_clean['Дата'], errors='coerce', dayfirst=True)
    # Сортируем по дате если она есть, иначе просто берем последнюю запись
    df_clean_sorted = df_clean.sort_values('Дата_parsed', na_position='first')
    df_unique = df_clean_sorted.drop_duplicates(subset='Номер карты', keep='last')
else:
    # Просто берем последнюю запись
    df_unique = df_clean.drop_duplicates(subset='Номер карты', keep='last')

print(f"🎫 УНИКАЛЬНЫХ ЗАЯВОК (КАРТ): {len(df_unique):,}")
print("=" * 80)

# ========== АНАЛИЗ СТАТУСОВ ==========
print(f"\n📊 АНАЛИЗ ПО СТАТУСАМ:")

# Функция для классификации
def classify_status(status_text):
    if pd.isna(status_text):
        return 'Неизвестно'
    
    text = str(status_text).lower().strip()
    
    if 'положит' in text:
        return 'Положительный'
    elif 'отрицат' in text:
        return 'Отрицательный'
    elif 'нет ответа' in text or 'занято' in text:
        return 'Нет ответа/Занято'
    elif 'закрыта' in text or 'закрыт' in text:
        return 'Заявка закрыта'
    elif 'соед' in text or 'прервано' in text:
        return 'Соединение прервано'
    elif 'тишина' in text:
        return 'Тишина'
    elif 'тиббиёт' in text or 'ходими' in text:
        return 'Медработники'
    else:
        return 'Прочее'

df_unique['Категория'] = df_unique['Статус связи'].apply(classify_status)
categories = df_unique['Категория'].value_counts()

total = len(df_unique)

print(f"\n{'Категория':<30} {'Количество':>15} {'Процент':>10}")
print("-" * 80)
for cat, count in categories.items():
    print(f"{cat:<30} {count:>15,} {count/total*100:>9.2f}%")

# ========== ДОЗВОНИЛИСЬ / НЕ ДОЗВОНИЛИСЬ ==========
print(f"\n" + "=" * 80)
print("📞 СВОДКА ПО ДОЗВОНАМ:")
print("=" * 80)

positive = categories.get('Положительный', 0)
negative = categories.get('Отрицательный', 0)
no_answer = categories.get('Нет ответа/Занято', 0)
closed = categories.get('Заявка закрыта', 0)
disconnected = categories.get('Соединение прервано', 0)
silence = categories.get('Тишина', 0)
medical = categories.get('Медработники', 0)
other = categories.get('Прочее', 0)
unknown = categories.get('Неизвестно', 0)

# Дозвонились = положительные + отрицательные
dozonil = positive + negative

# Не дозвонились = все остальное
ne_dozonil = no_answer + closed + disconnected + silence

print(f"\n✅ ДОЗВОНИЛИСЬ и получили ответ: {dozonil:,} ({dozonil/total*100:.2f}%)")
print(f"   ├─ Положительных: {positive:,} ({positive/dozonil*100 if dozonil > 0 else 0:.2f}% от дозвонившихся)")
print(f"   └─ Отрицательных: {negative:,} ({negative/dozonil*100 if dozonil > 0 else 0:.2f}% от дозвонившихся)")

print(f"\n❌ НЕ дозвонились: {ne_dozonil:,} ({ne_dozonil/total*100:.2f}%)")
print(f"   ├─ Нет ответа/Занято: {no_answer:,}")
print(f"   ├─ Заявка закрыта: {closed:,}")
print(f"   ├─ Соединение прервано: {disconnected:,}")
print(f"   └─ Тишина: {silence:,}")

print(f"\n🏥 Медработники: {medical:,} ({medical/total*100:.2f}%)")

if other > 0:
    print(f"📝 Прочие: {other:,} ({other/total*100:.2f}%)")
if unknown > 0:
    print(f"❓ Неизвестно: {unknown:,} ({unknown/total*100:.2f}%)")

# ========== ТОП ОПЕРАТОРОВ ==========
print(f"\n" + "=" * 80)
print("👥 ТОП-20 ОПЕРАТОРОВ ПО КОЛИЧЕСТВУ УНИКАЛЬНЫХ ЗАЯВОК:")
print("=" * 80)

df_ops = df_unique[df_unique['Оператор'].notna() & (df_unique['Оператор'] != '-')]

if len(df_ops) > 0:
    ops_counts = df_ops['Оператор'].value_counts().head(20)
    
    print(f"\n{'№':<5} {'Оператор':<50} {'Заявок':>10}")
    print("-" * 80)
    for idx, (op, count) in enumerate(ops_counts.items(), 1):
        print(f"{idx:<5} {op:<50} {count:>10,}")
    
    total_ops = df_ops['Оператор'].nunique()
    print(f"\n   Всего уникальных операторов: {total_ops}")

# ========== ДЕТАЛЬНАЯ СТАТИСТИКА СТАТУСОВ ==========
print(f"\n" + "=" * 80)
print("📋 ТОП-20 СТАТУСОВ (детально):")
print("=" * 80)

status_counts = df_unique['Статус связи'].value_counts().head(20)

print(f"\n{'№':<5} {'Статус':<50} {'Количество':>15} {'%':>10}")
print("-" * 80)
for idx, (st, count) in enumerate(status_counts.items(), 1):
    print(f"{idx:<5} {str(st)[:50]:<50} {count:>15,} {count/total*100:>9.2f}%")

# Сохраняем уникальные карты
output_file = r'C:\Users\a.djurayev\Desktop\QAyta skript\UNIQUE_CARDS_FINAL.csv'
df_unique.to_csv(output_file, index=False, encoding='utf-8-sig')
print(f"\n💾 Уникальные карты сохранены: {output_file}")

print("\n" + "=" * 80)
print("✅ АНАЛИЗ ЗАВЕРШЕН")
print("=" * 80)

# Выводим итоги для отчета
print(f"\n📊 ИТОГИ ДЛЯ ОТЧЕТА:")
print(f"   🎫 Уникальных заявок: {total:,}")
print(f"   ✅ Положительных: {positive:,} ({positive/total*100:.2f}%)")
print(f"   ❌ Отрицательных: {negative:,} ({negative/total*100:.2f}%)")
print(f"   🚫 Закрытых: {closed:,} ({closed/total*100:.2f}%)")
print(f"   📞 Дозвонились: {dozonil:,} ({dozonil/total*100:.2f}%)")
print(f"   🎯 Конверсия: {positive/dozonil*100 if dozonil > 0 else 0:.2f}% (от дозвонившихся)")
