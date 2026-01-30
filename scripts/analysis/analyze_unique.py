"""
Анализ УНИКАЛЬНЫХ заявок (без дубликатов по звонкам)
"""

import pandas as pd

print("="*80)
print("СТАТИСТИКА ПО УНИКАЛЬНЫМ ЗАЯВКАМ ЗА 2025 ГОД")
print("="*80)

# Загружаем данные
df = pd.read_csv('ALL_DATA_2025.csv', encoding='utf-8-sig', low_memory=False)

print(f"\n📊 Всего записей звонков: {len(df):,}")

# Определяем колонки
номер_карты = 'Номер карты'
статус = 'Статус связи'
оператор = 'Оператор'
дата_фиксации = 'Дата фиксации'

# Проверяем наличие колонок
if номер_карты not in df.columns:
    print(f"❌ Колонка '{номер_карты}' не найдена!")
    exit(1)

# Убираем пустые номера карт
df_clean = df[df[номер_карты].notna()].copy()
print(f"📋 Записей с номером карты: {len(df_clean):,}")

# Считаем уникальные заявки
unique_cards = df_clean[номер_карты].nunique()
print(f"\n🎫 УНИКАЛЬНЫХ ЗАЯВОК: {unique_cards:,}")

# Для каждой уникальной карты берем последний статус (последний звонок)
if дата_фиксации in df_clean.columns:
    df_clean[дата_фиксации] = pd.to_datetime(df_clean[дата_фиксации], format='%d.%m.%Y %H:%M:%S', errors='coerce')
    # Сортируем по дате и берем последнюю запись для каждой карты
    df_unique = df_clean.sort_values(дата_фиксации).groupby(номер_карты).last().reset_index()
else:
    # Если нет даты фиксации, просто берем первую встречу каждой карты
    df_unique = df_clean.drop_duplicates(subset=номер_карты, keep='first')

print(f"📊 Уникальных заявок после дедупликации: {len(df_unique):,}")

# Анализируем статусы
if статус in df_unique.columns:
    print(f"\n{'='*80}")
    print("РЕЗУЛЬТАТЫ ПО УНИКАЛЬНЫМ ЗАЯВКАМ:")
    print(f"{'='*80}")
    
    # Классификация
    positive = df_unique[df_unique[статус].astype(str).str.lower().str.contains('положит', na=False)]
    negative = df_unique[df_unique[статус].astype(str).str.lower().str.contains('отрицат', na=False)]
    no_answer = df_unique[df_unique[статус].astype(str).str.lower().str.contains('нет ответа|занято', na=False)]
    closed = df_unique[df_unique[статус].astype(str).str.lower().str.contains('закрыта', na=False)]
    medical = df_unique[df_unique[статус].astype(str).str.lower().str.contains('тиббиёт|ходими', na=False)]
    disconnected = df_unique[df_unique[статус].astype(str).str.lower().str.contains('соед.прервано|прервано', na=False)]
    silence = df_unique[df_unique[статус].astype(str).str.lower().str.contains('тишина', na=False)]
    
    total = len(df_unique)
    
    print(f"\n✅ ПОЛОЖИТЕЛЬНЫЕ: {len(positive):,} ({len(positive)/total*100:.2f}%)")
    print(f"   Клиенты согласились на карту")
    
    print(f"\n❌ ОТРИЦАТЕЛЬНЫЕ: {len(negative):,} ({len(negative)/total*100:.2f}%)")
    print(f"   Клиенты отказались от карты")
    
    print(f"\n📞 НЕТ ОТВЕТА/ЗАНЯТО: {len(no_answer):,} ({len(no_answer)/total*100:.2f}%)")
    print(f"   Не удалось дозвониться")
    
    print(f"\n🚫 ЗАЯВКА ЗАКРЫТА: {len(closed):,} ({len(closed)/total*100:.2f}%)")
    print(f"   Закрыты после безуспешных попыток")
    
    print(f"\n📵 СОЕДИНЕНИЕ ПРЕРВАНО: {len(disconnected):,} ({len(disconnected)/total*100:.2f}%)")
    
    print(f"\n🔇 ТИШИНА: {len(silence):,} ({len(silence)/total*100:.2f}%)")
    
    print(f"\n🏥 МЕДРАБОТНИКИ: {len(medical):,} ({len(medical)/total*100:.2f}%)")
    
    # Итоги
    print(f"\n{'='*80}")
    print("ИТОГОВАЯ СВОДКА:")
    print(f"{'='*80}")
    
    dozonil = len(positive) + len(negative)
    ne_dozonil = len(no_answer) + len(closed) + len(disconnected) + len(silence)
    
    print(f"\n📊 Дозвонились и получили ответ: {dozonil:,} ({dozonil/total*100:.2f}%)")
    print(f"   ✅ Положительных: {len(positive):,} ({len(positive)/dozonil*100:.2f}% от дозвонившихся)")
    print(f"   ❌ Отрицательных: {len(negative):,} ({len(negative)/dozonil*100:.2f}% от дозвонившихся)")
    
    print(f"\n📵 НЕ дозвонились: {ne_dozonil:,} ({ne_dozonil/total*100:.2f}%)")
    
    print(f"\n🏥 Медработники: {len(medical):,} ({len(medical)/total*100:.2f}%)")
    
    # Топ статусов
    print(f"\n{'='*80}")
    print("ТОП-10 СТАТУСОВ (по уникальным заявкам):")
    print(f"{'='*80}")
    for idx, (st, count) in enumerate(df_unique[статус].value_counts().head(10).items(), 1):
        print(f"{idx:2}. {st:50} {count:>8,} ({count/total*100:5.2f}%)")

# Топ операторов
if оператор in df_unique.columns:
    print(f"\n{'='*80}")
    print("ТОП-10 ОПЕРАТОРОВ (по уникальным заявкам):")
    print(f"{'='*80}")
    
    # Убираем пустые значения и "-"
    df_ops = df_unique[df_unique[оператор].notna() & (df_unique[оператор] != '-')]
    
    for idx, (op, count) in enumerate(df_ops[оператор].value_counts().head(10).items(), 1):
        print(f"{idx:2}. {op:50} {count:>8,}")
