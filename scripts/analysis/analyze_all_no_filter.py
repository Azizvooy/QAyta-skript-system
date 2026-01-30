"""
Анализ ВСЕХ данных без фильтрации по дате
"""

import pandas as pd

print("="*80)
print("ПОЛНАЯ СТАТИСТИКА ПО ВСЕМ СОБРАННЫМ ДАННЫМ")
print("="*80)

df = pd.read_csv('ALL_DATA_COLLECTED.csv', encoding='utf-8-sig', low_memory=False)

print(f"\n📊 Всего записей: {len(df):,}")

# Ищем колонки
номер_карты = None
статус = None

for col in df.columns:
    if 'номер' in str(col).lower() and 'карт' in str(col).lower():
        номер_карты = col
    if 'статус' in str(col).lower() and 'связ' in str(col).lower():
        статус = col

print(f"\nКолонки:")
print(f"  Номер карты: {номер_карты}")
print(f"  Статус: {статус}")

if номер_карты:
    # Убираем пустые
    df_clean = df[df[номер_карты].notna()].copy()
    print(f"\n📋 Записей с номером карты: {len(df_clean):,}")
    
    # Уникальные карты
    unique = df_clean[номер_карты].nunique()
    print(f"🎫 УНИКАЛЬНЫХ КАРТ: {unique:,}")
    
    # Берем последний статус для каждой карты
    df_unique = df_clean.drop_duplicates(subset=номер_карты, keep='last')
    print(f"📊 Уникальных после дедупликации: {len(df_unique):,}")

if статус and номер_карты:
    print(f"\n{'='*80}")
    print("СТАТУСЫ ПО УНИКАЛЬНЫМ КАРТАМ:")
    print(f"{'='*80}")
    
    df_unique = df_clean.drop_duplicates(subset=номер_карты, keep='last')
    
    positive = df_unique[df_unique[статус].astype(str).str.lower().str.contains('положит', na=False)]
    negative = df_unique[df_unique[статус].astype(str).str.lower().str.contains('отрицат', na=False)]
    no_answer = df_unique[df_unique[статус].astype(str).str.lower().str.contains('нет ответа|занято', na=False)]
    closed = df_unique[df_unique[статус].astype(str).str.lower().str.contains('закрыта', na=False)]
    medical = df_unique[df_unique[статус].astype(str).str.lower().str.contains('тиббиёт|ходими', na=False)]
    
    total = len(df_unique)
    
    print(f"\n✅ ПОЛОЖИТЕЛЬНЫЕ: {len(positive):,} ({len(positive)/total*100:.2f}%)")
    print(f"❌ ОТРИЦАТЕЛЬНЫЕ: {len(negative):,} ({len(negative)/total*100:.2f}%)")
    print(f"📞 НЕТ ОТВЕТА: {len(no_answer):,} ({len(no_answer)/total*100:.2f}%)")
    print(f"🚫 ЗАКРЫТА: {len(closed):,} ({len(closed)/total*100:.2f}%)")
    print(f"🏥 МЕДРАБОТНИКИ: {len(medical):,} ({len(medical)/total*100:.2f}%)")
    
    dozonil = len(positive) + len(negative)
    print(f"\n📊 Дозвонились: {dozonil:,} ({dozonil/total*100:.2f}%)")
    print(f"   ✅ Положительных: {len(positive):,} ({len(positive)/dozonil*100:.2f}% от дозвонившихся)")
    print(f"   ❌ Отрицательных: {len(negative):,} ({len(negative)/dozonil*100:.2f}% от дозвонившихся)")
    
    print(f"\n{'='*80}")
    print("ТОП-10 СТАТУСОВ:")
    print(f"{'='*80}")
    for idx, (st, count) in enumerate(df_unique[статус].value_counts().head(10).items(), 1):
        print(f"{idx:2}. {st:50} {count:>10,} ({count/total*100:5.2f}%)")
