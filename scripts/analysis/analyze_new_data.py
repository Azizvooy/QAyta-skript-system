import pandas as pd

# Загружаем новые данные за 2025
file_path = 'ALL_DATA_2025.csv'

print("="*80)
print("ПОЛНАЯ СТАТИСТИКА ПО НОВЫМ ДАННЫМ ЗА 2025 ГОД")
print("="*80)

df = pd.read_csv(file_path, encoding='utf-8-sig', low_memory=False)

print(f"\n📊 Всего записей: {len(df):,}")

# Определяем нужные колонки
print(f"\nВсе колонки: {list(df.columns[:20])}")

# Ищем правильные колонки
номер_карты_col = 'Номер карты' if 'Номер карты' in df.columns else 'Код карты'
статус_col = None
for col in df.columns:
    if 'статус' in str(col).lower() and 'связи' in str(col).lower():
        статус_col = col
        break

if not статус_col:
    статус_col = 'Причина/Статус' if 'Причина/Статус' in df.columns else None

оператор_col = 'Оператор' if 'Оператор' in df.columns else None

print(f"\nИспользуемые колонки:")
print(f"  Номер карты: {номер_карты_col}")
print(f"  Статус: {статус_col}")
print(f"  Оператор: {оператор_col}")

if номер_карты_col and номер_карты_col in df.columns:
    unique_cards = df[номер_карты_col].nunique()
    print(f"\n🎫 Уникальных заявок: {unique_cards:,}")

if статус_col and статус_col in df.columns:
    print(f"\n📊 Статусы (топ-20):")
    print(df[статус_col].value_counts().head(20))
    
    # Анализ
    positive = df[df[статус_col].astype(str).str.lower().str.contains('положит', na=False)]
    negative = df[df[статус_col].astype(str).str.lower().str.contains('отрицат', na=False)]
    no_answer = df[df[статус_col].astype(str).str.lower().str.contains('нет ответа|занято', na=False)]
    closed = df[df[статус_col].astype(str).str.lower().str.contains('закрыта', na=False)]
    
    print(f"\n✅ Положительных: {len(positive):,} ({len(positive)/len(df)*100:.2f}%)")
    print(f"❌ Отрицательных: {len(negative):,} ({len(negative)/len(df)*100:.2f}%)")
    print(f"📞 Нет ответа/Занято: {len(no_answer):,} ({len(no_answer)/len(df)*100:.2f}%)")
    print(f"🚫 Закрытых: {len(closed):,} ({len(closed)/len(df)*100:.2f}%)")

if оператор_col and оператор_col in df.columns:
    print(f"\n👥 Топ-10 операторов:")
    for idx, (op, count) in enumerate(df[оператор_col].value_counts().head(10).items(), 1):
        print(f"  {idx}. {op}: {count:,}")
