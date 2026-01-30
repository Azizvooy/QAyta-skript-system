import pandas as pd
import os

# Путь к файлу
file_path = r'c:\Users\a.djurayev\Desktop\QAyta skript\ALL_DATA_CLEANED.csv'

print("Загрузка данных...")
print(f"Размер файла: {os.path.getsize(file_path) / (1024*1024):.2f} МБ")

# Загружаем данные
df = pd.read_csv(file_path, encoding='utf-8-sig', low_memory=False)

print(f"\nВсего строк в файле: {len(df):,}")
print(f"\nКолонки: {list(df.columns)}")
print("\nПервые строки:")
print(df.head())

# Определяем колонку с номером заявки
# Обычно это может быть: ID, Номер заявки, №, Application ID и т.д.
print("\n" + "="*80)
print("АНАЛИЗ ДАННЫХ")
print("="*80)

# Ищем колонку с результатом звонка
result_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in 
                  ['результат', 'статус', 'ответ', 'result', 'status', 'answer'])]
print(f"\nНайденные колонки с результатами: {result_columns}")

# Ищем колонку с номером заявки
id_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in 
              ['заявк', 'id', 'номер', '№', 'application', 'number'])]
print(f"Найденные колонки с ID заявки: {id_columns}")

# Подсчет уникальных заявок
if id_columns:
    main_id_col = id_columns[0]
    unique_applications = df[main_id_col].nunique()
    print(f"\n📊 Уникальных заявок (по колонке '{main_id_col}'): {unique_applications:,}")
    
    # Подсчет всех записей
    total_calls = len(df)
    print(f"📞 Всего записей звонков: {total_calls:,}")

# Анализ результатов
if result_columns:
    for col in result_columns[:3]:  # Берем первые 3 колонки с результатами
        print(f"\n--- Анализ колонки: {col} ---")
        print(df[col].value_counts().head(20))
        
        # Попытка классифицировать на положительные/отрицательные
        positive_keywords = ['да', 'yes', 'согласен', 'положит', 'подтвержд', 'успеш', 'готов', 'принят']
        negative_keywords = ['нет', 'no', 'отказ', 'отрицат', 'не согласен', 'недоступ', 'не отвечает']
        
        df_temp = df[col].astype(str).str.lower()
        positive = df_temp[df_temp.apply(lambda x: any(kw in x for kw in positive_keywords))].count()
        negative = df_temp[df_temp.apply(lambda x: any(kw in x for kw in negative_keywords))].count()
        
        print(f"\n✅ Положительные ответы: {positive:,}")
        print(f"❌ Отрицательные ответы: {negative:,}")
        print(f"❓ Прочие: {len(df) - positive - negative:,}")

# Дополнительная статистика
print("\n" + "="*80)
print("ДОПОЛНИТЕЛЬНАЯ СТАТИСТИКА")
print("="*80)

# Выводим информацию о данных
print(f"\nИнформация о типах данных:")
print(df.dtypes)

print("\nСтатистика по пропущенным значениям:")
missing = df.isnull().sum()
missing = missing[missing > 0].sort_values(ascending=False)
if len(missing) > 0:
    print(missing.head(10))
else:
    print("Пропущенных значений нет")
