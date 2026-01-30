"""
Правильная обработка экспортированных CSV файлов
"""
import os
import pandas as pd
from tqdm import tqdm
from pathlib import Path

print("=" * 80)
print("ПРАВИЛЬНАЯ ОБРАБОТКА ДАННЫХ")
print("=" * 80)

# Папка с экспортированными файлами
export_folder = Path(r'C:\Users\a.djurayev\Desktop\QAyta skript\exported_sheets')

# Ключевые колонки которые нам нужны
KEY_COLUMNS = [
    'Номер карты', 'Код карты', '№', 'Кодкарты',  # Варианты для номера карты
    'Статус связи', 'Причина/Статус',  # Варианты для статуса
    'Оператор', 'USER', 'Пользователь',  # Варианты для оператора
    'Дата фиксации', 'Время фиксации', 'Дата открытия карты'  # Варианты для дат
]

all_data = []
file_count = 0
error_count = 0

# Проходим по всем файлам
csv_files = list(export_folder.rglob('*.csv'))
print(f"\n📂 Найдено CSV файлов: {len(csv_files)}")

for csv_file in tqdm(csv_files, desc="Обработка файлов"):
    try:
        # Читаем файл
        df = pd.read_csv(csv_file, encoding='utf-8-sig', low_memory=False)
        
        # Пропускаем пустые файлы
        if len(df) == 0:
            continue
            
        # Определяем какие колонки есть в этом файле
        номер_карты_col = None
        статус_col = None
        оператор_col = None
        дата_col = None
        
        # Ищем колонку с номером карты
        for col in ['Номер карты', 'Код карты', 'Кодкарты']:
            if col in df.columns:
                номер_карты_col = col
                break
        
        # Ищем колонку со статусом
        for col in ['Статус связи', 'Причина/Статус']:
            if col in df.columns:
                статус_col = col
                break
        
        # Ищем колонку с оператором
        for col in ['Оператор', 'USER', 'Пользователь']:
            if col in df.columns:
                оператор_col = col
                break
        
        # Ищем колонку с датой
        for col in ['Дата фиксации', 'Время фиксации', 'Дата открытия карты']:
            if col in df.columns:
                дата_col = col
                break
        
        # Если нет нужных колонок - пропускаем
        if not номер_карты_col and not статус_col:
            continue
        
        # Создаем нормализованный DataFrame
        normalized_df = pd.DataFrame()
        
        if номер_карты_col:
            normalized_df['Номер карты'] = df[номер_карты_col]
        else:
            normalized_df['Номер карты'] = None
            
        if статус_col:
            normalized_df['Статус связи'] = df[статус_col]
        else:
            normalized_df['Статус связи'] = None
            
        if оператор_col:
            normalized_df['Оператор'] = df[оператор_col]
        else:
            # Пробуем извлечь имя оператора из пути к файлу
            operator_name = csv_file.parent.name
            if operator_name != 'exported_sheets':
                normalized_df['Оператор'] = operator_name
            else:
                normalized_df['Оператор'] = None
        
        if дата_col:
            normalized_df['Дата'] = df[дата_col]
        else:
            normalized_df['Дата'] = None
        
        # Добавляем источник (имя файла)
        normalized_df['Источник'] = csv_file.name
        
        # Добавляем в общий список
        all_data.append(normalized_df)
        file_count += 1
        
    except Exception as e:
        error_count += 1
        print(f"\n⚠️ Ошибка в {csv_file.name}: {str(e)[:100]}")

print(f"\n✅ Обработано файлов: {file_count}")
print(f"⚠️ Ошибок: {error_count}")

# Объединяем все данные
if all_data:
    print(f"\n🔄 Объединение данных...")
    combined_df = pd.concat(all_data, ignore_index=True)
    
    print(f"📊 Всего записей: {len(combined_df):,}")
    
    # Фильтруем записи с номером карты
    df_clean = combined_df[combined_df['Номер карты'].notna()].copy()
    print(f"📋 Записей с номером карты: {len(df_clean):,}")
    
    # Сохраняем
    output_file = r'C:\Users\a.djurayev\Desktop\QAyta skript\ALL_DATA_FIXED.csv'
    combined_df.to_csv(output_file, index=False, encoding='utf-8-sig')
    
    file_size = os.path.getsize(output_file) / (1024 * 1024)
    print(f"\n💾 Сохранено в: {output_file}")
    print(f"📦 Размер: {file_size:.2f} МБ")
    
    # Быстрая статистика
    print(f"\n📊 БЫСТРАЯ СТАТИСТИКА:")
    print(f"   Уникальных карт: {df_clean['Номер карты'].nunique():,}")
    
    if 'Оператор' in df_clean.columns:
        ops_count = df_clean['Оператор'].nunique()
        print(f"   Операторов: {ops_count}")
    
    if 'Статус связи' in df_clean.columns:
        print(f"\n   ТОП-5 статусов:")
        for i, (status, count) in enumerate(df_clean['Статус связи'].value_counts().head(5).items(), 1):
            print(f"   {i}. {status}: {count:,}")
    
    print("\n" + "=" * 80)
    print("✅ ГОТОВО!")
    print("=" * 80)
else:
    print("\n❌ Нет данных для обработки!")
