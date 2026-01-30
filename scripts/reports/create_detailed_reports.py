"""
Создание 2 отчетов:
1. Положительные - по статусам (лист 1: отчет, лист 2: детали)
2. Отрицательные - по комментариям (лист 1: отчет, лист 2: детали)
"""
import pandas as pd
from datetime import datetime
from pathlib import Path
import warnings
import sys
warnings.filterwarnings('ignore')

# Установка базовых путей
BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / 'data'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'

print("=" * 80)
print("СОЗДАНИЕ ДЕТАЛЬНЫХ ОТЧЕТОВ")
print("=" * 80)

# Загружаем данные
print("\n📂 Загрузка данных...")
data_file = DATA_DIR / 'ALL_DATA_FIXED.csv'

if not data_file.exists():
    print(f"❌ ОШИБКА: Файл не найден: {data_file}")
    sys.exit(1)

try:
    df = pd.read_csv(data_file, encoding='utf-8-sig')
    print(f"✅ Загружено строк: {len(df):,}")
except Exception as e:
    print(f"❌ ОШИБКА при загрузке данных: {e}")
    sys.exit(1)

# Функция для извлечения года и месяца
def extract_year_month(date_str):
    if pd.isna(date_str) or not str(date_str).strip():
        return None, None
    
    try:
        date_str = str(date_str).strip()
        
        if '.' in date_str:
            parts = date_str.split('.')
            if len(parts) >= 3:
                day, month, year = parts[0], parts[1], parts[2].split()[0]
                return year, month
        
        if '/' in date_str:
            parts = date_str.split('/')
            if len(parts) >= 3:
                day, month, year = parts[0], parts[1], parts[2].split()[0]
                return year, month
        
        if '-' in date_str:
            parts = date_str.split('-')
            if len(parts) >= 3:
                return parts[0], parts[1]
    except:
        pass
    
    return None, None

# Категоризация статусов
def categorize_status(status):
    if pd.isna(status):
        return 'Прочее'
    status_str = str(status).strip().lower()
    
    positive_keywords = ['положительн', 'qanoatlantir', 'қаноатлантир']
    
    for kw in positive_keywords:
        if kw in status_str:
            return 'Положительный'
    
    return 'Отрицательный'

print("\n🔄 Обработка данных...")
df['Год'], df['Месяц'] = zip(*df['Дата открытия карты'].apply(extract_year_month))
df['Категория'] = df['Статус'].apply(categorize_status)

# Фильтруем только строки с датой
df_with_date = df[df['Год'].notna()].copy()
print(f"✅ Строк с датой: {len(df_with_date):,}")

# Словарь месяцев
month_names = {
    '01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель',
    '05': 'Май', '06': 'Июнь', '07': 'Июль', '08': 'Август',
    '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'
}

df_with_date['Месяц_Название'] = df_with_date['Месяц'].map(month_names)

# ============================================================================
# ОТЧЕТ 1: ПОЛОЖИТЕЛЬНЫЕ СТАТУСЫ
# ============================================================================
print("\n" + "=" * 80)
print("📊 ОТЧЕТ 1: ПОЛОЖИТЕЛЬНЫЕ СТАТУСЫ")
print("=" * 80)

positive_df = df_with_date[df_with_date['Категория'] == 'Положительный'].copy()
print(f"Положительных записей: {len(positive_df):,}")

# Сводка по статусам
positive_report = positive_df.groupby(['Год', 'Месяц_Название', 'Служба', 'Статус']).agg({
    'Номер карты': 'count',
    'Оператор': 'nunique',
    'Оператор фиксировавший': 'nunique'
}).reset_index()

positive_report.columns = ['Год', 'Месяц', 'Служба', 'Статус', 'Количество заявок', 'Операторов', 'Фиксировавших']

# Сортируем
month_order = {'Январь': 1, 'Февраль': 2, 'Март': 3, 'Апрель': 4, 'Май': 5, 'Июнь': 6,
               'Июль': 7, 'Август': 8, 'Сентябрь': 9, 'Октябрь': 10, 'Ноябрь': 11, 'Декабрь': 12}
positive_report['Порядок'] = positive_report['Месяц'].map(month_order)
positive_report = positive_report.sort_values(['Год', 'Порядок', 'Служба', 'Количество заявок'], ascending=[True, True, True, False])
positive_report = positive_report.drop('Порядок', axis=1)

# Детальные записи
positive_details = positive_df[[
    'Год', 'Месяц_Название', 'Оператор', 'Архивный лист', 
    'Номер карты', 'Номер телефона', 'Дата открытия карты',
    'Статус', 'Служба', 'Комментарий', 
    'Оператор фиксировавший', 'Дата фиксации'
]].copy()
positive_details.columns = [
    'Год', 'Месяц', 'Оператор', 'Архивный лист',
    'Номер карты', 'Номер телефона', 'Дата открытия',
    'Статус', 'Служба', 'Комментарий',
    'Фиксировавший оператор', 'Дата фиксации'
]
positive_details['Порядок'] = positive_details['Месяц'].map(month_order)
positive_details = positive_details.sort_values(['Год', 'Порядок', 'Служба', 'Дата открытия'])
positive_details = positive_details.drop('Порядок', axis=1)

# Создаем директорию для отчетов
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Сохраняем
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
positive_file = OUTPUT_DIR / f'ОТЧЕТ_ПОЛОЖИТЕЛЬНЫЕ_{timestamp}.xlsx'

print("\n💾 Сохранение отчета по положительным...")
try:
    with pd.ExcelWriter(positive_file, engine='openpyxl') as writer:
        positive_report.to_excel(writer, sheet_name='Отчетность', index=False)
        positive_details.to_excel(writer, sheet_name='Детальная запись', index=False)
        
        # Автонастройка ширины колонок
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"✅ Сохранено: {positive_file}")
    print(f"   - Лист 1 (Отчетность): {len(positive_report)} строк")
    print(f"   - Лист 2 (Детальная запись): {len(positive_details):,} строк")
except Exception as e:
    print(f"❌ ОШИБКА при сохранении: {e}")

# ============================================================================
# ОТЧЕТ 2: ОТРИЦАТЕЛЬНЫЕ ПО КОММЕНТАРИЯМ
# ============================================================================
print("\n" + "=" * 80)
print("📊 ОТЧЕТ 2: ОТРИЦАТЕЛЬНЫЕ ПО КОММЕНТАРИЯМ")
print("=" * 80)

negative_df = df_with_date[df_with_date['Категория'] == 'Отрицательный'].copy()
print(f"Отрицательных записей: {len(negative_df):,}")

# Заполняем пустые комментарии
negative_df['Комментарий'] = negative_df['Комментарий'].fillna('Без комментария')
negative_df.loc[negative_df['Комментарий'].str.strip() == '', 'Комментарий'] = 'Без комментария'

# Сводка по комментариям
negative_report = negative_df.groupby(['Год', 'Месяц_Название', 'Служба', 'Статус', 'Комментарий']).agg({
    'Номер карты': 'count',
    'Оператор': 'nunique',
    'Оператор фиксировавший': 'nunique'
}).reset_index()

negative_report.columns = ['Год', 'Месяц', 'Служба', 'Статус', 'Комментарий', 'Количество заявок', 'Операторов', 'Фиксировавших']

# Сортируем
negative_report['Порядок'] = negative_report['Месяц'].map(month_order)
negative_report = negative_report.sort_values(['Год', 'Порядок', 'Служба', 'Количество заявок'], ascending=[True, True, True, False])
negative_report = negative_report.drop('Порядок', axis=1)

# Детальные записи
negative_details = negative_df[[
    'Год', 'Месяц_Название', 'Оператор', 'Архивный лист',
    'Номер карты', 'Номер телефона', 'Дата открытия карты',
    'Статус', 'Служба', 'Комментарий',
    'Оператор фиксировавший', 'Дата фиксации'
]].copy()
negative_details.columns = [
    'Год', 'Месяц', 'Оператор', 'Архивный лист',
    'Номер карты', 'Номер телефона', 'Дата открытия',
    'Статус', 'Служба', 'Комментарий',
    'Фиксировавший оператор', 'Дата фиксации'
]
negative_details['Порядок'] = negative_details['Месяц'].map(month_order)
negative_details = negative_details.sort_values(['Год', 'Порядок', 'Служба', 'Комментарий', 'Дата открытия'])
negative_details = negative_details.drop('Порядок', axis=1)

# Сохраняем (разбиваем детали на листы по 1 млн строк)
print("\n💾 Сохранение отчета по отрицательным...")

MAX_ROWS = 1000000  # Excel лимит ~1,048,576, берем с запасом
negative_file = OUTPUT_DIR / f'ОТЧЕТ_ОТРИЦАТЕЛЬНЫЕ_{timestamp}.xlsx'

try:
    with pd.ExcelWriter(negative_file, engine='openpyxl') as writer:
        negative_report.to_excel(writer, sheet_name='Отчетность', index=False)
        
        # Разбиваем детали на части
        total_rows = len(negative_details)
        num_parts = (total_rows // MAX_ROWS) + 1
        
        if num_parts == 1:
            negative_details.to_excel(writer, sheet_name='Детальная запись', index=False)
            print(f"   - Детальная запись: {total_rows:,} строк")
        else:
            for i in range(num_parts):
                start_idx = i * MAX_ROWS
                end_idx = min((i + 1) * MAX_ROWS, total_rows)
                part_df = negative_details.iloc[start_idx:end_idx]
                sheet_name = f'Детали - часть {i + 1}'
                part_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"   - {sheet_name}: {len(part_df):,} строк")
        
        # Автонастройка ширины колонок для всех листов
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"✅ Сохранено: {negative_file}")
    print(f"   - Лист 1 (Отчетность): {len(negative_report)} строк")
    print(f"   - Детальных записей всего: {len(negative_details):,} строк")
except Exception as e:
    print(f"❌ ОШИБКА при сохранении: {e}")

# ============================================================================
# СТАТИСТИКА
# ============================================================================
print("\n" + "=" * 80)
print("📊 ОБЩАЯ СТАТИСТИКА")
print("=" * 80)

print(f"\n📈 Положительные:")
print(f"   Всего: {len(positive_df):,}")
print(f"   Уникальных статусов: {positive_df['Статус'].nunique()}")
print(f"   Топ-5 статусов:")
for i, (status, count) in enumerate(positive_df['Статус'].value_counts().head(5).items(), 1):
    print(f"   {i}. {status}: {count:,}")

print(f"\n📉 Отрицательные:")
print(f"   Всего: {len(negative_df):,}")
print(f"   Уникальных комментариев: {negative_df['Комментарий'].nunique()}")
print(f"   Топ-10 комментариев:")
for i, (comment, count) in enumerate(negative_df['Комментарий'].value_counts().head(10).items(), 1):
    comment_short = comment[:60] + '...' if len(comment) > 60 else comment
    print(f"   {i}. {comment_short}: {count:,}")

print("\n" + "=" * 80)
print("✅ ОБА ОТЧЕТА ГОТОВЫ!")
print("=" * 80)
print("\nСозданные файлы:")
print("  1. ОТЧЕТ_ПОЛОЖИТЕЛЬНЫЕ.xlsx")
print("     - Лист 1: Сводка по статусам")
print("     - Лист 2: Все положительные записи")
print("\n  2. ОТЧЕТ_ОТРИЦАТЕЛЬНЫЕ.xlsx")
print("     - Лист 1: Сводка по комментариям")
print("     - Лист 2: Все отрицательные записи")
print("=" * 80)
