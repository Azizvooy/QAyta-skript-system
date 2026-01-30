"""
Создание отчета в формате QAYTA ALOQA REPORT - КАК В PDF
Структура: Год | Месяц | Регион | Службы (101, 102, 103, 104)
Для каждой службы: Jami (Всего) | Qanoatlantirildi (Положительный) | Qanoatlantirilmadi (Отрицательный)
"""
import pandas as pd
from datetime import datetime
from pathlib import Path
import warnings
import sys
warnings.filterwarnings('ignore')

# Базовые пути
BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / 'data'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'

print("=" * 100)
print("СОЗДАНИЕ ОТЧЕТА QAYTA ALOQA REPORT - ФОРМАТ КАК В PDF")
print("=" * 100)

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
    print(f"❌ ОШИБКА при загрузке: {e}")
    sys.exit(1)

# Функция для извлечения года и месяца из даты
def extract_year_month(date_str):
    if pd.isna(date_str) or not str(date_str).strip():
        return None, None
    
    try:
        # Пробуем разные форматы
        date_str = str(date_str).strip()
        
        # DD.MM.YYYY HH:MM:SS или DD.MM.YYYY
        if '.' in date_str:
            parts = date_str.split('.')[0:3]
            if len(parts) >= 3:
                day = parts[0].strip()
                month = parts[1].strip()
                year = parts[2].split()[0].strip()
                
                # Валидация и добавление ведущего нуля к месяцу
                if day.isdigit() and month.isdigit() and year.isdigit():
                    if len(year) == 4 and 1 <= int(month) <= 12:
                        return year, month.zfill(2)
        
        # DD/MM/YYYY
        if '/' in date_str:
            parts = date_str.split('/')
            if len(parts) >= 3:
                day = parts[0].strip()
                month = parts[1].strip()
                year = parts[2].split()[0].strip()
                
                if day.isdigit() and month.isdigit() and year.isdigit():
                    if len(year) == 4 and 1 <= int(month) <= 12:
                        return year, month.zfill(2)
        
        # YYYY-MM-DD
        if '-' in date_str:
            parts = date_str.split('-')
            if len(parts) >= 3:
                year = parts[0].strip()
                month = parts[1].strip()
                
                if year.isdigit() and month.isdigit():
                    if len(year) == 4 and 1 <= int(month) <= 12:
                        return year, month.zfill(2)
    except:
        pass
    
    return None, None

# Словари для преобразования
month_names = {
    '01': 'Yanvar', '02': 'Fevral', '03': 'Mart', '04': 'Aprel',
    '05': 'May', '06': 'Iyun', '07': 'Iyul', '08': 'Avgust',
    '09': 'Sentyabr', '10': 'Oktyabr', '11': 'Noyabr', '12': 'Dekabr'
}

# ВАЖНО: Нужен маппинг номера карты -> регион
# Пока создадим заглушку, потом добавим реальные регионы
regions = [
    'Toshkent shahri', 'Toshkent viloyati', 'Farg\'ona viloyati',
    'Andijon viloyati', 'Namangan viloyati', 'Sirdaryo viloyati',
    'Jizzax viloyati', 'Samarqand viloyati', 'Navoiy viloyati',
    'Buxoro viloyati', 'Qashqadaryo viloyati', 'Surxondaryo viloyati',
    'Xorazm viloyati', 'Qoraqalpog\'iston Respublikasi'
]

print("\n⚠️  ВНИМАНИЕ: Для полного отчета нужен файл с регионами!")
print("   Номер карты → Регион")
print("   Пока будет использовано: 'Регион не указан'\n")

# Добавляем колонки для анализа
print("🔄 Обработка данных...")
df['Год'], df['Месяц'] = zip(*df['Дата открытия карты'].apply(extract_year_month))

# Фильтруем только строки с датой
df_filtered = df[df['Год'].notna()].copy()

# Фильтруем только нормальные года (2024-2025)
df_filtered = df_filtered[df_filtered['Год'].isin(['2024', '2025'])].copy()

print(f"✅ Строк после фильтрации: {len(df_filtered):,}")

# Заглушка для региона (пока нет данных)
df_filtered['Регион'] = 'Регион не указан'

# Категоризация статусов
def categorize_status(status):
    if pd.isna(status):
        return 'Прочее'
    status = str(status).strip().lower()
    
    positive_keywords = ['положительн', 'положит', 'qanoatlantir', 'қаноатлантир']
    negative_keywords = ['отрицательн', 'отриц', 'qanoatlantirilmadi', 'қаноатлантирилмади']
    
    for kw in positive_keywords:
        if kw in status:
            return 'Положительный'
    
    for kw in negative_keywords:
        if kw in status:
            return 'Отрицательный'
    
    return 'Прочее'

df_filtered['Категория'] = df_filtered['Статус'].apply(categorize_status)

# Создаем сводный отчет
print("\n📊 Создание отчета...")

services = ['101', '102', '103', '104']

# Группируем данные
result = []

for (year, month, region), group in df_filtered.groupby(['Год', 'Месяц', 'Регион']):
    row = {
        'Yil': year,
        'Oy': month_names.get(month, month),
        'Hudud': region
    }
    
    for service in services:
        service_data = group[group['Служба'] == service]
        
        total = len(service_data)
        positive = len(service_data[service_data['Категория'] == 'Положительный'])
        negative = len(service_data[service_data['Категория'] == 'Отрицательный'])
        
        row[f'{service}_Jami'] = total
        row[f'{service}_Qanoatlantirildi'] = positive
        row[f'{service}_Qanoatlantirilmadi'] = negative
    
    result.append(row)

# Создаем DataFrame
report_df = pd.DataFrame(result)

# Сортируем по году и месяцу
month_order = {'Yanvar': 1, 'Fevral': 2, 'Mart': 3, 'Aprel': 4, 'May': 5, 'Iyun': 6,
               'Iyul': 7, 'Avgust': 8, 'Sentyabr': 9, 'Oktyabr': 10, 'Noyabr': 11, 'Dekabr': 12}
report_df['Month_Order'] = report_df['Oy'].map(month_order)
report_df = report_df.sort_values(['Yil', 'Month_Order', 'Hudud'])
report_df = report_df.drop('Month_Order', axis=1)

# Сохраняем в Excel
print("\n💾 Сохранение отчета...")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = OUTPUT_DIR / f'QAYTA_ALOQA_REPORT_FINAL_{timestamp}.xlsx'

try:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Основной отчет
    report_df.to_excel(writer, sheet_name='Отчет по месяцам', index=False)
    
    # Детальная статистика по месяцам и службам
    month_names_ru = {
        '01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель',
        '05': 'Май', '06': 'Июнь', '07': 'Июль', '08': 'Август',
        '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'
    }
    
    detail_stats = []
    for (year, month), group in df_filtered.groupby(['Год', 'Месяц']):
        for service in services:
            service_data = group[group['Служба'] == service]
            if len(service_data) > 0:
                detail_stats.append({
                    'Год': year,
                    'Месяц': month_names_ru.get(month, month),
                    'Служба': service,
                    'Всего заявок': len(service_data),
                    'Положительных': len(service_data[service_data['Категория'] == 'Положительный']),
                    'Отрицательных': len(service_data[service_data['Категория'] == 'Отрицательный']),
                    'Уникальных карт': service_data['Номер карты'].nunique(),
                    'Операторов': service_data['Оператор'].nunique()
                })
    
    detail_df = pd.DataFrame(detail_stats)
    detail_df.to_excel(writer, sheet_name='Детальная статистика', index=False)
    
    # Статистика по операторам
    operator_stats = []
    for (year, month, service, operator), group in df_filtered.groupby(['Год', 'Месяц', 'Служба', 'Оператор']):
        if len(group) > 0:
            operator_stats.append({
                'Год': year,
                'Месяц': month_names_ru.get(month, month),
                'Служба': service,
                'Оператор': operator,
                'Всего заявок': len(group),
                'Положительных': len(group[group['Категория'] == 'Положительный']),
                'Отрицательных': len(group[group['Категория'] == 'Отрицательный']),
                'Уникальных карт': group['Номер карты'].nunique()
            })
    
    operator_df = pd.DataFrame(operator_stats)
    if len(operator_df) > 0:
        operator_df = operator_df.sort_values(['Год', 'Месяц', 'Служба', 'Всего заявок'], ascending=[True, True, True, False])
    operator_df.to_excel(writer, sheet_name='По операторам', index=False)

    print(f"✅ Отчет сохранен: {output_file}")
except Exception as e:
    print(f"❌ ОШИБКА при сохранении: {e}")
    sys.exit(1)

# Выводим предварительный просмотр
print("\n" + "=" * 100)
print("📋 ПРЕДВАРИТЕЛЬНЫЙ ПРОСМОТР ОТЧЕТА")
print("=" * 100)

# Показываем первые 15 строк
print(report_df.head(15).to_string(index=False))

# Итоговая статистика
print("\n" + "=" * 100)
print("📊 ИТОГОВАЯ СТАТИСТИКА")
print("=" * 100)

print(f"\nВсего периодов в отчете: {len(report_df)}")
print(f"Охваченные года: {', '.join(sorted(report_df['Yil'].unique()))}")
print(f"Охваченные месяцы: {', '.join(report_df['Oy'].unique())}")

print("\n📊 Итоги по службам (за весь период):")
for service in services:
    total = report_df[f'{service}_Jami'].sum()
    positive = report_df[f'{service}_Qanoatlantirildi'].sum()
    negative = report_df[f'{service}_Qanoatlantirilmadi'].sum()
    
    if total > 0:
        pos_pct = (positive / total * 100)
        neg_pct = (negative / total * 100)
        print(f"\n  Служба {service}:")
        print(f"    Jami (Всего): {total:,}")
        print(f"    Qanoatlantirildi (Положительных): {positive:,} ({pos_pct:.1f}%)")
        print(f"    Qanoatlantirilmadi (Отрицательных): {negative:,} ({neg_pct:.1f}%)")

# Статистика по месяцам
month_names_ru = {
    '01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель',
    '05': 'Май', '06': 'Июнь', '07': 'Июль', '08': 'Август',
    '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'
}

print("\n📅 Статистика по месяцам:")
for (year, month), group in df_filtered.groupby(['Год', 'Месяц']):
    month_name = month_names_ru.get(month, month)
    total = len(group)
    positive = len(group[group['Категория'] == 'Положительный'])
    negative = len(group[group['Категория'] == 'Отрицательный'])
    
    print(f"\n  {month_name} {year}:")
    print(f"    Всего: {total:,}")
    print(f"    Положительных: {positive:,} ({positive/total*100:.1f}%)")
    print(f"    Отрицательных: {negative:,} ({negative/total*100:.1f}%)")

print("\n" + "=" * 100)
print("✅ ОТЧЕТ ГОТОВ!")
print("=" * 100)
print("\n⚠️  ПРИМЕЧАНИЕ:")
print("   Регион указан как 'Регион не указан' - нужен файл с маппингом:")
print("   Номер карты → Регион → Район")
print("=" * 100)
