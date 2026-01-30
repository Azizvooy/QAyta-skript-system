#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
СОЗДАНИЕ ДЕТАЛЬНОЙ ОТЧЁТНОСТИ ПО СЛУЖБАМ
=============================================================================
Создаёт подробные отчёты по каждой службе (101, 102, 103, 104)
=============================================================================
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from collections import Counter

BASE_DIR = Path(__file__).parent

print('\n' + '='*80)
print('📊 СОЗДАНИЕ ДЕТАЛЬНОЙ ОТЧЁТНОСТИ ПО СЛУЖБАМ')
print('='*80)

# Загружаем сопоставленные данные
data_dir = BASE_DIR / 'data'
match_files = list(data_dir.glob('СОПОСТАВЛЕНИЕ_ПОЛНОЕ_*.csv'))

if not match_files:
    print('❌ Не найдены сопоставленные данные')
    exit(1)

match_file = max(match_files, key=lambda p: p.stat().st_mtime)
print(f'\nЗагрузка: {match_file.name}')

df = pd.read_csv(match_file, encoding='utf-8-sig', low_memory=False)
print(f'Всего записей: {len(df):,}\n')

# Создаём директорию для отчётов по службам
reports_dir = BASE_DIR / 'reports' / 'службы'
reports_dir.mkdir(exist_ok=True)

timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')

# Получаем список служб
services = df['Служба'].dropna().unique()
services = sorted([s for s in services if str(s) != 'nan'])

print(f'Найдено служб: {len(services)}')
for service in services:
    print(f'  • Служба {service}')

print('\n' + '='*80)
print('📋 СОЗДАНИЕ ОТЧЁТОВ ПО СЛУЖБАМ')
print('='*80 + '\n')

# Словарь названий служб
service_names = {
    '101.0': 'Пожарная служба',
    '102.0': 'Скорая медицинская помощь',
    '103.0': 'Газовая служба',
    '104.0': 'Аварийная служба'
}

# Создаём отчёт по каждой службе
for service in services:
    service_key = str(service)
    service_name = service_names.get(service_key, f'Служба {service_key}')
    
    print(f'[{service_key}] {service_name}...')
    
    # Фильтруем данные по службе
    df_service = df[df['Служба'] == service].copy()
    
    print(f'  Записей: {len(df_service):,}')
    
    # Сохраняем CSV
    csv_file = reports_dir / f'СЛУЖБА_{service_key.replace(".0", "")}_{timestamp}.csv'
    df_service.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f'  ✓ CSV: {csv_file.name}')
    
    # Создаём текстовый отчёт
    report_file = reports_dir / f'ОТЧЁТ_СЛУЖБА_{service_key.replace(".0", "")}_{timestamp}.txt'
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write('='*80 + '\n')
        f.write(f'ОТЧЁТ ПО СЛУЖБЕ {service_key}: {service_name.upper()}\n')
        f.write('='*80 + '\n')
        f.write(f'Дата: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
        f.write(f'Всего записей: {len(df_service):,}\n')
        f.write('='*80 + '\n\n')
        
        # Статистика по статусам
        f.write('📊 СТАТИСТИКА ПО СТАТУСАМ\n')
        f.write('-'*80 + '\n')
        if 'Статус' in df_service.columns:
            status_stats = df_service['Статус'].value_counts()
            total = len(df_service)
            for status, count in status_stats.items():
                pct = count / total * 100
                f.write(f'{str(status):<50} - {count:>8,} ({pct:>5.1f}%)\n')
        f.write('\n')
        
        # Статистика по документам/агентам
        f.write('='*80 + '\n')
        f.write('👥 СТАТИСТИКА ПО АГЕНТАМ\n')
        f.write('='*80 + '\n\n')
        if 'Документ' in df_service.columns:
            doc_stats = df_service['Документ'].value_counts()
            for idx, (doc, count) in enumerate(doc_stats.items(), 1):
                f.write(f'{idx:3}. {doc:<50} - {count:>8,} записей\n')
        f.write('\n')
        
        # Статистика по районам
        f.write('='*80 + '\n')
        f.write('🗺️  СТАТИСТИКА ПО РАЙОНАМ\n')
        f.write('='*80 + '\n\n')
        if 'Район' in df_service.columns:
            district_stats = df_service['Район'].value_counts().head(20)
            for idx, (district, count) in enumerate(district_stats.items(), 1):
                f.write(f'{idx:3}. {str(district):<50} - {count:>8,}\n')
        f.write('\n')
        
        # Статистика по операторам 112
        f.write('='*80 + '\n')
        f.write('🎧 СТАТИСТИКА ПО ОПЕРАТОРАМ СЛУЖБЫ 112\n')
        f.write('='*80 + '\n\n')
        if 'Оператор' in df_service.columns:
            operator_stats = df_service['Оператор'].value_counts().head(20)
            for idx, (operator, count) in enumerate(operator_stats.items(), 1):
                f.write(f'{idx:3}. {str(operator):<50} - {count:>8,}\n')
        f.write('\n')
        
        # Статистика по датам
        f.write('='*80 + '\n')
        f.write('📅 СТАТИСТИКА ПО ДАТАМ\n')
        f.write('='*80 + '\n\n')
        if 'Дата' in df_service.columns:
            date_stats = df_service['Дата'].value_counts().head(10)
            for idx, (date, count) in enumerate(date_stats.items(), 1):
                f.write(f'{idx:3}. {str(date):<20} - {count:>8,} обращений\n')
        f.write('\n')
        
        # Статистика по листам
        f.write('='*80 + '\n')
        f.write('📄 СТАТИСТИКА ПО ТИПАМ ЛИСТОВ\n')
        f.write('='*80 + '\n\n')
        if 'Лист' in df_service.columns:
            sheet_stats = df_service['Лист'].value_counts()
            for sheet, count in sheet_stats.items():
                pct = count / len(df_service) * 100
                f.write(f'{str(sheet):<60} - {count:>8,} ({pct:>5.1f}%)\n')
        f.write('\n')
        
        f.write('='*80 + '\n')
        f.write('✅ ОТЧЁТ ЗАВЕРШЁН\n')
        f.write('='*80 + '\n')
    
    print(f'  ✓ Отчёт: {report_file.name}')
    print()

# Создаём сводный отчёт по всем службам
print('='*80)
print('📊 СОЗДАНИЕ СВОДНОГО ОТЧЁТА ПО ВСЕМ СЛУЖБАМ')
print('='*80 + '\n')

summary_file = BASE_DIR / 'reports' / f'СВОДНЫЙ_ОТЧЁТ_ПО_СЛУЖБАМ_{timestamp}.txt'

with open(summary_file, 'w', encoding='utf-8') as f:
    f.write('='*80 + '\n')
    f.write('СВОДНЫЙ ОТЧЁТ ПО ВСЕМ СЛУЖБАМ\n')
    f.write('='*80 + '\n')
    f.write(f'Дата: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
    f.write(f'Период: январь 2026\n')
    f.write('='*80 + '\n\n')
    
    f.write('📊 ОБЩАЯ СТАТИСТИКА\n')
    f.write('-'*80 + '\n')
    f.write(f'Всего записей: {len(df):,}\n')
    f.write(f'Служб: {len(services)}\n\n')
    
    # Статистика по каждой службе
    f.write('='*80 + '\n')
    f.write('📋 ДЕТАЛЬНАЯ СТАТИСТИКА ПО СЛУЖБАМ\n')
    f.write('='*80 + '\n\n')
    
    total_records = len(df)
    
    for service in sorted(services):
        service_key = str(service)
        service_name = service_names.get(service_key, f'Служба {service_key}')
        df_service = df[df['Служба'] == service]
        count = len(df_service)
        pct = count / total_records * 100
        
        f.write(f'\n{"="*80}\n')
        f.write(f'СЛУЖБА {service_key}: {service_name.upper()}\n')
        f.write(f'{"="*80}\n\n')
        f.write(f'Всего обращений: {count:,} ({pct:.1f}% от общего числа)\n\n')
        
        # ТОП-5 агентов по этой службе
        if 'Документ' in df_service.columns:
            f.write('ТОП-5 АГЕНТОВ:\n')
            doc_stats = df_service['Документ'].value_counts().head(5)
            for idx, (doc, cnt) in enumerate(doc_stats.items(), 1):
                f.write(f'  {idx}. {doc:<40} - {cnt:>8,} записей\n')
            f.write('\n')
        
        # ТОП-5 районов
        if 'Район' in df_service.columns:
            f.write('ТОП-5 РАЙОНОВ:\n')
            district_stats = df_service['Район'].value_counts().head(5)
            for idx, (district, cnt) in enumerate(district_stats.items(), 1):
                f.write(f'  {idx}. {str(district):<40} - {cnt:>8,} обращений\n')
            f.write('\n')
        
        # Статусы
        if 'Статус' in df_service.columns:
            f.write('СТАТУСЫ:\n')
            status_stats = df_service['Статус'].value_counts()
            for status, cnt in status_stats.items():
                status_pct = cnt / count * 100
                f.write(f'  • {str(status):<40} - {cnt:>8,} ({status_pct:>5.1f}%)\n')
            f.write('\n')
    
    # Сравнительная таблица
    f.write('\n' + '='*80 + '\n')
    f.write('📊 СРАВНИТЕЛЬНАЯ ТАБЛИЦА СЛУЖБ\n')
    f.write('='*80 + '\n\n')
    
    f.write(f'{"Служба":<30} {"Обращений":>15} {"Процент":>10} {"Обработано":>10}\n')
    f.write('-'*80 + '\n')
    
    for service in sorted(services):
        service_key = str(service)
        service_name = service_names.get(service_key, f'Служба {service_key}')
        df_service = df[df['Служба'] == service]
        count = len(df_service)
        pct = count / total_records * 100
        
        # Процент обработанных
        if 'Статус' in df_service.columns:
            processed = len(df_service[df_service['Статус'] == 'Обработан'])
            processed_pct = processed / count * 100 if count > 0 else 0
        else:
            processed_pct = 0
        
        f.write(f'{service_name:<30} {count:>15,} {pct:>9.1f}% {processed_pct:>9.1f}%\n')
    
    f.write('-'*80 + '\n')
    f.write(f'{"ИТОГО":<30} {total_records:>15,} {100.0:>9.1f}%\n')
    
    f.write('\n' + '='*80 + '\n')
    f.write('✅ СВОДНЫЙ ОТЧЁТ ЗАВЕРШЁН\n')
    f.write('='*80 + '\n')

print(f'✓ Сводный отчёт: {summary_file.name}\n')

# Создаём Excel файл со всеми службами на разных листах
print('='*80)
print('📊 СОЗДАНИЕ EXCEL ОТЧЁТА')
print('='*80 + '\n')

excel_file = BASE_DIR / 'reports' / f'ОТЧЁТ_ВСЕ_СЛУЖБЫ_{timestamp}.xlsx'

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    # Лист со сводкой
    summary_data = []
    for service in sorted(services):
        service_key = str(service)
        service_name = service_names.get(service_key, f'Служба {service_key}')
        df_service = df[df['Служба'] == service]
        
        summary_data.append({
            'Служба': service_key,
            'Название': service_name,
            'Всего обращений': len(df_service),
            'Процент': f'{len(df_service)/len(df)*100:.1f}%'
        })
    
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Сводка', index=False)
    
    # Листы по каждой службе (ограничиваем первыми 50000 строк)
    for service in sorted(services):
        service_key = str(service).replace('.0', '')
        df_service = df[df['Служба'] == service].head(50000)
        
        # Отбираем ключевые колонки
        key_cols = ['Номер_карты_norm', 'Телефон_norm', 'Колонка_4', 'Колонка_5', 
                    'Документ', 'Лист', 'Номер инцидента', 'Дата', 'Статус', 
                    'Оператор', 'Район', 'Адрес']
        
        available_cols = [col for col in key_cols if col in df_service.columns]
        df_service_export = df_service[available_cols]
        
        sheet_name = f'Служба {service_key}'[:31]  # Ограничение Excel
        df_service_export.to_excel(writer, sheet_name=sheet_name, index=False)

print(f'✓ Excel отчёт: {excel_file.name}\n')

print('='*80)
print('✅ ВСЯ ОТЧЁТНОСТЬ СОЗДАНА')
print('='*80)
print(f'\n📁 Созданные файлы:\n')
print(f'1. Сводный отчёт: {summary_file.name}')
print(f'2. Excel отчёт: {excel_file.name}')
print(f'3. Отчёты по службам в папке: reports/службы/')
print(f'   • {len(services)} текстовых отчётов')
print(f'   • {len(services)} CSV файлов')
print('\n' + '='*80)
