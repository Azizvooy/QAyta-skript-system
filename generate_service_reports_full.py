#!/usr/bin/env python3
"""
Генерация отчетов по каждой службе отдельно на основе полного отчета
"""

import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import xlsxwriter

def load_full_report():
    """Загружает данные из полного отчета"""
    print("Загрузка данных из полного отчета...")
    
    # Ищем полный отчет
    report_dir = Path("reports/2026-01_full_112")
    if not report_dir.exists():
        print("⚠️ Не найдена папка с полным отчетом")
        return None
    
    xlsx_files = list(report_dir.glob("ОТЧЁТ_*.xlsx"))
    if not xlsx_files:
        print("⚠️ Не найден полный отчет")
        return None
    
    latest_file = sorted(xlsx_files)[-1]
    print(f"Используем файл: {latest_file.name}")
    
    # Читаем лист с детальными данными
    df = pd.read_excel(latest_file, sheet_name='Детальные')
    print(f"Загружено {len(df):,} записей")
    
    return df

def process_service(service_num, df_all):
    """Обрабатывает данные для одной службы"""
    print(f"\n{'='*60}")
    print(f"Обработка службы {service_num}")
    print(f"{'='*60}")
    
    # Фильтруем данные по службе
    df_service = df_all[df_all['Служба_112'] == service_num].copy()
    print(f"Записей для службы {service_num}: {len(df_service):,}")
    
    if len(df_service) == 0:
        print(f"⚠️ Нет данных для службы {service_num}")
        return None
    
    # Статистика
    if 'Дата_112' in df_service.columns:
        print(f"Период: {df_service['Дата_112'].min()} — {df_service['Дата_112'].max()}")
    
    complaints_count = df_service['Есть_жалоба'].sum() if 'Есть_жалоба' in df_service.columns else 0
    print(f"Записей с жалобами: {complaints_count:,}")
    
    return df_service

def create_summary_tables(df):
    """Создает сводные таблицы"""
    # 1. Жалобы по регионам
    complaints_df = df[df['Есть_жалоба'] == True].copy()
    if len(complaints_df) > 0:
        region_complaints = complaints_df.groupby('Регион_112').size().reset_index(name='Количество_жалоб')
        region_complaints = region_complaints.sort_values('Количество_жалоб', ascending=False)
    else:
        region_complaints = pd.DataFrame(columns=['Регион_112', 'Количество_жалоб'])
    
    # 2. Регионы и жалобы (все записи)
    region_summary = df.groupby('Регион_112').agg({
        'Инцидент_112': 'count',
        'Есть_жалоба': 'sum'
    }).reset_index()
    region_summary.columns = ['Регион', 'Всего_обращений', 'Количество_жалоб']
    region_summary = region_summary.sort_values('Всего_обращений', ascending=False)
    
    # 3. Отрицательные оценки и жалобы
    negative_df = df[df['Есть_жалоба'] == True].copy()
    
    return region_complaints, region_summary, negative_df

def save_service_report(df, service_num, output_dir):
    """Сохраняет отчет для службы в Excel"""
    
    # Создаем сводные таблицы
    region_complaints, region_summary, negative_df = create_summary_tables(df)
    
    # Формируем имя файла
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"СЛУЖБА_{service_num}_2026-01_ПОЛНЫЙ_{timestamp}.xlsx"
    filepath = output_dir / filename
    
    print(f"\nСоздание Excel файла: {filename}")
    
    # Создаем Excel файл
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Форматы
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Лист 1: Жалобы по регионам
        region_complaints.to_excel(writer, sheet_name='Жалобы_по_регионам', index=False)
        worksheet = writer.sheets['Жалобы_по_регионам']
        for col_num, value in enumerate(region_complaints.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 30)
        
        # Лист 2: Регионы и жалобы
        region_summary.to_excel(writer, sheet_name='Регионы_и_жалобы', index=False)
        worksheet = writer.sheets['Регионы_и_жалобы']
        for col_num, value in enumerate(region_summary.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 30)
        
        # Лист 3: Детальные данные
        df.to_excel(writer, sheet_name='Детальные', index=False)
        worksheet = writer.sheets['Детальные']
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Лист 4: Отрицательные и жалобы
        if len(negative_df) > 0:
            negative_df.to_excel(writer, sheet_name='Отрицательные_и_жалобы', index=False)
            worksheet = writer.sheets['Отрицательные_и_жалобы']
            for col_num, value in enumerate(negative_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
    
    print(f"✅ Сохранено: {filepath}")
    return filepath

def main():
    print("\n" + "="*60)
    print("ГЕНЕРАЦИЯ ОТЧЕТОВ ПО СЛУЖБАМ (ПОЛНЫЕ ДАННЫЕ)")
    print("="*60 + "\n")
    
    # Загружаем данные из полного отчета
    df_all = load_full_report()
    if df_all is None:
        return
    
    print(f"Всего записей: {len(df_all):,}")
    if 'Дата_112' in df_all.columns and df_all['Дата_112'].notna().any():
        print(f"Период: {df_all['Дата_112'].min()} — {df_all['Дата_112'].max()}")
    
    # Создаем выходную директорию
    output_dir = Path("reports/2026-01_services")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Обрабатываем каждую службу
    services = [101, 102, 103, 104]
    
    for service_num in services:
        df_service = process_service(service_num, df_all)
        if df_service is not None:
            save_service_report(df_service, service_num, output_dir)
    
    print("\n" + "="*60)
    print("✅ ГОТОВО! Все отчеты по службам созданы")
    print(f"📁 Папка: {output_dir}")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
