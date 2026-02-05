#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Создает отчеты, где базой являются все записи 112 (4 файла),
а данные из Google Sheets подставляются как жалобы/статусы.
"""

import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from process_period_data import build_summary_tables, normalize_phone, rename_statuses


def extract_services(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    text = str(value)
    found = re.findall(r"\b(101|102|103|104)\b", text)
    if found:
        return list(dict.fromkeys(found))
    parts = re.split(r"[;,/\\|\s]+", text)
    parts = [p.strip() for p in parts if p.strip()]
    return list(dict.fromkeys(parts))


def update_map(store, key, complaint, status, positive):
    if key not in store:
        store[key] = {'complaints': set(), 'status': '', 'positive': ''}
    entry = store[key]
    if complaint:
        entry['complaints'].add(complaint)
    if not entry['status'] and status:
        entry['status'] = status
    if not entry['positive'] and positive:
        entry['positive'] = positive


def load_sheets_maps():
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    if not local_files:
        raise FileNotFoundError("Нет файла КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv в data")

    latest_file = max(local_files, key=lambda p: p.stat().st_ctime)
    usecols = ['Колонка_2', 'Колонка_4', 'Колонка_5', 'Колонка_6', 'Колонка_7', 'Колонка_8']

    by_service = {}
    by_incident = {}

    for chunk in pd.read_csv(latest_file, usecols=usecols, chunksize=200000, low_memory=False):
        chunk = chunk.rename(columns={
            'Колонка_2': 'Инцидент_Sheets',
            'Колонка_4': 'Дата_открытия',
            'Колонка_5': 'Статус_связи',
            'Колонка_6': 'Служба_Sheets',
            'Колонка_7': 'Жалоба',
            'Колонка_8': 'Положительно'
        })

        # Фильтр по датам 04–31.01.2026
        date_series = pd.to_datetime(chunk['Дата_открытия'], errors='coerce', dayfirst=True)
        start_date = pd.Timestamp('2026-01-04')
        end_date = pd.Timestamp('2026-01-31 23:59:59')
        chunk = chunk[(date_series >= start_date) & (date_series <= end_date)].copy()
        if chunk.empty:
            continue

        chunk['Инцидент_Sheets_norm'] = chunk['Инцидент_Sheets'].astype(str).str.strip()
        chunk['Статус_связи'] = chunk['Статус_связи'].apply(rename_statuses)

        for _, row in chunk.iterrows():
            incident = str(row['Инцидент_Sheets_norm']).strip()
            if incident in ('', 'nan', 'None'):
                continue
            complaint = str(row.get('Жалоба', '')).strip()
            if complaint in ('nan', 'None'):
                complaint = ''
            status = str(row.get('Статус_связи', '')).strip()
            if status in ('nan', 'None'):
                status = ''
            positive = str(row.get('Положительно', '')).strip()
            if positive in ('nan', 'None'):
                positive = ''

            services = extract_services(row.get('Служба_Sheets'))
            if services:
                for service in services:
                    update_map(by_service, (incident, str(service)), complaint, status, positive)
            else:
                update_map(by_incident, incident, complaint, status, positive)

    return by_service, by_incident


def load_112_data():
    files = sorted(Path('123').glob('*.xlsx'))
    if not files:
        raise FileNotFoundError("Нет файлов в папке 123")

    all_data = []
    for file in files:
        df = pd.read_excel(file)
        all_data.append(df)

    df_112 = pd.concat(all_data, ignore_index=True)

    df_112 = df_112.rename(columns={
        'Карточка рақами': 'Карта_112',
        'Ҳодиса рақами': 'Инцидент_112',
        'Хизмат': 'Служба_112',
        'Мурожаатчи телефон рақами': 'Телефон_112',
        'Ҳолат': 'Статус_112',
        'Вилоят': 'Регион_112',
        'Туман': 'Район_112',
        'Оператор': 'Оператор_112',
        'Сана': 'Дата_112'
    })

    df_112['Телефон_нормализованный'] = df_112['Телефон_112'].apply(normalize_phone)
    df_112['Статус_112'] = df_112['Статус_112'].apply(rename_statuses)
    df_112['Инцидент_112_norm'] = df_112['Инцидент_112'].astype(str).str.strip()
    df_112['Служба_112'] = df_112['Служба_112'].astype(str)

    return df_112


def main():
    by_service, by_incident = load_sheets_maps()
    df_112 = load_112_data()

    complaints = []
    statuses = []
    positives = []

    for incident, service in zip(df_112['Инцидент_112_norm'], df_112['Служба_112']):
        entry = by_service.get((incident, service)) or by_incident.get(incident)
        if entry:
            complaint = '; '.join(sorted(entry['complaints'])) if entry['complaints'] else ''
            status = entry['status']
            positive = entry['positive']
        else:
            complaint = ''
            status = ''
            positive = ''
        complaints.append(complaint)
        statuses.append(status)
        positives.append(positive)

    df_112['Жалоба'] = complaints
    df_112['Статус_связи'] = statuses
    df_112['Положительно'] = positives
    df_112['Есть_жалоба'] = df_112['Жалоба'].astype(str).str.strip() != ''

    # Сводки
    complaints_region, complaints_region_type, negative_df = build_summary_tables(df_112)

    # Сохранение
    reports_dir = Path('reports') / '2026-01_full_112'
    reports_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

    out_xlsx = reports_dir / f'ОТЧЁТ_2026-01_FULL_112_{timestamp}.xlsx'
    with pd.ExcelWriter(out_xlsx, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
        complaints_region.to_excel(writer, index=False, sheet_name='Жалобы_по_регионам')
        complaints_region_type.to_excel(writer, index=False, sheet_name='Регионы_и_жалобы')
        df_112.to_excel(writer, index=False, sheet_name='Детальные')
        negative_df.to_excel(writer, index=False, sheet_name='Отрицательные_и_жалобы')

    print(f"Готово: {out_xlsx}")


if __name__ == '__main__':
    main()
