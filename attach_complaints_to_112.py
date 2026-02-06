#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Подставляет жалобы из Google Sheets в файлы 112 (папка 123).
Сохраняет новые файлы с добавленными колонками.
"""

import os
import re
import pandas as pd
from pathlib import Path
from datetime import datetime


def normalize_phone(phone):
    if pd.isna(phone):
        return ''
    phone_str = str(phone).replace('.0', '')
    phone_clean = ''.join(filter(str.isdigit, phone_str))
    if phone_clean.startswith('998') and len(phone_clean) == 12:
        phone_clean = phone_clean[3:]
    return phone_clean


def rename_statuses(status):
    if pd.isna(status):
        return status
    status_str = str(status).strip()
    statuses_to_rename = [
        'НЕТ ОТВЕТА (ЗАНЯТО)',
        'Заявка закрыта (не удалось дозвониться)',
        'Тиббиёт ходими аризаси',
        'Открыть карту',
        'Не удалось дозвониться'
    ]
    if status_str in statuses_to_rename:
        return 'Не удалось дозвониться'
    return status_str


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
    # Ищем последний файл
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    if not local_files:
        raise FileNotFoundError("Нет файлов КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv в папке data")

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

        # Опциональный фильтр по датам
        if str(os.environ.get('APPLY_DATE_FILTER', '')).strip() == '1':
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


def aggregate_unique(series):
    values = []
    for v in series:
        if v is None:
            continue
        s = str(v).strip()
        if s in ('', 'nan', 'None'):
            continue
        if s not in values:
            values.append(s)
    return '; '.join(values)


def first_nonempty(series):
    for v in series:
        if v is None:
            continue
        s = str(v).strip()
        if s not in ('', 'nan', 'None'):
            return s
    return ''


def main():
    by_service, by_incident = load_sheets_maps()

    input_files = sorted(Path('123').glob('*.xlsx'))
    if not input_files:
        print("❌ Нет файлов в папке 123")
        return

    output_dir = Path('output') / '123_с_жалобами'
    output_dir.mkdir(parents=True, exist_ok=True)

    start_from = int(os.environ.get("START_FROM", "1"))
    for idx, file_path in enumerate(input_files, 1):
        if idx < start_from:
            continue
        df_orig = pd.read_excel(file_path)
        df_112 = df_orig.rename(columns={
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

        df_112['row_id'] = df_112.index
        df_112['Инцидент_112_norm'] = df_112['Инцидент_112'].astype(str).str.strip()
        df_112['Служба_112'] = df_112['Служба_112'].astype(str)

        agg_rows = []
        for _, row in df_112[['row_id', 'Инцидент_112_norm', 'Служба_112']].iterrows():
            incident = str(row['Инцидент_112_norm']).strip()
            service = str(row['Служба_112']).strip()

            entry = by_service.get((incident, service))
            if entry is None:
                entry = by_incident.get(incident)

            complaint = ''
            status = ''
            positive = ''
            if entry:
                complaint = '; '.join(sorted(entry['complaints'])) if entry['complaints'] else ''
                status = entry['status']
                positive = entry['positive']

            agg_rows.append({
                'row_id': row['row_id'],
                'Жалоба': complaint,
                'Статус_связи': status,
                'Положительно': positive
            })

        agg = pd.DataFrame(agg_rows)

        df_out = df_orig.copy()
        df_out['Жалоба'] = ''
        df_out['Статус_связи'] = ''
        df_out['Положительно'] = ''

        df_out = df_out.join(agg.set_index('row_id'), how='left', rsuffix='_new')
        df_out['Жалоба'] = df_out['Жалоба_new'].fillna('')
        df_out['Статус_связи'] = df_out['Статус_связи_new'].fillna('')
        df_out['Положительно'] = df_out['Положительно_new'].fillna('')

        # Если жалобы нет, но совпадение было — ставим "Положительно"
        has_match = df_out['Статус_связи'].astype(str).str.strip().ne('') | df_out['Жалоба'].astype(str).str.strip().ne('')
        no_complaint = df_out['Жалоба'].astype(str).str.strip().eq('')
        empty_positive = df_out['Положительно'].astype(str).str.strip().eq('')
        df_out.loc[has_match & no_complaint & empty_positive, 'Положительно'] = 'Положительно'
        df_out = df_out.drop(columns=['Жалоба_new', 'Статус_связи_new', 'Положительно_new'])

        out_name = file_path.stem + '_с_жалобами.xlsx'
        out_path = output_dir / out_name
        if out_path.exists():
            print(f"↷ {out_name} уже существует, пропускаю")
            continue
        df_out.to_excel(out_path, index=False)
        print(f"✓ {out_name}")

    print(f"Готово. Файлы: {output_dir}")


if __name__ == '__main__':
    main()
