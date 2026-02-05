#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Заполняет неполные отчеты (output/reports) данными из Google Sheets.
"""

import re
import pandas as pd
from pathlib import Path

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


def normalize_incident(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    return str(value).strip()


def is_empty(value):
    if value is None:
        return True
    s = str(value).strip()
    return s in ('', 'nan', 'None')


def update_entry(entry, phone, date_open, status, service_raw, complaint, positive):
    if not entry['Телефон_Sheets'] and phone:
        entry['Телефон_Sheets'] = phone
    if not entry['Дата_открытия'] and date_open:
        entry['Дата_открытия'] = date_open
    if not entry['Статус_связи'] and status:
        entry['Статус_связи'] = status
    if not entry['Служба_Sheets'] and service_raw:
        entry['Служба_Sheets'] = service_raw
    if positive and not entry['Положительно']:
        entry['Положительно'] = positive
    if complaint:
        entry['Жалоба'].add(complaint)


def main():
    reports_dir = Path('output/reports')
    report_files = sorted(reports_dir.glob('ОТЧЁТ_2026-01_СЛУЖБА_*.xlsx'))
    if not report_files:
        print('❌ Нет файлов отчетов в output/reports')
        return

    # Собираем ключи из отчетов
    target_pairs = set()
    target_incidents = set()
    for f in report_files:
        df = pd.read_excel(f, sheet_name='Детальные')
        for col in [
            'Телефон_Sheets', 'Дата_открытия', 'Статус_связи', 'Служба_Sheets',
            'Жалоба', 'Положительно', 'Есть_жалоба',
            'Инцидент_Sheets', 'Инцидент_Sheets_norm'
        ]:
            if col in df.columns:
                df[col] = df[col].astype('object')
        if 'Инцидент_112_norm' in df.columns:
            inc_col = 'Инцидент_112_norm'
        else:
            inc_col = 'Инцидент_Sheets_norm'
        for inc, serv in zip(df[inc_col], df['Служба_112']):
            inc_norm = normalize_incident(inc)
            serv_norm = normalize_incident(serv)
            if inc_norm:
                target_pairs.add((inc_norm, serv_norm))
                target_incidents.add(inc_norm)

    # Читаем последний CSV
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    if not local_files:
        print('❌ Нет файла КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv в data')
        return
    latest_file = max(local_files, key=lambda p: p.stat().st_ctime)

    by_service = {}
    by_incident = {}

    usecols = ['Колонка_2', 'Колонка_3', 'Колонка_4', 'Колонка_5', 'Колонка_6', 'Колонка_7', 'Колонка_8']
    for chunk in pd.read_csv(latest_file, usecols=usecols, chunksize=200000, low_memory=False):
        chunk = chunk.rename(columns={
            'Колонка_2': 'Инцидент_Sheets',
            'Колонка_3': 'Телефон_Sheets',
            'Колонка_4': 'Дата_открытия',
            'Колонка_5': 'Статус_связи',
            'Колонка_6': 'Служба_Sheets',
            'Колонка_7': 'Жалоба',
            'Колонка_8': 'Положительно'
        })

        for _, row in chunk.iterrows():
            incident = normalize_incident(row.get('Инцидент_Sheets'))
            if not incident:
                continue
            if incident not in target_incidents:
                continue

            phone = '' if is_empty(row.get('Телефон_Sheets')) else str(row.get('Телефон_Sheets')).strip()
            date_open = '' if is_empty(row.get('Дата_открытия')) else str(row.get('Дата_открытия')).strip()
            status = '' if is_empty(row.get('Статус_связи')) else str(row.get('Статус_связи')).strip()
            service_raw = '' if is_empty(row.get('Служба_Sheets')) else str(row.get('Служба_Sheets')).strip()
            complaint = '' if is_empty(row.get('Жалоба')) else str(row.get('Жалоба')).strip()
            positive = '' if is_empty(row.get('Положительно')) else str(row.get('Положительно')).strip()

            services = extract_services(row.get('Служба_Sheets'))
            if services:
                for service in services:
                    key = (incident, str(service))
                    if key not in target_pairs:
                        continue
                    if key not in by_service:
                        by_service[key] = {
                            'Телефон_Sheets': '',
                            'Дата_открытия': '',
                            'Статус_связи': '',
                            'Служба_Sheets': '',
                            'Жалоба': set(),
                            'Положительно': ''
                        }
                    update_entry(by_service[key], phone, date_open, status, service_raw, complaint, positive)
            else:
                if incident not in by_incident:
                    by_incident[incident] = {
                        'Телефон_Sheets': '',
                        'Дата_открытия': '',
                        'Статус_связи': '',
                        'Служба_Sheets': '',
                        'Жалоба': set(),
                        'Положительно': ''
                    }
                update_entry(by_incident[incident], phone, date_open, status, service_raw, complaint, positive)

    # Заполняем отчеты
    out_dir = reports_dir / 'filled'
    out_dir.mkdir(parents=True, exist_ok=True)

    for f in report_files:
        df = pd.read_excel(f, sheet_name='Детальные')
        if 'Инцидент_112_norm' in df.columns:
            inc_col = 'Инцидент_112_norm'
        else:
            inc_col = 'Инцидент_Sheets_norm'

        for idx, row in df.iterrows():
            incident = normalize_incident(row.get(inc_col))
            service = normalize_incident(row.get('Служба_112'))
            entry = by_service.get((incident, service)) or by_incident.get(incident)
            if not entry:
                continue

            if 'Телефон_Sheets' in df.columns and is_empty(df.at[idx, 'Телефон_Sheets']):
                df.at[idx, 'Телефон_Sheets'] = entry['Телефон_Sheets']
            if 'Дата_открытия' in df.columns and is_empty(df.at[idx, 'Дата_открытия']):
                df.at[idx, 'Дата_открытия'] = entry['Дата_открытия']
            if 'Статус_связи' in df.columns and is_empty(df.at[idx, 'Статус_связи']):
                df.at[idx, 'Статус_связи'] = entry['Статус_связи']
            if 'Служба_Sheets' in df.columns and is_empty(df.at[idx, 'Служба_Sheets']):
                df.at[idx, 'Служба_Sheets'] = entry['Служба_Sheets']
            if 'Жалоба' in df.columns and is_empty(df.at[idx, 'Жалоба']):
                df.at[idx, 'Жалоба'] = '; '.join(sorted(entry['Жалоба'])) if entry['Жалоба'] else ''
            if 'Положительно' in df.columns:
                if is_empty(df.at[idx, 'Положительно']) and entry['Положительно']:
                    df.at[idx, 'Положительно'] = entry['Положительно']
            if 'Есть_жалоба' in df.columns:
                df.at[idx, 'Есть_жалоба'] = bool(str(df.at[idx, 'Жалоба']).strip())
            if 'Инцидент_Sheets' in df.columns and is_empty(df.at[idx, 'Инцидент_Sheets']):
                df.at[idx, 'Инцидент_Sheets'] = incident
            if 'Инцидент_Sheets_norm' in df.columns and is_empty(df.at[idx, 'Инцидент_Sheets_norm']):
                df.at[idx, 'Инцидент_Sheets_norm'] = incident

        out_path = out_dir / f.name
        with pd.ExcelWriter(out_path, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            # пересобираем служебные листы из текущих данных
            from process_period_data import build_summary_tables
            complaints_region, complaints_region_type, negative_df = build_summary_tables(df)
            complaints_region.to_excel(writer, index=False, sheet_name='Жалобы_по_регионам')
            complaints_region_type.to_excel(writer, index=False, sheet_name='Регионы_и_жалобы')
            df.to_excel(writer, index=False, sheet_name='Детальные')
            negative_df.to_excel(writer, index=False, sheet_name='Отрицательные_и_жалобы')

        print(f"✓ {out_path.name}")

    print(f"Готово. Заполненные отчеты: {out_dir}")


if __name__ == '__main__':
    main()
