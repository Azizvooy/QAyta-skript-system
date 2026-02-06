#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Создает 4 отчета по службам (101-104) с сопоставлением по номеру инцидента.
Правило:
- Если в Sheets есть жалоба и указана служба -> жалоба только этой службе.
- Если в Sheets есть жалоба без службы -> жалоба всем службам по инциденту.
- Статус/положительно применяются по инциденту ко всем службам.
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


def is_valid_complaint(value):
    """Проверка что жалоба не является номером инцидента"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    text = str(value).strip()
    if text in ('', 'nan', 'None'):
        return False
    # Фильтруем номера инцидентов (формат XX.XXXXXX/XX)
    incident_pattern = r'^\d{2}\.[A-Z0-9]{6,8}\/\d{2}$'
    if re.match(incident_pattern, text):
        return False
    return True


def extract_service_from_complaint(complaint_text):
    """Извлекает номер службы из начала жалобы (1. -> 101, 2. -> 102, и т.д.) и очищает текст"""
    if not complaint_text or pd.isna(complaint_text):
        return None, ''
    
    text = str(complaint_text).strip()
    # Паттерн: цифра + точка + пробелы в начале
    match = re.match(r'^(\d)\.\s*(.+)$', text)
    if match:
        service_digit = match.group(1)
        clean_text = match.group(2).strip()
        
        # Преобразуем цифру в номер службы
        service_map = {'1': '101', '2': '102', '3': '103', '4': '104'}
        service = service_map.get(service_digit)
        
        return service, clean_text
    
    return None, text


def parse_sheets_datetime(series):
    series_str = series.astype(str).str.strip()
    parsed = pd.to_datetime(series_str, format='%d.%m.%Y %H:%M:%S', errors='coerce')
    parsed = parsed.fillna(pd.to_datetime(series_str, format='%d.%m.%Y %H:%M', errors='coerce'))
    numeric = pd.to_numeric(series, errors='coerce')
    numeric = numeric.where((numeric >= 1) & (numeric <= 60000))
    excel_dates = pd.to_datetime(numeric, unit='D', origin='1899-12-30', errors='coerce')
    return parsed.fillna(excel_dates)


def load_sheets_maps():
    local_files = list(Path('data').glob('КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv'))
    if not local_files:
        raise FileNotFoundError("Нет файла КОНСОЛИДИРОВАННЫЕ_ДАННЫЕ_*.csv в data")

    latest_file = max(local_files, key=lambda p: p.stat().st_ctime)
    usecols = ['Колонка_2', 'Колонка_4', 'Колонка_5', 'Колонка_6', 'Колонка_7', 'Колонка_8']

    incident_status = {}
    incident_positive = {}
    incident_all_complaints = {}
    incident_service_complaints = {}

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
        date_series = parse_sheets_datetime(chunk['Дата_открытия'])
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
            # Фильтруем ошибочные номера инцидентов в жалобах
            if complaint and not is_valid_complaint(complaint):
                complaint = ''
            
            # Извлекаем службу из текста жалобы (1. -> 101, 2. -> 102 и т.д.)
            complaint_service, clean_complaint = extract_service_from_complaint(complaint)
            if clean_complaint:
                complaint = clean_complaint
            
            status = str(row.get('Статус_связи', '')).strip()
            if status in ('nan', 'None'):
                status = ''
            positive = str(row.get('Положительно', '')).strip()
            if positive in ('nan', 'None'):
                positive = ''

            # статусы по инциденту для всех служб
            if incident not in incident_status and status:
                incident_status[incident] = status
            if incident not in incident_positive and positive:
                incident_positive[incident] = positive

            # Определяем службу для жалобы: приоритет у службы из текста жалобы
            services = extract_services(row.get('Служба_Sheets'))
            if complaint_service:
                # Если служба указана в тексте жалобы - используем её
                key = (incident, complaint_service)
                if key not in incident_service_complaints:
                    incident_service_complaints[key] = set()
                if complaint:
                    incident_service_complaints[key].add(complaint)
            elif services:
                # Иначе используем службу из колонки Служба_Sheets
                for service in services:
                    key = (incident, str(service))
                    if key not in incident_service_complaints:
                        incident_service_complaints[key] = set()
                    if complaint:
                        incident_service_complaints[key].add(complaint)
            else:
                # Если служба вообще не указана - жалоба для всех служб
                if incident not in incident_all_complaints:
                    incident_all_complaints[incident] = set()
                if complaint:
                    incident_all_complaints[incident].add(complaint)

    return incident_status, incident_positive, incident_all_complaints, incident_service_complaints


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


def apply_sheets_data(df_112, incident_status, incident_positive, incident_all_complaints, incident_service_complaints):
    complaints = []
    statuses = []
    positives = []

    for incident, service in zip(df_112['Инцидент_112_norm'], df_112['Служба_112']):
        complaint_set = set()
        key = (incident, str(service))

        if key in incident_service_complaints:
            complaint_set |= incident_service_complaints[key]
        elif incident in incident_all_complaints:
            complaint_set |= incident_all_complaints[incident]

        complaint = '; '.join(sorted(complaint_set)) if complaint_set else ''
        status = incident_status.get(incident, '')
        positive = incident_positive.get(incident, '')

        complaints.append(complaint)
        statuses.append(status)
        positives.append(positive)

    df_112['Жалоба'] = complaints
    df_112['Статус_связи'] = statuses
    df_112['Положительно'] = positives
    df_112['Есть_жалоба'] = df_112['Жалоба'].astype(str).str.strip() != ''

    return df_112


def save_service_report(df_service, service_num, out_dir):
    complaints_region, complaints_region_type, negative_df = build_summary_tables(df_service)

    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    out_xlsx = out_dir / f'СЛУЖБА_{service_num}_ИНЦИДЕНТЫ_{timestamp}.xlsx'

    with pd.ExcelWriter(out_xlsx, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
        complaints_region.to_excel(writer, index=False, sheet_name='Жалобы_по_регионам')
        complaints_region_type.to_excel(writer, index=False, sheet_name='Регионы_и_жалобы')
        df_service.to_excel(writer, index=False, sheet_name='Детальные')
        negative_df.to_excel(writer, index=False, sheet_name='Отрицательные_и_жалобы')

    return out_xlsx


def main():
    print("=== Загрузка данных из Sheets ===")
    incident_status, incident_positive, incident_all_complaints, incident_service_complaints = load_sheets_maps()

    print("=== Загрузка данных 112 ===")
    df_112 = load_112_data()

    print("=== Применение данных Sheets по номеру инцидента ===")
    df_112 = apply_sheets_data(
        df_112,
        incident_status,
        incident_positive,
        incident_all_complaints,
        incident_service_complaints
    )

    out_dir = Path('reports') / '2026-01_services_by_incident'
    out_dir.mkdir(parents=True, exist_ok=True)

    for service_num in ['101', '102', '103', '104']:
        df_service = df_112[df_112['Служба_112'] == service_num].copy()
        if df_service.empty:
            print(f"⚠️ Нет данных для службы {service_num}")
            continue
        out_file = save_service_report(df_service, service_num, out_dir)
        print(f"✅ Сохранено: {out_file}")

    print(f"\nГотово. Папка: {out_dir}")


if __name__ == '__main__':
    main()
