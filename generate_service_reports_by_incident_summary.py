#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Создает 4 файла по службам с 2 листами:
1) Сводка_количество — только количества (всего, жалобы, отрицательные, отрицательные+жалобы)
2) Матрица_жалоб — строки: регионы, колонки: категории жалоб, значения: количество

Логика сопоставления по инциденту:
- Если в Sheets есть жалоба и указана служба -> жалоба только этой службе.
- Если в Sheets есть жалоба без службы -> жалоба всем службам по инциденту.
- Статус применяется по инциденту ко всем службам.
"""

import pandas as pd
import re
from pathlib import Path
from datetime import datetime
from generate_service_reports_by_incident import load_sheets_maps, load_112_data, apply_sheets_data


def build_summary_counts(df):
    total = len(df)
    complaints = int(df['Есть_жалоба'].sum()) if 'Есть_жалоба' in df.columns else 0

    status_col = None
    if 'Статус_связи' in df.columns:
        status_col = 'Статус_связи'
    elif 'Статус_Sheets' in df.columns:
        status_col = 'Статус_Sheets'

    if status_col:
        negative_mask = df[status_col].astype(str).str.contains(
            r"отриц|не удалось|недозвон|не дозвон|не ответ|занят|занято|сброс",
            case=False,
            na=False
        )
    else:
        negative_mask = pd.Series(False, index=df.index)

    negative = int(negative_mask.sum())
    negative_or_complaint = int((negative_mask | (df['Есть_жалоба'] == True)).sum()) if 'Есть_жалоба' in df.columns else negative

    summary = pd.DataFrame([
        {"Показатель": "Всего_записей", "Количество": total},
        {"Показатель": "Жалобы", "Количество": complaints},
        {"Показатель": "Отрицательные", "Количество": negative},
        {"Показатель": "Отрицательные_или_жалобы", "Количество": negative_or_complaint},
    ])
    return summary


def build_region_complaint_matrix(df):
    if 'Регион_112' not in df.columns or 'Жалоба' not in df.columns:
        return pd.DataFrame({"Примечание": ["Нет данных о регионах или жалобах"]})

    df = df.copy()
    df['Регион_112'] = df['Регион_112'].fillna('Неизвестно')

    all_regions = sorted(df['Регион_112'].astype(str).unique().tolist())

    complaints_df = df[df['Жалоба'].astype(str).str.strip() != ''].copy()
    if complaints_df.empty:
        matrix = pd.DataFrame(index=all_regions)
        matrix['ИТОГО'] = 0
        return matrix

    matrix = complaints_df.pivot_table(
        index='Регион_112',
        columns='Жалоба',
        values='Инцидент_112',
        aggfunc='count',
        fill_value=0
    )

    matrix = matrix.reindex(all_regions, fill_value=0)
    matrix['ИТОГО'] = matrix.sum(axis=1)
    matrix = matrix.sort_values('ИТОГО', ascending=False)

    return matrix


def classify_status(value: str) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 'Неизвестно'
    text = str(value).strip().lower()
    if text in ('', 'nan', 'none'):
        return 'Неизвестно'

    no_answer_pattern = r"не удалось|недозвон|не дозвон|не ответ|занят|занято|сброс"
    negative_pattern = r"отриц"
    positive_pattern = r"полож"

    if re.search(no_answer_pattern, text):
        return 'Не дозвонились'
    if re.search(negative_pattern, text):
        return 'Отрицательные'
    if re.search(positive_pattern, text):
        return 'Положительные'
    return 'Прочее'


def build_region_status_stats(df):
    if 'Регион_112' not in df.columns:
        return pd.DataFrame({"Примечание": ["Нет данных о регионах"]})

    status_col = None
    if 'Статус_связи' in df.columns:
        status_col = 'Статус_связи'
    elif 'Статус_Sheets' in df.columns:
        status_col = 'Статус_Sheets'

    if not status_col:
        return pd.DataFrame({"Примечание": ["Нет данных о статусе связи"]})

    df = df.copy()
    df['Регион_112'] = df['Регион_112'].fillna('Неизвестно')
    df['Статус_категория'] = df[status_col].apply(classify_status)

    pivot = df.pivot_table(
        index='Регион_112',
        columns='Статус_категория',
        values='Инцидент_112',
        aggfunc='count',
        fill_value=0
    )

    for col in ['Положительные', 'Отрицательные', 'Не дозвонились', 'Прочее', 'Неизвестно']:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot['Всего'] = pivot.sum(axis=1)
    pivot = pivot[['Положительные', 'Отрицательные', 'Не дозвонились', 'Прочее', 'Неизвестно', 'Всего']]
    pivot = pivot.sort_values('Всего', ascending=False)
    return pivot


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

    out_dir = Path('reports') / '2026-01_services_by_incident_summary'
    out_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

    for service_num in ['101', '102', '103', '104']:
        df_service = df_112[df_112['Служба_112'] == service_num].copy()
        if df_service.empty:
            print(f"⚠️ Нет данных для службы {service_num}")
            continue

        summary = build_summary_counts(df_service)
        matrix = build_region_complaint_matrix(df_service)
        region_status = build_region_status_stats(df_service)

        out_xlsx = out_dir / f'СЛУЖБА_{service_num}_СВОДКА_И_МАТРИЦА_{timestamp}.xlsx'
        with pd.ExcelWriter(out_xlsx, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            summary.to_excel(writer, index=False, sheet_name='Сводка_количество')
            matrix.to_excel(writer, sheet_name='Матрица_жалоб')
            df_service.to_excel(writer, index=False, sheet_name='Детальные')
            region_status.to_excel(writer, sheet_name='Статусы_по_регионам')

        print(f"✅ Сохранено: {out_xlsx}")

    print(f"\nГотово. Папка: {out_dir}")


if __name__ == '__main__':
    main()
