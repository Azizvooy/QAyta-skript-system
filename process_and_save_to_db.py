#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Обработка данных и сохранение результатов в PostgreSQL
Вместо создания Excel файлов, создает таблицы в БД:
- detailed_reports - детальные отчеты
- negative_complaints - отрицательные и жалобы
- complaints_by_region - жалобы по регионам
- regions_complaints_pivot - pivot регионы-жалобы
- not_found_applications - не найденные заявки
"""

import os
import re
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
from pathlib import Path
from dotenv import load_dotenv
from datetime import datetime

# Загрузка конфигурации
BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / 'config'
load_dotenv(CONFIG_DIR / 'postgresql.env')

# Параметры подключения к БД
DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': os.getenv('DB_PORT', '5432'),
    'database': os.getenv('DB_NAME', 'qayta_data'),
    'user': os.getenv('DB_USER', 'qayta_user'),
    'password': os.getenv('DB_PASSWORD', 'qayta_password_2026')
}

# Колонки для удаления
DROP_COLUMNS = [
    'Дата_112',
    'Қўнғироқ давомийлиги',
    'Бригадага узатилган вақт',
    'Қўнғироқ якунланган вақт',
    'Статус_112',
    'Қўнғироқ қилувчи Ф.И.Ш',
    'Телефон_112',
    'Ўзи рад этган',
    'Есть_жалоба'
]

# Колонки которые должны остаться (17 колонок)
KEEP_COLUMNS = [
    'Қўнғироқ қабул қилинган вақт',
    'Регион_112',
    'Служба_112',
    'Телефон_нормализованный',
    'Инцидент_112',
    'Жалоба',
    'Статус_связи',
    'Телефон_Sheets',
    'Дата_открытия',
    'Служба_Sheets',
    'Положительно',
    'Дата_закрытия',
    'Инцидент_Sheets',
    'Статус_Sheets',
    'Комментарий',
    'Примечание',
    'Категория'
]

DATE_COLUMN = 'Қўнғироқ қабул қилинган вақт'


def get_db_connection():
    """Создать подключение к БД"""
    return psycopg2.connect(**DB_CONFIG)


def classify_status(val):
    """Классификация статуса"""
    if pd.isna(val):
        return 'Неизвестно'
    text = str(val).strip().lower()
    if not text:
        return 'Неизвестно'
    
    negative_keywords = ['отриц', 'негатив', 'неудов', 'плох']
    positive_keywords = ['полож', 'позитив', 'удов', 'хорош']
    neutral_keywords = ['закрыт', 'завер', 'выполн']
    no_contact = ['не дозвон', 'недозвон', 'не отвеча', 'нет связ']
    
    for kw in negative_keywords:
        if kw in text:
            return 'Отрицательно'
    for kw in positive_keywords:
        if kw in text:
            return 'Положительно'
    for kw in no_contact:
        if kw in text:
            return 'Не дозвонились'
    for kw in neutral_keywords:
        if kw in text:
            return 'Заявка закрыта'
    
    return 'Другое'


def clean_complaint(text):
    """Удалить префиксы 1./2./3./4. из жалоб"""
    if pd.isna(text) or not text:
        return ''
    text = str(text).strip()
    text = re.sub(r'^[1-4]\.\s*', '', text)
    return text


def load_data_from_db():
    """Загрузить данные из БД"""
    print("📊 Загрузка данных из PostgreSQL...")
    conn = get_db_connection()
    
    query = """
        SELECT 
            f.fixation_id,
            f.card_number,
            f.call_date AS "Қўнғироқ қабул қилинган вақт",
            r.region_name AS "Регион_112",
            s.service_name AS "Служба_112",
            f.phone AS "Телефон_нормализованный",
            f.incident_number AS "Инцидент_112",
            f.complaint AS "Жалоба",
            f.status AS "Статус_связи",
            f.status_category AS "Категория_статуса",
            f.reason,
            f.description AS "Комментарий",
            o.operator_name,
            f.source_file,
            f.collection_date AS "Дата_закрытия"
        FROM fixations f
        LEFT JOIN operators o ON f.operator_id = o.operator_id
        LEFT JOIN regions r ON f.region_id = r.region_id
        LEFT JOIN services s ON f.service_id = s.service_id
        ORDER BY f.call_date
    """
    
    df = pd.read_sql(query, conn)
    conn.close()
    
    print(f"✓ Загружено {len(df):,} записей из БД")
    return df


def process_detailed_data(df):
    """Обработать детальные данные"""
    print("\n📋 Обработка детальных данных...")
    
    df_detailed = df.copy()
    
    # Колонки которые должны быть 
    required_cols = [
        'Қўнғироқ қабул қилинган вақт',
        'Регион_112',
        'Служба_112',
        'Телефон_нормализованный',
        'Инцидент_112',
        'Жалоба',
        'Статус_связи',
        'Категория_статуса',
        'Комментарий',
        'Дата_закрытия'
    ]
    
    # Оставить только существующие колонки
    existing_cols = [col for col in required_cols if col in df_detailed.columns]
    df_detailed = df_detailed[existing_cols]
    
    # Преобразовать дату
    if 'Қўнғироқ қабул қилинган вақт' in df_detailed.columns:
        df_detailed['Қўнғироқ қабул қилинган вақт'] = pd.to_datetime(
            df_detailed['Қўнғироқ қабул қилинган вақт'], 
            errors='coerce'
        )
        df_detailed = df_detailed.sort_values('Қўнғироқ қабул қилинган вақт')
    
    # Добавить нумерацию
    df_detailed.insert(0, '№', range(1, len(df_detailed) + 1))
    
    # Очистить жалобы от префиксов
    if 'Жалоба' in df_detailed.columns:
        df_detailed['Жалоба'] = df_detailed['Жалоба'].apply(clean_complaint)
    
    # Если нет категории статуса, классифицировать
    if 'Категория_статуса' not in df_detailed.columns and 'Статус_связи' in df_detailed.columns:
        df_detailed['Категория_статуса'] = df_detailed['Статус_связи'].apply(classify_status)
    
    print(f"✓ Обработано {len(df_detailed):,} записей")
    return df_detailed


def create_negative_complaints(df_detailed):
    """Создать таблицу отрицательных отзывов и жалоб"""
    print("\n❌ Создание таблицы отрицательных отзывов...")
    
    if 'Статус_связи' not in df_detailed.columns:
        return pd.DataFrame()
    
    # Поиск отрицательных статусов в колонке Статус_связи
    negative_keywords = ['отриц', 'негатив', 'неудов', 'плох', 'жалоб']
    
    # Проверить каждую строку
    mask = df_detailed['Статус_связи'].fillna('').str.lower().apply(
        lambda x: any(kw in x for kw in negative_keywords) if x else False
    )
    
    df_negative = df_detailed[mask].copy()
    
    print(f"✓ Найдено {len(df_negative):,} отрицательных записей")
    return df_negative


def create_complaints_by_region(df_detailed):
    """Создать сводку жалоб по регионам"""
    print("\n🗺️ Создание сводки по регионам...")
    
    if 'Регион_112' not in df_detailed.columns:
        return pd.DataFrame()
    
    # Группировка по регионам с подсчетом различных типов статусов
    summary = df_detailed.groupby('Регион_112').agg({
        'Телефон_нормализованный': 'count'
    }).reset_index()
    
    summary.columns = ['Регион', 'Всего_звонков']
    
    # Добавить подсчет жалоб (записи со статусом содержащим 'жалоб')
    complaints_mask = df_detailed['Статус_связи'].fillna('').str.lower().str.contains('жалоб', na=False)
    complaints_by_region = df_detailed[complaints_mask].groupby('Регион_112').size().reset_index(name='Жалоб')
    
    summary = summary.merge(complaints_by_region, left_on='Регион', right_on='Регион_112', how='left')
    summary['Жалоб'] = summary['Жалоб'].fillna(0).astype(int)
    
    if 'Регион_112' in summary.columns:
        summary = summary.drop('Регион_112', axis=1)
    
    print(f"✓ Создана сводка по {len(summary)} регионам")
    return summary


def create_regions_pivot(df_detailed):
    """Создать pivot таблицу регионы-жалобы"""
    print("\n📊 Создание pivot таблицы...")
    
    if 'Регион_112' not in df_detailed.columns or 'Жалоба' not in df_detailed.columns:
        return pd.DataFrame()
    
    # Фильтровать только записи с жалобами
    df_complaints = df_detailed[df_detailed['Жалоба'].notna() & (df_detailed['Жалоба'] != '')].copy()
    
    if len(df_complaints) == 0:
        return pd.DataFrame()
    
    # Создать pivot
    pivot = pd.crosstab(
        df_complaints['Регион_112'],
        df_complaints['Жалоба'],
        margins=True,
        margins_name='ИТОГО'
    )
    
    print(f"✓ Создана pivot таблица {pivot.shape[0]} x {pivot.shape[1]}")
    return pivot


def create_not_found_applications(df_detailed):
    """Создать таблицу не найденных заявок"""
    print("\n🔍 Создание таблицы не найденных заявок...")
    
    # Записи без информации о жалобах и с неизвестным статусом
    if 'Статус_связи' in df_detailed.columns:
        unknown_keywords = ['неизвестно', 'другое', 'ошибка', 'системы']
        mask = df_detailed['Статус_связи'].fillna('').str.lower().apply(
            lambda x: any(kw in x for kw in unknown_keywords) if x else False
        )
        df_not_found = df_detailed[mask].copy()
    else:
        df_not_found = pd.DataFrame()
    
    print(f"✓ Найдено {len(df_not_found):,} не найденных заявок")
    return df_not_found


def save_to_database(table_name, df, conn):
    """Сохранить DataFrame в таблицу PostgreSQL"""
    if len(df) == 0:
        print(f"⚠️  Таблица {table_name} пустая, пропуск...")
        return
    
    print(f"💾 Сохранение {len(df):,} записей в таблицу {table_name}...")
    
    # Заменить NaT и NaN на None для PostgreSQL
    df = df.replace({pd.NaT: None, pd.NA: None})
    df = df.where(pd.notna(df), None)
    
    cursor = conn.cursor()
    
    # Удалить таблицу если существует
    cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
    
    # Создать таблицу
    columns_def = []
    for col in df.columns:
        dtype = df[col].dtype
        if dtype == 'int64':
            sql_type = 'INTEGER'
        elif dtype == 'float64':
            sql_type = 'REAL'
        elif dtype == 'datetime64[ns]':
            sql_type = 'TIMESTAMP'
        else:
            sql_type = 'TEXT'
        
        col_name = col.replace(' ', '_').replace('№', 'num')
        columns_def.append(f'"{col_name}" {sql_type}')
    
    create_query = f"""
        CREATE TABLE {table_name} (
            {', '.join(columns_def)}
        )
    """
    cursor.execute(create_query)
    
    # Вставить данные
    columns = [col.replace(' ', '_').replace('№', 'num') for col in df.columns]
    values = [tuple(row) for row in df.values]
    
    insert_query = f"""
        INSERT INTO {table_name} ({', '.join([f'"{col}"' for col in columns])})
        VALUES %s
    """
    
    execute_values(cursor, insert_query, values)
    conn.commit()
    
    print(f"✓ Таблица {table_name} создана и заполнена")


def main():
    """Основная функция"""
    print("="*70)
    print("ОБРАБОТКА ДАННЫХ И СОХРАНЕНИЕ В POSTGRESQL")
    print("="*70)
    
    start_time = datetime.now()
    
    try:
        # 1. Загрузить данные
        df = load_data_from_db()
        
        # 2. Обработать детальные данные
        df_detailed = process_detailed_data(df)
        
        # 3. Создать производные таблицы
        print("\nСоздание производных таблиц...")
        df_negative = create_negative_complaints(df_detailed)
        
        # Пропустить медленные операции, которые зависают
        # df_regions = create_complaints_by_region(df_detailed)
        # df_pivot = create_regions_pivot(df_detailed)
        df_not_found = create_not_found_applications(df_detailed)
        
        # 4. Сохранить в БД
        print("\n" + "="*70)
        print("СОХРАНЕНИЕ В БАЗУ ДАННЫХ")
        print("="*70)
        
        conn = get_db_connection()
        
        save_to_database('detailed_reports', df_detailed, conn)
        save_to_database('negative_complaints', df_negative, conn)
        # skip: save_to_database('complaints_by_region', df_regions, conn)
        # skip: save_to_database('regions_complaints_pivot', df_pivot, conn)
        save_to_database('not_found_applications', df_not_found, conn)
        
        conn.close()
        
        # 5. Итоги
        elapsed = datetime.now() - start_time
        print("\n" + "="*70)
        print("✅ ОБРАБОТКА ЗАВЕРШЕНА")
        print("="*70)
        print(f"Время выполнения: {elapsed}")
        print(f"\nСозданные таблицы:")
        print(f"  • detailed_reports: {len(df_detailed):,} записей")
        print(f"  • negative_complaints: {len(df_negative):,} записей")
        print(f"  • not_found_applications: {len(df_not_found):,} записей")
        print("="*70)
        
    except Exception as e:
        print(f"\n❌ ОШИБКА: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == '__main__':
    exit(main())
