#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Создание таблиц отчетов напрямую в PostgreSQL с использованием SQL
Намного быстрее, чем pandas + psycopg2
"""

import os
import psycopg2
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


def get_db_connection():
    """Создать подключение к БД"""
    return psycopg2.connect(**DB_CONFIG)


def run_query(conn, query, description):
    """Выполнить SQL-запрос и показать результат"""
    print(f"📊 {description}...")
    cursor = conn.cursor()
    try:
        cursor.execute(query)
        conn.commit()
        row_count = cursor.rowcount
        print(f"✓ {description} - {row_count} строк")
        return True
    except Exception as e:
        conn.rollback()
        print(f"❌ Ошибка: {e}")
        return False
    finally:
        cursor.close()


def main():
    """Создать отчеты в SQL"""
    print("="*70)
    print("СОЗДАНИЕ ТАБЛИЦ ОТЧЕТОВ В POSTGRESQL")
    print("="*70)
    
    conn = get_db_connection()
    start = datetime.now()
    
    cursor = conn.cursor()
    
    # 1. Создать основную таблицу reportов
    print("\n1️⃣  Создание таблицы detailed_reports...")
    query_detailed = """
    DROP TABLE IF EXISTS detailed_reports;
    CREATE TABLE detailed_reports AS
    SELECT 
        f.fixation_id,
        f.card_number,
        f.call_date AS "call_datetime",
        r.region_name AS "region",
        s.service_name AS "service",
        f.phone AS "phone",
        f.incident_number AS "incident",
        f.complaint AS "complaint",
        f.status AS "status_text",
        f.reason,
        f.description AS "comment",
        o.operator_name,
        f.source_file,
        f.collection_date AS "closed_date"
    FROM fixations f
    LEFT JOIN operators o ON f.operator_id = o.operator_id
    LEFT JOIN regions r ON f.region_id = r.region_id
    LEFT JOIN services s ON f.service_id = s.service_id
    ORDER BY f.call_date;
    """
    
    cursor.execute(query_detailed)
    conn.commit()
    count = cursor.rowcount
    print(f"✓ Создана таблица detailed_reports с {count:,} записями")
    
    # 2. Создать таблицу жалоб
    print("\n2️⃣  Создание таблицы negative_complaints...")
    query_negative = """
    DROP TABLE IF EXISTS negative_complaints;
    CREATE TABLE negative_complaints AS
    SELECT *
    FROM detailed_reports
    WHERE LOWER(status_text) LIKE '%жалоб%'
        OR LOWER(status_text) LIKE '%отриц%'
        OR LOWER(status_text) LIKE '%негатив%';
    """
    
    cursor.execute(query_negative)
    conn.commit()
    count = cursor.rowcount
    print(f"✓ Создана таблица negative_complaints с {count:,} записями")
    
    # 3. Статистика по регионам
    print("\n3️⃣  Создание таблицы complaints_by_region...")
    query_regions = """
    DROP TABLE IF EXISTS complaints_by_region;
    CREATE TABLE complaints_by_region AS
    SELECT 
        COALESCE(region, '(не указан)') as region,
        COUNT(*) as total_calls,
        COUNT(*) FILTER (WHERE LOWER(status_text) LIKE '%жалоб%') as complaints_count,
        COUNT(*) FILTER (WHERE LOWER(status_text) LIKE '%отриц%') as negative_count,
        COUNT(*) FILTER (WHERE LOWER(status_text) LIKE '%Положительно%') as positive_count
    FROM detailed_reports
    GROUP BY region
    ORDER BY total_calls DESC;
    """
    
    cursor.execute(query_regions)
    conn.commit()
    count = cursor.rowcount
    print(f"✓ Создана таблица complaints_by_region с {count:,} записями")
    
    # 4. Итоговые статистики
    print("\n4️⃣  Создание таблицы summary_statistics...")
    query_summary = """
    DROP TABLE IF EXISTS summary_statistics;
    CREATE TABLE summary_statistics AS
    SELECT 
        COUNT(*) as total_calls,
        COUNT(*) FILTER (WHERE LOWER(status_text) LIKE '%жалоб%') as total_complaints,
        COUNT(*) FILTER (WHERE LOWER(status_text) LIKE '%отриц%') as negative_feedback,
        COUNT(*) FILTER (WHERE LOWER(status_text) LIKE '%Положительно%') as positive_feedback,
        COUNT(DISTINCT COALESCE(region, '')) as regions_count,
        COUNT(DISTINCT operator_name) as operators_count,
        MIN(call_datetime) as first_call,
        MAX(call_datetime) as last_call
    FROM detailed_reports;
    """
    
    cursor.execute(query_summary)
    conn.commit()
    print(f"✓ Создана таблица summary_statistics")
    
    cursor.close()
    conn.close()
    
    # Показать результаты
    elapsed = datetime.now() - start
    print("\n" + "="*70)
    print("✅ ВСЕ ТАБЛИЦЫ СОЗДАНЫ УСПЕШНО")
    print("="*70)
    print(f"Время выполнения: {elapsed}")
    print("\nСозданные таблицы:")
    print("  • detailed_reports - основной отчет")
    print("  • negative_complaints - жалобы и отрицательные отзывы")
    print("  • complaints_by_region - статистика по регионам")
    print("  • summary_statistics - итоговая статистика")
    print("="*70)


if __name__ == '__main__':
    main()
