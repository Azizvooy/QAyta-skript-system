#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
СОЗДАНИЕ СХЕМЫ POSTGRESQL ДЛЯ СИСТЕМЫ QAYTA
=============================================================================
Создает оптимизированную структуру БД PostgreSQL для анализа данных
=============================================================================
"""

import psycopg2
from psycopg2 import sql
from pathlib import Path
from datetime import datetime
import os
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent.parent
CONFIG_DIR = BASE_DIR / 'config'

# Загрузка конфигурации
load_dotenv(CONFIG_DIR / 'postgresql.env')

# Параметры подключения
DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', 5432)),
    'database': os.getenv('DB_NAME', 'qayta_data'),
    'user': os.getenv('DB_USER', 'qayta_user'),
    'password': os.getenv('DB_PASSWORD', 'qayta_password_2026')
}

print('\n' + '='*80)
print('🐘 СОЗДАНИЕ СХЕМЫ POSTGRESQL ДЛЯ QAYTA')
print('='*80)
print(f'\nПодключение к: {DB_CONFIG["host"]}:{DB_CONFIG["port"]}/{DB_CONFIG["database"]}')

def create_database():
    """Создание базы данных если не существует"""
    try:
        # Подключаемся к postgres для создания БД
        conn = psycopg2.connect(
            host=DB_CONFIG['host'],
            port=DB_CONFIG['port'],
            database='postgres',
            user=DB_CONFIG['user'],
            password=DB_CONFIG['password']
        )
        conn.autocommit = True
        cursor = conn.cursor()
        
        # Проверяем существование БД
        cursor.execute(
            "SELECT 1 FROM pg_database WHERE datname = %s",
            (DB_CONFIG['database'],)
        )
        
        if not cursor.fetchone():
            cursor.execute(
                sql.SQL("CREATE DATABASE {}").format(
                    sql.Identifier(DB_CONFIG['database'])
                )
            )
            print(f"✅ База данных '{DB_CONFIG['database']}' создана")
        else:
            print(f"ℹ️  База данных '{DB_CONFIG['database']}' уже существует")
        
        cursor.close()
        conn.close()
        
    except psycopg2.Error as e:
        print(f"❌ Ошибка при создании БД: {e}")
        return False
    
    return True

def create_schema():
    """Создание структуры таблиц"""
    
    if not create_database():
        return
    
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        print('\n[1/7] Создание таблицы операторов...')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS operators (
                operator_id SERIAL PRIMARY KEY,
                operator_name VARCHAR(255) UNIQUE NOT NULL,
                phone VARCHAR(50),
                email VARCHAR(255),
                position VARCHAR(100) DEFAULT 'Оператор',
                active BOOLEAN DEFAULT TRUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            CREATE INDEX IF NOT EXISTS idx_operator_name ON operators(operator_name);
            CREATE INDEX IF NOT EXISTS idx_operator_active ON operators(active);
        """)
        print('  ✅ Таблица operators создана')
        
        print('\n[2/7] Создание таблицы служб...')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS services (
                service_id SERIAL PRIMARY KEY,
                service_code VARCHAR(10) UNIQUE NOT NULL,
                service_name VARCHAR(255) NOT NULL,
                description TEXT,
                active BOOLEAN DEFAULT TRUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            -- Добавляем стандартные службы
            INSERT INTO services (service_code, service_name) 
            VALUES 
                ('101', 'Пожарная служба'),
                ('102', 'Скорая медицинская помощь'),
                ('103', 'Полиция'),
                ('104', 'Аварийная газовая служба')
            ON CONFLICT (service_code) DO NOTHING;
            
            CREATE INDEX IF NOT EXISTS idx_service_code ON services(service_code);
        """)
        print('  ✅ Таблица services создана')
        
        print('\n[3/7] Создание таблицы регионов...')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS regions (
                region_id SERIAL PRIMARY KEY,
                region_name VARCHAR(255) UNIQUE NOT NULL,
                region_code VARCHAR(50),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            CREATE INDEX IF NOT EXISTS idx_region_name ON regions(region_name);
        """)
        print('  ✅ Таблица regions создана')
        
        print('\n[4/7] Создание таблицы фиксаций...')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS fixations (
                fixation_id BIGSERIAL PRIMARY KEY,
                card_number VARCHAR(50),
                operator_id INTEGER REFERENCES operators(operator_id),
                service_id INTEGER REFERENCES services(service_id),
                region_id INTEGER REFERENCES regions(region_id),
                
                -- Данные обращения
                call_date TIMESTAMP,
                incident_number VARCHAR(100),
                phone VARCHAR(50),
                caller_name VARCHAR(255),
                address TEXT,
                district VARCHAR(255),
                
                -- Статус и результат
                status VARCHAR(255),
                status_category VARCHAR(50),
                reason TEXT,
                complaint TEXT,
                description TEXT,
                
                -- Метаданные
                source_file VARCHAR(500),
                import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                collection_date TIMESTAMP,
                
                -- Индексы для быстрого поиска
                CONSTRAINT fixations_card_number_idx UNIQUE (card_number, call_date)
            );
            
            CREATE INDEX IF NOT EXISTS idx_fixations_card ON fixations(card_number);
            CREATE INDEX IF NOT EXISTS idx_fixations_operator ON fixations(operator_id);
            CREATE INDEX IF NOT EXISTS idx_fixations_service ON fixations(service_id);
            CREATE INDEX IF NOT EXISTS idx_fixations_region ON fixations(region_id);
            CREATE INDEX IF NOT EXISTS idx_fixations_date ON fixations(call_date);
            CREATE INDEX IF NOT EXISTS idx_fixations_status ON fixations(status_category);
            CREATE INDEX IF NOT EXISTS idx_fixations_import ON fixations(import_date);
        """)
        print('  ✅ Таблица fixations создана')
        
        print('\n[5/7] Создание таблицы инцидентов 112...')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS incidents_112 (
                incident_id BIGSERIAL PRIMARY KEY,
                incident_number VARCHAR(100) UNIQUE,
                card_number VARCHAR(50),
                service_id INTEGER REFERENCES services(service_id),
                region_id INTEGER REFERENCES regions(region_id),
                
                -- Данные инцидента
                call_time TIMESTAMP,
                caller_phone VARCHAR(50),
                caller_name VARCHAR(255),
                address TEXT,
                district VARCHAR(255),
                reason TEXT,
                status VARCHAR(255),
                
                -- Данные обработки
                operator_112 VARCHAR(255),
                close_time TIMESTAMP,
                duration INTERVAL,
                
                -- Метаданные
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            CREATE INDEX IF NOT EXISTS idx_incidents_number ON incidents_112(incident_number);
            CREATE INDEX IF NOT EXISTS idx_incidents_card ON incidents_112(card_number);
            CREATE INDEX IF NOT EXISTS idx_incidents_service ON incidents_112(service_id);
            CREATE INDEX IF NOT EXISTS idx_incidents_region ON incidents_112(region_id);
            CREATE INDEX IF NOT EXISTS idx_incidents_time ON incidents_112(call_time);
        """)
        print('  ✅ Таблица incidents_112 создана')
        
        print('\n[6/7] Создание аналитических представлений...')
        cursor.execute("""
            -- Представление: Полная информация о фиксациях
            CREATE OR REPLACE VIEW v_fixations_full AS
            SELECT 
                f.fixation_id,
                f.card_number,
                f.incident_number,
                f.call_date,
                f.phone,
                f.caller_name,
                f.address,
                f.district,
                f.status,
                f.status_category,
                f.reason,
                f.complaint,
                o.operator_name,
                s.service_code,
                s.service_name,
                r.region_name,
                f.source_file,
                f.import_date
            FROM fixations f
            LEFT JOIN operators o ON f.operator_id = o.operator_id
            LEFT JOIN services s ON f.service_id = s.service_id
            LEFT JOIN regions r ON f.region_id = r.region_id;
            
            -- Представление: Статистика по операторам
            CREATE OR REPLACE VIEW v_operator_statistics AS
            SELECT 
                o.operator_id,
                o.operator_name,
                COUNT(f.fixation_id) as total_fixations,
                COUNT(CASE WHEN f.status_category = 'Положительно' THEN 1 END) as positive_count,
                COUNT(CASE WHEN f.status_category = 'Отрицательно' THEN 1 END) as negative_count,
                COUNT(CASE WHEN f.status_category = 'Не дозвонились' THEN 1 END) as no_answer_count,
                ROUND(
                    COUNT(CASE WHEN f.status_category = 'Положительно' THEN 1 END)::NUMERIC / 
                    NULLIF(COUNT(f.fixation_id), 0) * 100, 2
                ) as positive_percentage
            FROM operators o
            LEFT JOIN fixations f ON o.operator_id = f.operator_id
            GROUP BY o.operator_id, o.operator_name;
            
            -- Представление: Статистика по службам
            CREATE OR REPLACE VIEW v_service_statistics AS
            SELECT 
                s.service_id,
                s.service_code,
                s.service_name,
                COUNT(f.fixation_id) as total_fixations,
                COUNT(CASE WHEN f.status_category = 'Положительно' THEN 1 END) as positive_count,
                COUNT(CASE WHEN f.status_category = 'Отрицательно' THEN 1 END) as negative_count,
                COUNT(DISTINCT f.region_id) as regions_count
            FROM services s
            LEFT JOIN fixations f ON s.service_id = f.service_id
            GROUP BY s.service_id, s.service_code, s.service_name;
            
            -- Представление: Статистика по регионам
            CREATE OR REPLACE VIEW v_region_statistics AS
            SELECT 
                r.region_id,
                r.region_name,
                COUNT(f.fixation_id) as total_fixations,
                COUNT(CASE WHEN f.status_category = 'Положительно' THEN 1 END) as positive_count,
                COUNT(CASE WHEN f.status_category = 'Отрицательно' THEN 1 END) as negative_count,
                COUNT(DISTINCT f.service_id) as services_count
            FROM regions r
            LEFT JOIN fixations f ON r.region_id = f.region_id
            GROUP BY r.region_id, r.region_name;
        """)
        print('  ✅ Аналитические представления созданы')
        
        print('\n[7/7] Создание функций для анализа...')
        cursor.execute("""
            -- Функция: Категоризация статуса
            CREATE OR REPLACE FUNCTION categorize_status(status_text TEXT)
            RETURNS VARCHAR(50) AS $$
            BEGIN
                IF status_text IS NULL THEN
                    RETURN 'Прочее';
                END IF;
                
                status_text := LOWER(status_text);
                
                IF status_text LIKE '%положительн%' OR 
                   status_text LIKE '%qanoatlantir%' OR
                   status_text LIKE '%қаноатлантир%' THEN
                    RETURN 'Положительно';
                ELSIF status_text LIKE '%отрицательн%' OR
                      status_text LIKE '%qanoatlantirilmadi%' OR
                      status_text LIKE '%нет ответа%' OR
                      status_text LIKE '%жалоб%' THEN
                    RETURN 'Отрицательно';
                ELSIF status_text LIKE '%занято%' OR
                      status_text LIKE '%не дозвон%' THEN
                    RETURN 'Не дозвонились';
                ELSE
                    RETURN 'Прочее';
                END IF;
            END;
            $$ LANGUAGE plpgsql IMMUTABLE;
            
            -- Триггер: Автоматическая категоризация при вставке
            CREATE OR REPLACE FUNCTION trigger_categorize_status()
            RETURNS TRIGGER AS $$
            BEGIN
                NEW.status_category := categorize_status(NEW.status);
                NEW.updated_at := CURRENT_TIMESTAMP;
                RETURN NEW;
            END;
            $$ LANGUAGE plpgsql;
            
            DROP TRIGGER IF EXISTS before_insert_fixation ON fixations;
            CREATE TRIGGER before_insert_fixation
                BEFORE INSERT OR UPDATE ON fixations
                FOR EACH ROW
                EXECUTE FUNCTION trigger_categorize_status();
        """)
        print('  ✅ Функции и триггеры созданы')
        
        conn.commit()
        cursor.close()
        conn.close()
        
        print('\n' + '='*80)
        print('✅ СХЕМА POSTGRESQL УСПЕШНО СОЗДАНА!')
        print('='*80)
        print(f'\n📊 Созданные объекты:')
        print('   Таблицы:')
        print('     • operators (операторы)')
        print('     • services (службы: 101, 102, 103, 104)')
        print('     • regions (регионы)')
        print('     • fixations (фиксации)')
        print('     • incidents_112 (инциденты 112)')
        print('\n   Представления:')
        print('     • v_fixations_full (полная информация)')
        print('     • v_operator_statistics (статистика операторов)')
        print('     • v_service_statistics (статистика служб)')
        print('     • v_region_statistics (статистика регионов)')
        print('\n   Функции:')
        print('     • categorize_status() (категоризация статусов)')
        print('     • trigger_categorize_status() (авто-категоризация)')
        print('\n' + '='*80)
        
    except psycopg2.Error as e:
        print(f'\n❌ Ошибка при создании схемы: {e}')
        if conn:
            conn.rollback()

if __name__ == '__main__':
    create_schema()
