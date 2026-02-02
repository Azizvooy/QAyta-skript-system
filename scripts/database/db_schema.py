#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
СХЕМА БАЗЫ ДАННЫХ С ПРАВИЛЬНОЙ СТРУКТУРОЙ И СВЯЗЯМИ
=============================================================================
Создает оптимизированную БД без пустых значений и неправильных связей
=============================================================================
"""

import sqlite3
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
LOG_DIR = BASE_DIR / 'logs' / 'database'

# Создаем директорию для логов
LOG_DIR.mkdir(parents=True, exist_ok=True)

def log_operation(operation, details, status='SUCCESS'):
    """Логирование всех операций с БД"""
    log_file = LOG_DIR / f'db_log_{datetime.now().strftime("%Y%m%d")}.log'
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'[{timestamp}] [{status}] {operation}: {details}\n')

def create_database_schema():
    """Создать полную схему БД с правильными связями"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    log_operation('CREATE_SCHEMA', 'Начало создания схемы БД')
    
    try:
        # =============================================================================
        # 1. ТАБЛИЦА ОПЕРАТОРОВ (справочник)
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS operators (
                operator_id INTEGER PRIMARY KEY AUTOINCREMENT,
                operator_name TEXT UNIQUE NOT NULL,
                phone TEXT,
                email TEXT,
                position TEXT DEFAULT 'Оператор',
                active BOOLEAN DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Индекс для быстрого поиска
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_operator_name 
            ON operators(operator_name)
        ''')
        
        # =============================================================================
        # 2. ТАБЛИЦА СЛУЖБ (справочник)
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS services (
                service_id INTEGER PRIMARY KEY AUTOINCREMENT,
                service_code TEXT UNIQUE NOT NULL,
                service_name TEXT NOT NULL,
                description TEXT,
                active BOOLEAN DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Добавляем стандартные службы
        services_data = [
            ('101', 'Пожарная служба'),
            ('102', 'Полиция'),
            ('103', 'Скорая помощь'),
            ('104', 'Аварийная газовая служба')
        ]
        
        for code, name in services_data:
            cursor.execute('''
                INSERT OR IGNORE INTO services (service_code, service_name)
                VALUES (?, ?)
            ''', (code, name))
        
        # =============================================================================
        # 3. ТАБЛИЦА РЕГИОНОВ (справочник)
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS regions (
                region_id INTEGER PRIMARY KEY AUTOINCREMENT,
                region_name TEXT UNIQUE NOT NULL,
                region_code TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # =============================================================================
        # 4. ТАБЛИЦА ЗАЯВОК (основная)
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS applications (
                application_id INTEGER PRIMARY KEY AUTOINCREMENT,
                application_number TEXT UNIQUE NOT NULL,
                card_number TEXT,
                incident_number TEXT,
                
                -- Связь со службой
                service_id INTEGER,
                
                -- Связь с регионом
                region_id INTEGER,
                
                -- Данные заявителя
                caller_name TEXT,
                caller_phone TEXT,
                
                -- Данные о происшествии
                call_date DATE,
                call_time TEXT,
                address TEXT,
                district TEXT,
                reason TEXT,
                description TEXT,
                location_type TEXT,
                
                -- Статус заявки
                status TEXT DEFAULT 'Новая',
                close_time TEXT,
                duration TEXT,
                
                -- Примечания
                notes TEXT,
                
                -- Метаданные
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                
                -- Внешние ключи
                FOREIGN KEY (service_id) REFERENCES services(service_id),
                FOREIGN KEY (region_id) REFERENCES regions(region_id)
            )
        ''')
        
        # Индексы для быстрого поиска
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_application_number 
            ON applications(application_number)
        ''')
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_card_number 
            ON applications(card_number)
        ''')
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_call_date 
            ON applications(call_date)
        ''')
        
        # =============================================================================
        # 5. ТАБЛИЦА ФИКСАЦИИ (обзвоны операторов)
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS fixations (
                fixation_id INTEGER PRIMARY KEY AUTOINCREMENT,
                
                -- Связь с заявкой
                application_id INTEGER,
                
                -- Связь с оператором
                operator_id INTEGER NOT NULL,
                
                -- Данные фиксации
                collection_date DATE,
                fixation_date DATETIME,
                phone_called TEXT,
                
                -- Результат обзвона
                status TEXT NOT NULL,
                feedback_type TEXT,
                
                -- Примечания
                notes TEXT,
                
                -- Метаданные
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                
                -- Внешние ключи
                FOREIGN KEY (application_id) REFERENCES applications(application_id),
                FOREIGN KEY (operator_id) REFERENCES operators(operator_id)
            )
        ''')
        
        # Индексы
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_fixation_date 
            ON fixations(collection_date)
        ''')
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_fixation_operator 
            ON fixations(operator_id)
        ''')
        
        # =============================================================================
        # 6. ТАБЛИЦА ЛОГОВ ОПЕРАЦИЙ
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS operation_logs (
                log_id INTEGER PRIMARY KEY AUTOINCREMENT,
                operation_type TEXT NOT NULL,
                table_name TEXT,
                record_id INTEGER,
                old_value TEXT,
                new_value TEXT,
                operator_id INTEGER,
                operation_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                ip_address TEXT,
                success BOOLEAN DEFAULT 1,
                error_message TEXT
            )
        ''')
        
        # =============================================================================
        # 7. ТАБЛИЦА СТАТИСТИКИ (агрегированные данные)
        # =============================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS daily_statistics (
                stat_id INTEGER PRIMARY KEY AUTOINCREMENT,
                stat_date DATE NOT NULL,
                
                -- Связь с оператором
                operator_id INTEGER,
                
                -- Связь со службой
                service_id INTEGER,
                
                -- Статистика
                total_calls INTEGER DEFAULT 0,
                positive_count INTEGER DEFAULT 0,
                negative_count INTEGER DEFAULT 0,
                no_answer_count INTEGER DEFAULT 0,
                
                -- Расчетные поля
                positive_percent REAL,
                efficiency_rate REAL,
                
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                
                FOREIGN KEY (operator_id) REFERENCES operators(operator_id),
                FOREIGN KEY (service_id) REFERENCES services(service_id),
                
                -- Уникальность по дате + оператор + служба
                UNIQUE(stat_date, operator_id, service_id)
            )
        ''')
        
        # =============================================================================
        # 8. ПРЕДСТАВЛЕНИЯ (VIEWS) для удобных запросов
        # =============================================================================
        
        # Полная информация о заявках
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_applications_full AS
            SELECT 
                a.application_id,
                a.application_number,
                a.card_number,
                a.incident_number,
                s.service_code,
                s.service_name,
                r.region_name,
                a.caller_name,
                a.caller_phone,
                a.call_date,
                a.call_time,
                a.address,
                a.district,
                a.reason,
                a.description,
                a.status,
                a.created_at
            FROM applications a
            LEFT JOIN services s ON a.service_id = s.service_id
            LEFT JOIN regions r ON a.region_id = r.region_id
        ''')
        
        # Полная информация о фиксациях
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_fixations_full AS
            SELECT 
                f.fixation_id,
                f.collection_date,
                f.fixation_date,
                o.operator_name,
                a.application_number,
                a.card_number,
                s.service_code,
                s.service_name,
                r.region_name,
                f.phone_called,
                f.status,
                f.feedback_type,
                f.notes,
                f.created_at
            FROM fixations f
            LEFT JOIN operators o ON f.operator_id = o.operator_id
            LEFT JOIN applications a ON f.application_id = a.application_id
            LEFT JOIN services s ON a.service_id = s.service_id
            LEFT JOIN regions r ON a.region_id = r.region_id
        ''')
        
        # Статистика по операторам
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_operator_stats AS
            SELECT 
                o.operator_name,
                COUNT(f.fixation_id) as total_fixations,
                SUM(CASE WHEN f.status LIKE '%Положительн%' THEN 1 ELSE 0 END) as positive,
                SUM(CASE WHEN f.status LIKE '%Отрицательн%' THEN 1 ELSE 0 END) as negative,
                SUM(CASE WHEN f.status LIKE '%Нет ответа%' OR f.status LIKE '%занято%' THEN 1 ELSE 0 END) as no_answer,
                ROUND(
                    CAST(SUM(CASE WHEN f.status LIKE '%Положительн%' THEN 1 ELSE 0 END) AS REAL) * 100.0 / 
                    NULLIF(COUNT(f.fixation_id), 0), 
                    2
                ) as positive_percent
            FROM operators o
            LEFT JOIN fixations f ON o.operator_id = f.operator_id
            GROUP BY o.operator_id, o.operator_name
        ''')
        
        conn.commit()
        log_operation('CREATE_SCHEMA', 'Схема БД успешно создана', 'SUCCESS')
        
        print('OK: База данных успешно создана')
        print(f'Путь к БД: {DB_PATH}')
        print(f'Логи: {LOG_DIR}')
        
        # Показываем созданные таблицы
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = cursor.fetchall()
        print('\nСозданные таблицы:')
        for table in tables:
            print(f'   - {table[0]}')
        
        # Показываем созданные представления
        cursor.execute("SELECT name FROM sqlite_master WHERE type='view' ORDER BY name")
        views = cursor.fetchall()
        print('\nСозданные представления (views):')
        for view in views:
            print(f'   - {view[0]}')
            
    except Exception as e:
        log_operation('CREATE_SCHEMA', f'Ошибка: {str(e)}', 'ERROR')
        print(f'ОШИБКА при создании схемы: {e}')
        conn.rollback()
        raise
    finally:
        conn.close()

if __name__ == '__main__':
    print('=' * 80)
    print('СОЗДАНИЕ ОПТИМИЗИРОВАННОЙ СХЕМЫ БАЗЫ ДАННЫХ')
    print('=' * 80)
    create_database_schema()
    print('\n' + '=' * 80)
    print('ГОТОВО!')
    print('=' * 80)
