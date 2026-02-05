#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Создание таблицы для отслеживания контекста разговоров
Чтобы помнить что уже было сделано и не спрашивать повторно
"""
import psycopg2
from dotenv import load_dotenv
import os
from datetime import datetime

# Загружаем конфигурацию
load_dotenv('config/postgresql.env')

try:
    conn = psycopg2.connect(
        host=os.getenv('DB_HOST'),
        port=os.getenv('DB_PORT'),
        dbname=os.getenv('DB_NAME'),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD')
    )
    cur = conn.cursor()
    
    print('📝 Создание таблицы для отслеживания контекста разговоров...')
    
    # Создаем таблицу для хранения контекста разговоров
    cur.execute('''
        CREATE TABLE IF NOT EXISTS conversation_context (
            id SERIAL PRIMARY KEY,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            context_key VARCHAR(255) UNIQUE NOT NULL,  -- Ключ контекста (например, "credentials_configured", "sheets_access_granted")
            context_value TEXT,  -- Значение (может быть JSON)
            description TEXT,  -- Описание контекста
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Создаем таблицу для истории действий
    cur.execute('''
        CREATE TABLE IF NOT EXISTS action_history (
            id SERIAL PRIMARY KEY,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            action_type VARCHAR(100) NOT NULL,  -- Тип действия (install, configure, import, etc.)
            action_name VARCHAR(255) NOT NULL,  -- Название действия
            status VARCHAR(50) NOT NULL,  -- success, failed, in_progress
            details TEXT,  -- Дополнительные детали (JSON)
            error_message TEXT  -- Сообщение об ошибке если было
        )
    ''')
    
    # Создаем индексы
    cur.execute('CREATE INDEX IF NOT EXISTS idx_context_key ON conversation_context(context_key)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_action_timestamp ON action_history(timestamp DESC)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_action_type ON action_history(action_type)')
    
    # Создаем триггер для автообновления updated_at
    cur.execute('''
        CREATE OR REPLACE FUNCTION update_conversation_context_timestamp()
        RETURNS TRIGGER AS $$
        BEGIN
            NEW.updated_at = CURRENT_TIMESTAMP;
            RETURN NEW;
        END;
        $$ LANGUAGE plpgsql
    ''')
    
    cur.execute('''
        DROP TRIGGER IF EXISTS trigger_update_conversation_context ON conversation_context
    ''')
    
    cur.execute('''
        CREATE TRIGGER trigger_update_conversation_context
        BEFORE UPDATE ON conversation_context
        FOR EACH ROW
        EXECUTE FUNCTION update_conversation_context_timestamp()
    ''')
    
    # Записываем текущее состояние системы
    now = datetime.now().isoformat()
    
    context_records = [
        ('credentials_configured', 'true', 'Google Sheets API credentials.json настроен'),
        ('sheets_access_granted', 'true', 'Доступ к Google Sheets предоставлен для service account'),
        ('postgresql_setup', 'true', 'PostgreSQL база данных настроена и запущена в Docker'),
        ('docker_running', 'true', 'Docker Desktop запущен и работает'),
        ('wsl_updated', 'true', 'WSL обновлен до последней версии'),
        ('data_imported', 'true', f'Данные импортированы из Google Sheets API ({now})'),
        ('schema_created', 'true', 'Схема БД создана: operators, services, regions, fixations, incidents_112'),
        ('views_created', 'true', 'Созданы представления: v_fixations_full, v_operator_statistics, v_service_statistics, v_region_statistics'),
        ('auto_categorization', 'true', 'Автоматическая категоризация статусов работает (trigger)'),
    ]
    
    for key, value, description in context_records:
        cur.execute('''
            INSERT INTO conversation_context (context_key, context_value, description)
            VALUES (%s, %s, %s)
            ON CONFLICT (context_key) 
            DO UPDATE SET 
                context_value = EXCLUDED.context_value,
                description = EXCLUDED.description,
                updated_at = CURRENT_TIMESTAMP
        ''', (key, value, description))
    
    # Записываем историю действий
    action_records = [
        ('setup', 'Docker Compose Configuration', 'success', 'Создан docker-compose.yml с PostgreSQL 16 и pgAdmin'),
        ('setup', 'Database Schema Creation', 'success', '5 таблиц, 4 представления, 2 функции с триггерами'),
        ('setup', 'WSL Update', 'success', 'WSL обновлен командой wsl --update'),
        ('setup', 'Docker Containers Start', 'success', 'Контейнеры qayta-postgres и qayta-pgadmin запущены'),
        ('install', 'Python Dependencies', 'success', 'psycopg2-binary, python-dotenv, tqdm, google-api-python-client'),
        ('install', 'Google Sheets API Setup', 'success', 'Credentials.json настроен, доступ к таблицам предоставлен'),
        ('import', 'Google Sheets API Import', 'success', f'Данные импортированы успешно ({now})'),
    ]
    
    for action_type, action_name, status, details in action_records:
        cur.execute('''
            INSERT INTO action_history (action_type, action_name, status, details)
            VALUES (%s, %s, %s, %s)
        ''', (action_type, action_name, status, details))
    
    conn.commit()
    
    print('✅ Таблицы для отслеживания контекста созданы!')
    print('\n📊 Записано в контекст:')
    cur.execute('SELECT context_key, description FROM conversation_context ORDER BY id')
    for key, desc in cur.fetchall():
        print(f'  ✓ {desc}')
    
    print('\n📜 История действий:')
    cur.execute('SELECT action_name, status FROM action_history ORDER BY id')
    for name, status in cur.fetchall():
        print(f'  ✓ {name}: {status}')
    
    conn.close()
    
    print('\n✅ Теперь система будет помнить контекст разговоров!')
    
except Exception as e:
    print(f'❌ Ошибка: {e}')
    import traceback
    traceback.print_exc()
