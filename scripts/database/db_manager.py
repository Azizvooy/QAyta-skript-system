#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
ГЛАВНЫЙ МЕНЕДЖЕР БАЗЫ ДАННЫХ
=============================================================================
Управление БД: создание, импорт, отчеты, обслуживание
=============================================================================
"""

import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent.parent
sys.path.insert(0, str(BASE_DIR / 'scripts'))

from database.db_schema import create_database_schema, log_operation
from database.db_import import import_applications_from_excel, import_fixations_from_csv
from database.db_reports import create_all_reports

def show_menu():
    """Показать меню"""
    print('\n' + '=' * 80)
    print('УПРАВЛЕНИЕ БАЗОЙ ДАННЫХ')
    print('=' * 80)
    print('\n📋 Выберите действие:')
    print('  1. Создать/обновить схему БД')
    print('  2. Импортировать данные из Excel')
    print('  3. Импортировать данные из CSV')
    print('  4. Создать все отчеты')
    print('  5. Полная инициализация (схема + импорт + отчеты)')
    print('  6. Проверить БД')
    print('  0. Выход')
    print('=' * 80)

def check_database():
    """Проверка состояния БД"""
    import sqlite3
    
    db_path = BASE_DIR / 'data' / 'fiksa_database.db'
    
    if not db_path.exists():
        print('❌ База данных не найдена!')
        print(f'   Путь: {db_path}')
        return False
    
    print(f'✅ База данных найдена: {db_path}')
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Проверяем таблицы
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = cursor.fetchall()
        
        print(f'\n📊 Таблицы ({len(tables)}):')
        for table in tables:
            cursor.execute(f'SELECT COUNT(*) FROM {table[0]}')
            count = cursor.fetchone()[0]
            print(f'   - {table[0]}: {count:,} записей')
        
        # Проверяем представления
        cursor.execute("SELECT name FROM sqlite_master WHERE type='view' ORDER BY name")
        views = cursor.fetchall()
        
        print(f'\n👁️  Представления ({len(views)}):')
        for view in views:
            print(f'   - {view[0]}')
        
        conn.close()
        return True
        
    except Exception as e:
        print(f'❌ Ошибка при проверке БД: {e}')
        return False

def full_initialization():
    """Полная инициализация"""
    print('\n' + '🚀' * 40)
    print('ПОЛНАЯ ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ')
    print('🚀' * 40)
    
    # 1. Создание схемы
    print('\n1️⃣  СОЗДАНИЕ СХЕМЫ БД')
    print('-' * 80)
    create_database_schema()
    
    # 2. Импорт заявок
    print('\n2️⃣  ИМПОРТ ЗАЯВОК ИЗ EXCEL')
    print('-' * 80)
    try:
        app_imported, app_skipped = import_applications_from_excel()
        print(f'✅ Импортировано: {app_imported}, Пропущено: {app_skipped}')
    except Exception as e:
        print(f'❌ Ошибка импорта заявок: {e}')
    
    # 3. Импорт фиксаций
    print('\n3️⃣  ИМПОРТ ФИКСАЦИЙ ИЗ CSV')
    print('-' * 80)
    try:
        fix_imported, fix_skipped = import_fixations_from_csv()
        print(f'✅ Импортировано: {fix_imported}, Пропущено: {fix_skipped}')
    except Exception as e:
        print(f'❌ Ошибка импорта фиксаций: {e}')
    
    # 4. Создание отчетов
    print('\n4️⃣  СОЗДАНИЕ ОТЧЕТОВ')
    print('-' * 80)
    try:
        reports = create_all_reports()
        print(f'✅ Создано отчетов: {len(reports)}')
    except Exception as e:
        print(f'❌ Ошибка создания отчетов: {e}')
    
    print('\n' + '✅' * 40)
    print('ИНИЦИАЛИЗАЦИЯ ЗАВЕРШЕНА!')
    print('✅' * 40)

def main():
    """Главная функция"""
    while True:
        show_menu()
        
        try:
            choice = input('\n👉 Ваш выбор: ').strip()
            
            if choice == '0':
                print('\n👋 До свидания!')
                break
            
            elif choice == '1':
                print('\n' + '=' * 80)
                create_database_schema()
                print('=' * 80)
                input('\n✅ Нажмите Enter для продолжения...')
            
            elif choice == '2':
                print('\n' + '=' * 80)
                app_imported, app_skipped = import_applications_from_excel()
                print(f'\n✅ Импортировано: {app_imported}')
                print(f'⚠️  Пропущено: {app_skipped}')
                print('=' * 80)
                input('\n✅ Нажмите Enter для продолжения...')
            
            elif choice == '3':
                print('\n' + '=' * 80)
                fix_imported, fix_skipped = import_fixations_from_csv()
                print(f'\n✅ Импортировано: {fix_imported}')
                print(f'⚠️  Пропущено: {fix_skipped}')
                print('=' * 80)
                input('\n✅ Нажмите Enter для продолжения...')
            
            elif choice == '4':
                print('\n' + '=' * 80)
                reports = create_all_reports()
                print(f'\n✅ Создано отчетов: {len(reports)}')
                print('=' * 80)
                input('\n✅ Нажмите Enter для продолжения...')
            
            elif choice == '5':
                full_initialization()
                input('\n✅ Нажмите Enter для продолжения...')
            
            elif choice == '6':
                print('\n' + '=' * 80)
                check_database()
                print('=' * 80)
                input('\n✅ Нажмите Enter для продолжения...')
            
            else:
                print('\n❌ Неверный выбор. Попробуйте снова.')
                input('\nНажмите Enter для продолжения...')
        
        except KeyboardInterrupt:
            print('\n\n👋 Работа прервана пользователем')
            break
        except Exception as e:
            print(f'\n❌ Ошибка: {e}')
            input('\nНажмите Enter для продолжения...')

if __name__ == '__main__':
    main()
