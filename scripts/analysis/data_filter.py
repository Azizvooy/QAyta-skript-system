"""
=============================================================================
ФИЛЬТР И АНАЛИЗ ДАННЫХ ИЗ БАЗЫ ДАННЫХ
=============================================================================
Позволяет фильтровать и анализировать собранные данные
=============================================================================
"""

import sqlite3
import json
from datetime import datetime, timedelta
import pandas as pd

DB_PATH = 'data/fiksa_database.db'

# =============================================================================
# ФИЛЬТРЫ ДАННЫХ
# =============================================================================

def get_records_by_date(start_date=None, end_date=None):
    """Получить записи за период"""
    conn = sqlite3.connect(DB_PATH)
    
    if start_date and end_date:
        query = '''
            SELECT * FROM fiksa_records 
            WHERE collection_date BETWEEN ? AND ?
            ORDER BY collection_date DESC
        '''
        df = pd.read_sql_query(query, conn, params=(start_date, end_date))
    elif start_date:
        query = '''
            SELECT * FROM fiksa_records 
            WHERE collection_date >= ?
            ORDER BY collection_date DESC
        '''
        df = pd.read_sql_query(query, conn, params=(start_date,))
    else:
        df = pd.read_sql_query('SELECT * FROM fiksa_records ORDER BY collection_date DESC', conn)
    
    conn.close()
    return df

def get_records_by_operator(operator_name, start_date=None, end_date=None):
    """Получить записи по оператору"""
    conn = sqlite3.connect(DB_PATH)
    
    if start_date and end_date:
        query = '''
            SELECT * FROM fiksa_records 
            WHERE operator_name = ? AND collection_date BETWEEN ? AND ?
            ORDER BY collection_date DESC
        '''
        df = pd.read_sql_query(query, conn, params=(operator_name, start_date, end_date))
    else:
        query = '''
            SELECT * FROM fiksa_records 
            WHERE operator_name = ?
            ORDER BY collection_date DESC
        '''
        df = pd.read_sql_query(query, conn, params=(operator_name,))
    
    conn.close()
    return df

def get_records_by_status(status, start_date=None, end_date=None):
    """Получить записи по статусу"""
    conn = sqlite3.connect(DB_PATH)
    
    if start_date and end_date:
        query = '''
            SELECT * FROM fiksa_records 
            WHERE status = ? AND collection_date BETWEEN ? AND ?
            ORDER BY collection_date DESC
        '''
        df = pd.read_sql_query(query, conn, params=(status, start_date, end_date))
    else:
        query = '''
            SELECT * FROM fiksa_records 
            WHERE status = ?
            ORDER BY collection_date DESC
        '''
        df = pd.read_sql_query(query, conn, params=(status,))
    
    conn.close()
    return df

def search_by_card_or_name(search_text):
    """Поиск по номеру карты или имени"""
    conn = sqlite3.connect(DB_PATH)
    
    query = '''
        SELECT * FROM fiksa_records 
        WHERE card_number LIKE ? OR full_name LIKE ?
        ORDER BY collection_date DESC
    '''
    
    search_pattern = f'%{search_text}%'
    df = pd.read_sql_query(query, conn, params=(search_pattern, search_pattern))
    
    conn.close()
    return df

# =============================================================================
# СТАТИСТИКА
# =============================================================================

def get_daily_stats(date=None):
    """Статистика за день"""
    if date is None:
        date = datetime.now().date()
    
    conn = sqlite3.connect(DB_PATH)
    
    # Общая статистика
    query = '''
        SELECT 
            COUNT(*) as total_records,
            COUNT(DISTINCT operator_name) as total_operators,
            COUNT(DISTINCT card_number) as unique_cards
        FROM fiksa_records
        WHERE collection_date = ?
    '''
    
    cursor = conn.cursor()
    cursor.execute(query, (date,))
    result = cursor.fetchone()
    
    stats = {
        'date': date,
        'total_records': result[0],
        'total_operators': result[1],
        'unique_cards': result[2]
    }
    
    # Статистика по статусам
    query = '''
        SELECT status, COUNT(*) as count
        FROM fiksa_records
        WHERE collection_date = ?
        GROUP BY status
        ORDER BY count DESC
    '''
    
    cursor.execute(query, (date,))
    status_stats = {row[0]: row[1] for row in cursor.fetchall()}
    stats['status_breakdown'] = status_stats
    
    # Топ операторов
    query = '''
        SELECT operator_name, COUNT(*) as count
        FROM fiksa_records
        WHERE collection_date = ?
        GROUP BY operator_name
        ORDER BY count DESC
        LIMIT 10
    '''
    
    cursor.execute(query, (date,))
    top_operators = [(row[0], row[1]) for row in cursor.fetchall()]
    stats['top_operators'] = top_operators
    
    conn.close()
    return stats

def get_period_stats(start_date, end_date):
    """Статистика за период"""
    conn = sqlite3.connect(DB_PATH)
    
    # Общая статистика
    query = '''
        SELECT 
            COUNT(*) as total_records,
            COUNT(DISTINCT operator_name) as total_operators,
            COUNT(DISTINCT card_number) as unique_cards,
            COUNT(DISTINCT collection_date) as days_count
        FROM fiksa_records
        WHERE collection_date BETWEEN ? AND ?
    '''
    
    cursor = conn.cursor()
    cursor.execute(query, (start_date, end_date))
    result = cursor.fetchone()
    
    stats = {
        'period': f'{start_date} - {end_date}',
        'total_records': result[0],
        'total_operators': result[1],
        'unique_cards': result[2],
        'days_count': result[3]
    }
    
    # Динамика по дням
    query = '''
        SELECT collection_date, COUNT(*) as count
        FROM fiksa_records
        WHERE collection_date BETWEEN ? AND ?
        GROUP BY collection_date
        ORDER BY collection_date
    '''
    
    df = pd.read_sql_query(query, conn, params=(start_date, end_date))
    stats['daily_dynamics'] = df
    
    # Статистика по статусам
    query = '''
        SELECT status, COUNT(*) as count
        FROM fiksa_records
        WHERE collection_date BETWEEN ? AND ?
        GROUP BY status
        ORDER BY count DESC
    '''
    
    cursor.execute(query, (start_date, end_date))
    status_stats = {row[0]: row[1] for row in cursor.fetchall()}
    stats['status_breakdown'] = status_stats
    
    conn.close()
    return stats

def get_operator_performance(start_date=None, end_date=None):
    """Производительность операторов"""
    conn = sqlite3.connect(DB_PATH)
    
    if start_date and end_date:
        query = '''
            SELECT 
                operator_name,
                COUNT(*) as total_calls,
                COUNT(DISTINCT card_number) as unique_cards,
                COUNT(DISTINCT collection_date) as work_days,
                COUNT(*) * 1.0 / COUNT(DISTINCT collection_date) as avg_per_day
            FROM fiksa_records
            WHERE collection_date BETWEEN ? AND ?
            GROUP BY operator_name
            ORDER BY total_calls DESC
        '''
        df = pd.read_sql_query(query, conn, params=(start_date, end_date))
    else:
        query = '''
            SELECT 
                operator_name,
                COUNT(*) as total_calls,
                COUNT(DISTINCT card_number) as unique_cards,
                COUNT(DISTINCT collection_date) as work_days,
                COUNT(*) * 1.0 / COUNT(DISTINCT collection_date) as avg_per_day
            FROM fiksa_records
            GROUP BY operator_name
            ORDER BY total_calls DESC
        '''
        df = pd.read_sql_query(query, conn)
    
    conn.close()
    return df

# =============================================================================
# ИНТЕРФЕЙС КОМАНДНОЙ СТРОКИ
# =============================================================================

def main():
    print("=" * 80)
    print("🔍 ФИЛЬТР И АНАЛИЗ ДАННЫХ ФИКСАЦИИ")
    print("=" * 80)
    
    while True:
        print("\nВыберите действие:")
        print("1. Статистика за сегодня")
        print("2. Статистика за период")
        print("3. Поиск по карте/имени")
        print("4. Фильтр по оператору")
        print("5. Фильтр по статусу")
        print("6. Производительность операторов")
        print("7. Экспорт в Excel")
        print("0. Выход")
        
        choice = input("\nВведите номер: ").strip()
        
        if choice == '0':
            break
            
        elif choice == '1':
            stats = get_daily_stats()
            print(f"\n📊 Статистика за {stats['date']}:")
            print(f"   Всего записей: {stats['total_records']}")
            print(f"   Операторов: {stats['total_operators']}")
            print(f"   Уникальных карт: {stats['unique_cards']}")
            print("\n   По статусам:")
            for status, count in stats['status_breakdown'].items():
                print(f"      {status}: {count}")
            print("\n   Топ операторов:")
            for name, count in stats['top_operators']:
                print(f"      {name}: {count}")
        
        elif choice == '2':
            start = input("Дата начала (YYYY-MM-DD): ").strip()
            end = input("Дата окончания (YYYY-MM-DD): ").strip()
            
            stats = get_period_stats(start, end)
            print(f"\n📊 Статистика за {stats['period']}:")
            print(f"   Всего записей: {stats['total_records']}")
            print(f"   Операторов: {stats['total_operators']}")
            print(f"   Уникальных карт: {stats['unique_cards']}")
            print(f"   Дней: {stats['days_count']}")
            print("\n   По статусам:")
            for status, count in stats['status_breakdown'].items():
                print(f"      {status}: {count}")
        
        elif choice == '3':
            search = input("Введите номер карты или имя: ").strip()
            df = search_by_card_or_name(search)
            print(f"\n🔍 Найдено записей: {len(df)}")
            if len(df) > 0:
                print(df[['collection_date', 'operator_name', 'card_number', 'full_name', 'status']].to_string())
        
        elif choice == '4':
            operator = input("Имя оператора: ").strip()
            df = get_records_by_operator(operator)
            print(f"\n📋 Записей оператора {operator}: {len(df)}")
            if len(df) > 0:
                print(df[['collection_date', 'card_number', 'full_name', 'status']].head(20).to_string())
        
        elif choice == '5':
            status = input("Статус: ").strip()
            df = get_records_by_status(status)
            print(f"\n📋 Записей со статусом '{status}': {len(df)}")
            if len(df) > 0:
                print(df[['collection_date', 'operator_name', 'card_number', 'full_name']].head(20).to_string())
        
        elif choice == '6':
            df = get_operator_performance()
            print("\n📊 Производительность операторов:")
            print(df.to_string(index=False))
        
        elif choice == '7':
            start = input("Дата начала (YYYY-MM-DD, Enter для всех): ").strip()
            end = input("Дата окончания (YYYY-MM-DD, Enter для всех): ").strip()
            
            if start and end:
                df = get_records_by_date(start, end)
            else:
                df = get_records_by_date()
            
            filename = f"output/filtered_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(filename, index=False)
            print(f"\n✅ Экспортировано в {filename}")

if __name__ == "__main__":
    main()
