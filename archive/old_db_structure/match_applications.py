"""
=============================================================================
СОПОСТАВЛЕНИЕ ЗАЯВОК С ДАННЫМИ ФИКСАЦИИ
=============================================================================
Сопоставляет заявки с данными ФИКСА по номеру заявки и телефону
=============================================================================
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
OUTPUT_DIR = BASE_DIR / 'output' / 'reports'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def match_applications_with_fiksa():
    """Сопоставление заявок с данными фиксации"""
    conn = get_db_connection()
    
    print("\n" + "=" * 80)
    print("🔍 СОПОСТАВЛЕНИЕ ЗАЯВОК С ДАННЫМИ ФИКСАЦИИ")
    print("=" * 80)
    
    # Получаем все заявки
    applications = pd.read_sql_query("""
        SELECT 
            a.id,
            a.application_number,
            a.phone,
            a.address,
            a.import_date,
            a.notes
        FROM applications a
        WHERE a.application_number IS NOT NULL AND a.application_number != ''
        ORDER BY a.import_date DESC
    """, conn)
    
    if applications.empty:
        print("⚠️  Нет загруженных заявок")
        return
    
    print(f"📋 Найдено заявок: {len(applications)}")
    
    # Сопоставляем с данными фиксации
    results = []
    matched_count = 0
    
    print("\n🔄 Обработка...")
    for idx, app in applications.iterrows():
        if idx % 1000 == 0 and idx > 0:
            print(f"  Обработано: {idx}/{len(applications)}... (найдено: {matched_count})")
            
        app_number = app['application_number']
        phone = app['phone'].replace('+998', '') if app['phone'] else ''
        
        # Ищем по номеру заявки (в колонке full_name)
        fiksa_data = pd.read_sql_query("""
            SELECT 
                operator_name,
                card_number,
                full_name,
                phone,
                status,
                call_date,
                notes
            FROM fiksa_records
            WHERE full_name = ? OR phone LIKE ?
            LIMIT 1
        """, conn, params=(app_number, f'%{phone}%'))
        
        if not fiksa_data.empty:
            matched_count += 1
            results.append({
                'Номер заявки': app_number,
                'Телефон из заявки': app['phone'],
                'Адрес из заявки': app['address'],
                'Дата из заявки': app['notes'],
                'Оператор': fiksa_data['operator_name'].iloc[0],
                'Телефон из ФИКСА': fiksa_data['phone'].iloc[0],
                'Статус': fiksa_data['status'].iloc[0],
                'Дата звонка': fiksa_data['call_date'].iloc[0],
                'Примечания ФИКСА': fiksa_data['notes'].iloc[0],
                'Найдено': '✅ ДА'
            })
        else:
            results.append({
                'Номер заявки': app_number,
                'Телефон из заявки': app['phone'],
                'Адрес из заявки': app['address'],
                'Дата из заявки': app['notes'],
                'Оператор': '',
                'Телефон из ФИКСА': '',
                'Статус': '',
                'Дата звонка': '',
                'Примечания ФИКСА': '',
                'Найдено': '❌ НЕТ'
            })
    
    # Создаем DataFrame
    df_results = pd.DataFrame(results)
    
    # Статистика
    not_matched = len(results) - matched_count
    
    print(f"\n📊 СТАТИСТИКА:")
    print(f"  ✅ Найдено в ФИКСА: {matched_count} ({matched_count/len(results)*100:.1f}%)")
    print(f"  ❌ Не найдено: {not_matched} ({not_matched/len(results)*100:.1f}%)")
    
    # Сохраняем отчет
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_path = OUTPUT_DIR / f"match_report_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_results.to_excel(writer, sheet_name='Сопоставление', index=False)
        
        # Лист со статистикой по статусам
        if matched_count > 0:
            found_df = df_results[df_results['Найдено'] == '✅ ДА']
            status_stats = found_df['Статус'].value_counts().reset_index()
            status_stats.columns = ['Статус', 'Количество']
            status_stats.to_excel(writer, sheet_name='Статистика по статусам', index=False)
            
            # Статистика по операторам
            operator_stats = found_df['Оператор'].value_counts().reset_index()
            operator_stats.columns = ['Оператор', 'Количество']
            operator_stats.to_excel(writer, sheet_name='По операторам', index=False)
    
    print(f"\n✅ Отчет сохранен: {excel_path}")
    
    conn.close()
    return df_results

if __name__ == "__main__":
    match_applications_with_fiksa()

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
