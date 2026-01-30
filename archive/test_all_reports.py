"""
Тестирование всех отчетов по отдельности
Проверяет каждый лист и показывает детали
"""
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime
import openpyxl

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
REPORTS_DIR = BASE_DIR / 'reports'

def print_header(text):
    """Красивый заголовок"""
    print('\n' + '=' * 80)
    print(f'  {text}')
    print('=' * 80)

def check_database():
    """Проверка базы данных"""
    print_header('ПРОВЕРКА БАЗЫ ДАННЫХ')
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Список таблиц
    tables = {
        'fiksa_records': 'Записи FIKSA',
        'call_history_112': 'История звонков 112',
        'applications': 'Заявки',
        'operator_stats_daily': 'Статистика операторов (ежедневная)',
        'service_feedback': 'Фидбэки служб'
    }
    
    for table, description in tables.items():
        try:
            c.execute(f'SELECT COUNT(*) FROM {table}')
            count = c.fetchone()[0]
            print(f'✅ {description:45} {count:,} записей')
        except sqlite3.OperationalError as e:
            print(f'❌ {description:45} ОШИБКА: {e}')
    
    conn.close()
    print()

def test_operator_stats():
    """Тест 1: Статистика операторов"""
    print_header('ТЕСТ 1: СТАТИСТИКА ОПЕРАТОРОВ')
    
    conn = sqlite3.connect(DB_PATH)
    
    # Проверяем данные
    query = '''
        SELECT 
            operator_name,
            COUNT(*) as total,
            COUNT(CASE WHEN status LIKE '%Положительн%' OR status LIKE '%положительн%' THEN 1 END) as positive,
            COUNT(CASE WHEN status LIKE '%Отрицательн%' OR status LIKE '%отрицательн%' THEN 1 END) as negative
        FROM fiksa_records
        WHERE status IS NOT NULL AND status != ''
        GROUP BY operator_name
        ORDER BY total DESC
        LIMIT 10
    '''
    
    df = pd.read_sql_query(query, conn)
    
    if df.empty:
        print('❌ НЕТ ДАННЫХ для статистики операторов')
        conn.close()
        return False
    
    print(f'✅ Найдено операторов: {len(df)}')
    print(f'\nТОП-5 операторов:')
    print('-' * 80)
    
    for idx, row in df.head().iterrows():
        percent = (row['positive'] / row['total'] * 100) if row['total'] > 0 else 0
        print(f"{row['operator_name']:40} Всего: {row['total']:4}  Положит: {row['positive']:4}  ({percent:.1f}%)")
    
    conn.close()
    
    # Пробуем создать Excel
    print('\nПопытка создать Excel...')
    try:
        output_file = REPORTS_DIR / 'analytics' / 'test_operators.xlsx'
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        conn = sqlite3.connect(DB_PATH)
        query_full = '''
            SELECT 
                operator_name as "Оператор",
                COUNT(*) as "Всего",
                COUNT(CASE WHEN status LIKE '%Положительн%' OR status LIKE '%положительн%' THEN 1 END) as "Положительные",
                COUNT(CASE WHEN status LIKE '%Отрицательн%' OR status LIKE '%отрицательн%' THEN 1 END) as "Отрицательные",
                COUNT(CASE WHEN status LIKE '%Нет ответа%' OR status LIKE '%занято%' THEN 1 END) as "Нет ответа",
                ROUND(COUNT(CASE WHEN status LIKE '%Положительн%' OR status LIKE '%положительн%' THEN 1 END) * 100.0 / COUNT(*), 1) as "% Успешных"
            FROM fiksa_records
            WHERE status IS NOT NULL AND status != ''
            GROUP BY operator_name
            ORDER BY COUNT(*) DESC
        '''
        
        df_full = pd.read_sql_query(query_full, conn)
        df_full.to_excel(output_file, index=False, sheet_name='Статистика')
        conn.close()
        
        print(f'✅ Excel создан: {output_file}')
        print(f'   Строк: {len(df_full)}, Колонок: {len(df_full.columns)}')
        return True
        
    except Exception as e:
        print(f'❌ ОШИБКА при создании Excel: {e}')
        return False

def test_service_feedback():
    """Тест 2: Фидбэки служб"""
    print_header('ТЕСТ 2: ФИДБЭКИ СЛУЖБ')
    
    conn = sqlite3.connect(DB_PATH)
    
    # Проверяем данные - группируем по статусам
    query = '''
        SELECT 
            status,
            COUNT(*) as total,
            COUNT(CASE WHEN status LIKE '%не%' OR status LIKE '%отказ%' OR status LIKE '%Отрицательн%' THEN 1 END) as problems
        FROM fiksa_records
        WHERE status IS NOT NULL AND status != ''
        GROUP BY status
        ORDER BY total DESC
        LIMIT 10
    '''
    
    df = pd.read_sql_query(query, conn)
    
    if df.empty:
        print('❌ НЕТ ДАННЫХ по статусам')
        conn.close()
        return False
    
    print(f'✅ Типов статусов: {len(df)}')
    print('\nСтатистика по статусам (ТОП-10):')
    print('-' * 80)
    
    for _, row in df.iterrows():
        percent = (row['problems'] / row['total'] * 100) if row['total'] > 0 else 0
        print(f"{row['status'][:50]:50}  Всего: {row['total']:5}  Проблемных: {row['problems']:5}  ({percent:.1f}%)")
    
    # Пробуем создать Excel с 3 листами
    print('\nПопытка создать Excel с 3 листами...')
    try:
        output_file = REPORTS_DIR / 'analytics' / 'test_feedback.xlsx'
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Лист 1: Общая статистика по статусам
            df.to_excel(writer, sheet_name='Общая статистика', index=False)
            print('✅ Лист 1: Общая статистика - создан')
            
            # Лист 2: Детализация по операторам
            detail_query = '''
                SELECT 
                    operator_name as "Оператор",
                    status as "Статус",
                    COUNT(*) as "Количество"
                FROM fiksa_records
                WHERE status IS NOT NULL AND status != ''
                GROUP BY operator_name, status
                ORDER BY operator_name, COUNT(*) DESC
            '''
            df_detail = pd.read_sql_query(detail_query, conn)
            df_detail.to_excel(writer, sheet_name='Детализация', index=False)
            print(f'✅ Лист 2: Детализация - создан ({len(df_detail)} строк)')
            
            # Лист 3: Проблемные случаи
            problems_query = '''
                SELECT 
                    call_date as "Дата",
                    operator_name as "Оператор",
                    full_name as "Номер карточки",
                    status as "Статус",
                    phone as "Телефон",
                    notes as "Примечания"
                FROM fiksa_records
                WHERE status LIKE '%не%' OR status LIKE '%отказ%' OR status LIKE '%Отрицательн%'
                ORDER BY call_date DESC
                LIMIT 500
            '''
            df_problems = pd.read_sql_query(problems_query, conn)
            df_problems.to_excel(writer, sheet_name='Проблемные (топ-500)', index=False)
            print(f'✅ Лист 3: Проблемные - создан ({len(df_problems)} строк)')
        
        conn.close()
        print(f'\n✅ Excel создан: {output_file}')
        return True
        
    except Exception as e:
        print(f'❌ ОШИБКА при создании Excel: {e}')
        conn.close()
        return False

def test_citizen_responses():
    """Тест 3: Ответы граждан"""
    print_header('ТЕСТ 3: ОТВЕТЫ ГРАЖДАН')
    
    conn = sqlite3.connect(DB_PATH)
    
    # Проверяем типы ответов
    query = '''
        SELECT 
            status as "Тип ответа",
            COUNT(*) as "Количество",
            ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM fiksa_records WHERE status IS NOT NULL), 1) as "Процент"
        FROM fiksa_records
        WHERE status IS NOT NULL AND status != ''
        GROUP BY status
        ORDER BY COUNT(*) DESC
        LIMIT 15
    '''
    
    df = pd.read_sql_query(query, conn)
    
    if df.empty:
        print('❌ НЕТ ДАННЫХ по ответам')
        conn.close()
        return False
    
    print(f'✅ Типов ответов: {len(df)}')
    print('\nТОП-10 типов ответов:')
    print('-' * 80)
    
    for _, row in df.head(10).iterrows():
        print(f"{row['Тип ответа']:50} {row['Количество']:5} ({row['Процент']:5.1f}%)")
    
    # Пробуем создать Excel
    print('\nПопытка создать Excel...')
    try:
        output_file = REPORTS_DIR / 'analytics' / 'test_responses.xlsx'
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        df.to_excel(output_file, index=False, sheet_name='Ответы граждан')
        conn.close()
        
        print(f'✅ Excel создан: {output_file}')
        print(f'   Строк: {len(df)}')
        return True
        
    except Exception as e:
        print(f'❌ ОШИБКА при создании Excel: {e}')
        conn.close()
        return False

def test_service_report():
    """Тест 4: Отчет по конкретной службе"""
    print_header('ТЕСТ 4: ОТЧЕТ ПО СЛУЖБЕ 102 (ПРИМЕР)')
    
    conn = sqlite3.connect(DB_PATH)
    
    # Проверяем данные по службе 102
    query = '''
        SELECT 
            ch.call_date,
            ch.call_time,
            ch.incident_number,
            ch.caller_phone,
            ch.address,
            f.operator_name as fiksa_operator,
            f.call_date as fiksa_date,
            f.phone as fiksa_phone,
            f.status as fiksa_status
        FROM call_history_112 ch
        LEFT JOIN fiksa_records f ON f.full_name = ch.incident_number
        WHERE ch.service_code = '102'
        ORDER BY ch.call_date DESC, ch.call_time DESC
        LIMIT 100
    '''
    
    df = pd.read_sql_query(query, conn)
    
    if df.empty:
        print('❌ НЕТ ДАННЫХ по службе 102')
        conn.close()
        return False
    
    # Считаем связи с FIKSA
    with_fiksa = df['fiksa_operator'].notna().sum()
    total = len(df)
    percent = (with_fiksa / total * 100) if total > 0 else 0
    
    print(f'✅ Записей по службе 102: {total}')
    print(f'✅ Связано с FIKSA: {with_fiksa} ({percent:.1f}%)')
    
    print('\nПример данных (первые 5 записей):')
    print('-' * 80)
    
    for idx, row in df.head(5).iterrows():
        print(f"\nДата: {row['call_date']} {row['call_time']}")
        print(f"  Инцидент: {row['incident_number']}")
        print(f"  Телефон: {row['caller_phone']}")
        print(f"  Адрес: {row['address'][:50] if row['address'] else 'Нет'}...")
        if pd.notna(row['fiksa_operator']):
            print(f"  [FIKSA] Оператор: {row['fiksa_operator']}, Статус: {row['fiksa_status']}")
        else:
            print(f"  [FIKSA] Нет данных")
    
    # Пробуем создать Excel
    print('\nПопытка создать Excel...')
    try:
        output_file = REPORTS_DIR / 'services' / 'test_service_102.xlsx'
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Полный запрос для всех записей
        query_full = '''
            SELECT 
                ch.call_date as "Дата",
                ch.call_time as "Время",
                ch.incident_number as "Номер инцидента",
                ch.caller_phone as "Телефон звонившего",
                ch.address as "Адрес",
                ch.status as "Статус",
                f.operator_name as "Оператор FIKSA",
                f.call_date as "Дата звонка FIKSA",
                f.phone as "Телефон FIKSA",
                f.status as "Статус FIKSA",
                f.notes as "Примечания FIKSA"
            FROM call_history_112 ch
            LEFT JOIN fiksa_records f ON f.full_name = ch.incident_number
            WHERE ch.service_code = '102'
            ORDER BY ch.call_date DESC, ch.call_time DESC
        '''
        
        df_full = pd.read_sql_query(query_full, conn)
        df_full.to_excel(output_file, index=False, sheet_name='Служба 102')
        conn.close()
        
        print(f'✅ Excel создан: {output_file}')
        print(f'   Всего записей: {len(df_full)}')
        print(f'   Со связью FIKSA: {df_full["Оператор FIKSA"].notna().sum()}')
        return True
        
    except Exception as e:
        print(f'❌ ОШИБКА при создании Excel: {e}')
        conn.close()
        return False

def main():
    """Запуск всех тестов"""
    print('\n' + '=' * 80)
    print('  ТЕСТИРОВАНИЕ ВСЕХ ОТЧЕТОВ')
    print('  ' + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    print('=' * 80)
    
    # Проверка базы
    check_database()
    
    # Запуск тестов
    results = {}
    
    results['Статистика операторов'] = test_operator_stats()
    results['Фидбэки служб'] = test_service_feedback()
    results['Ответы граждан'] = test_citizen_responses()
    results['Отчет по службе'] = test_service_report()
    
    # Итоги
    print_header('ИТОГИ ТЕСТИРОВАНИЯ')
    
    for test_name, result in results.items():
        status = '✅ УСПЕШНО' if result else '❌ ОШИБКА'
        print(f'{test_name:30} {status}')
    
    total = len(results)
    passed = sum(results.values())
    
    print(f'\nВСЕГО ТЕСТОВ: {total}')
    print(f'ПРОЙДЕНО: {passed}')
    print(f'ОШИБОК: {total - passed}')
    
    if passed == total:
        print('\n🎉 ВСЕ ТЕСТЫ ПРОЙДЕНЫ!')
    else:
        print('\n⚠️ ЕСТЬ ПРОБЛЕМЫ - СМОТРИТЕ ДЕТАЛИ ВЫШЕ')
    
    print('\nТестовые файлы созданы в:')
    print(f'  {REPORTS_DIR / "analytics"}')
    print(f'  {REPORTS_DIR / "services"}')
    print('\n' + '=' * 80)

if __name__ == '__main__':
    main()
