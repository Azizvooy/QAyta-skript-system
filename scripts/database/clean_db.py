"""
Очистка базы данных от лишних строк
Удаляет операторов и связанные данные по паттернам исключения
"""
import sqlite3
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
LOG_DIR = BASE_DIR / 'logs' / 'database'

# Паттерны для удаления
EXCLUDE_PATTERNS = [
    'Тренды',
    'Текущий месяц - Сводка',
    'Предыдущий месяц - Сводка',
    'СВОДКА СОТРУДНИКИ',
    'Ноябрь 2025',
    'Декабрь 2025',
    'сводка',
    'итого',
    'total',
    'summary',
]

def log_message(message):
    """Логирование"""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f'cleanup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
    
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'[{timestamp}] {message}\n')
    
    print(message)

def clean_database():
    """Очистка БД от лишних записей"""
    print('=' * 80)
    print('ОЧИСТКА БАЗЫ ДАННЫХ ОТ ЛИШНИХ ЗАПИСЕЙ')
    print('=' * 80)
    print()
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        # Проверяем, есть ли таблица operators
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='operators'")
        if not cursor.fetchone():
            print('⚠️  Таблица operators не найдена. Возможно используется старая структура БД.')
            return
        
        # Находим операторов для удаления
        log_message('Поиск операторов для удаления...')
        
        operators_to_delete = []
        for pattern in EXCLUDE_PATTERNS:
            cursor.execute('''
                SELECT operator_id, operator_name 
                FROM operators 
                WHERE LOWER(operator_name) LIKE ?
            ''', (f'%{pattern.lower()}%',))
            
            found = cursor.fetchall()
            operators_to_delete.extend(found)
        
        if not operators_to_delete:
            log_message('✅ Не найдено операторов для удаления')
            return
        
        # Удаляем дубликаты
        operators_to_delete = list(set(operators_to_delete))
        
        log_message(f'\nНайдено операторов для удаления: {len(operators_to_delete)}')
        for op_id, op_name in operators_to_delete:
            log_message(f'  • ID {op_id}: {op_name}')
        
        print()
        confirm = input('❓ Удалить этих операторов и все их данные? (yes/no): ')
        if confirm.lower() not in ['yes', 'y', 'да']:
            log_message('❌ Отменено пользователем')
            return
        
        # Удаляем данные
        deleted_fixations = 0
        deleted_stats = 0
        
        for op_id, op_name in operators_to_delete:
            # Удаляем фиксации
            cursor.execute('DELETE FROM fixations WHERE operator_id = ?', (op_id,))
            deleted_fixations += cursor.rowcount
            
            # Удаляем статистику
            cursor.execute('DELETE FROM daily_statistics WHERE operator_id = ?', (op_id,))
            deleted_stats += cursor.rowcount
            
            # Удаляем самого оператора
            cursor.execute('DELETE FROM operators WHERE operator_id = ?', (op_id,))
            
            log_message(f'✓ Удален: {op_name}')
        
        conn.commit()
        
        print()
        log_message('=' * 80)
        log_message('ИТОГИ ОЧИСТКИ:')
        log_message('=' * 80)
        log_message(f'Удалено операторов: {len(operators_to_delete)}')
        log_message(f'Удалено фиксаций: {deleted_fixations}')
        log_message(f'Удалено записей статистики: {deleted_stats}')
        log_message('=' * 80)
        log_message('✅ Очистка завершена успешно')
        
    except Exception as e:
        conn.rollback()
        log_message(f'❌ ОШИБКА: {e}')
        raise
    
    finally:
        conn.close()

if __name__ == '__main__':
    clean_database()
