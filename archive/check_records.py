import sqlite3

conn = sqlite3.connect('data/fiksa_database.db')
cursor = conn.cursor()

# Получаем операторов
cursor.execute('SELECT operator_name FROM fiksa_records GROUP BY operator_name ORDER BY operator_name')
operators = cursor.fetchall()

print("\n" + "=" * 80)
print("📊 КОЛИЧЕСТВО ЗАПИСЕЙ ПО ОПЕРАТОРАМ")
print("=" * 80)

for i, (op_name,) in enumerate(operators, 1):
    cursor.execute('SELECT COUNT(*) FROM fiksa_records WHERE operator_name = ?', (op_name,))
    count = cursor.fetchone()[0]
    
    # Выделяем операторов с большим количеством
    if count >= 900:
        print(f"[{i:2d}] {op_name:50s}: {count:4d} записей ⚠️  МНОГО")
    else:
        print(f"[{i:2d}] {op_name:50s}: {count:4d} записей")

# Проверяем операторов #15 и #18 детально
print("\n" + "=" * 80)
print("🔍 ДЕТАЛЬНАЯ ПРОВЕРКА")
print("=" * 80)

# Находим операторов с 999 и 953 записями
cursor.execute('''
    SELECT operator_name, COUNT(*) as cnt 
    FROM fiksa_records 
    GROUP BY operator_name 
    HAVING cnt >= 900
    ORDER BY cnt DESC
''')

for op_name, count in cursor.fetchall():
    print(f"\n{op_name}: {count} записей")
    
    # Проверяем пустые строки
    cursor.execute('''
        SELECT COUNT(*) FROM fiksa_records 
        WHERE operator_name = ? AND (card_number = '' OR card_number IS NULL) AND (full_name = '' OR full_name IS NULL)
    ''', (op_name,))
    empty = cursor.fetchone()[0]
    
    # Проверяем дубликаты
    cursor.execute('''
        SELECT card_number, COUNT(*) as cnt 
        FROM fiksa_records 
        WHERE operator_name = ? AND card_number != ''
        GROUP BY card_number 
        HAVING cnt > 1 
        LIMIT 5
    ''', (op_name,))
    duplicates = cursor.fetchall()
    
    print(f"  Пустых строк: {empty}")
    if duplicates:
        print(f"  Дубликаты карт:")
        for card, cnt in duplicates:
            print(f"    - {card}: {cnt} раз")

conn.close()
print("\n" + "=" * 80)
