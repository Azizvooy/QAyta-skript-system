import sqlite3

conn = sqlite3.connect('data/fiksa_database.db')
c = conn.cursor()

total = c.execute('SELECT COUNT(*) FROM fiksa_records').fetchone()[0]
ops = c.execute('SELECT COUNT(DISTINCT operator_name) FROM fiksa_records').fetchone()[0]
phones = c.execute('SELECT COUNT(*) FROM fiksa_records WHERE phone IS NOT NULL AND phone != ""').fetchone()[0]
positive = c.execute('SELECT COUNT(*) FROM fiksa_records WHERE status LIKE "%Положительн%"').fetchone()[0]

print('\n' + '='*80)
print('📊 СТАТИСТИКА ОБНОВЛЕННОЙ БАЗЫ ДАННЫХ')
print('='*80)
print(f'\n📈 Общие показатели:')
print(f'  Всего записей: {total:,}')
print(f'  Операторов: {ops}')
print(f'  С телефонами: {phones:,} ({phones*100/total:.1f}%)')
print(f'  Положительных: {positive:,} ({positive*100/total:.1f}%)')

# Даты
dates = c.execute('SELECT MIN(call_date), MAX(call_date) FROM fiksa_records WHERE call_date IS NOT NULL').fetchone()
print(f'\n📅 Период данных:')
print(f'  С: {dates[0] if dates[0] else "н/д"}')
print(f'  По: {dates[1] if dates[1] else "н/д"}')

# Последние обновления
recent = c.execute('''
    SELECT operator_name, COUNT(*) as cnt 
    FROM fiksa_records 
    WHERE collection_date = date('now')
    GROUP BY operator_name 
    ORDER BY cnt DESC 
    LIMIT 5
''').fetchall()

if recent:
    print(f'\n🆕 Сегодня обновлено ({len(recent)} операторов):')
    for op, cnt in recent:
        print(f'  {op[:50]:50} - {cnt:,} записей')

print('\n' + '='*80)
conn.close()
