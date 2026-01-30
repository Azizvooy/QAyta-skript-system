import sqlite3

conn = sqlite3.connect('data/fiksa_database.db')

print('ТАБЛИЦА fiksa_records:')
for row in conn.execute('PRAGMA table_info(fiksa_records)').fetchall():
    print(f'  {row[1]} ({row[2]})')

print('\nТАБЛИЦА call_history_112:')
for row in conn.execute('PRAGMA table_info(call_history_112)').fetchall():
    print(f'  {row[1]} ({row[2]})')

conn.close()
