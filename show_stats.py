import sqlite3

c = sqlite3.connect('data/fiksa_database.db')
print(f'[БД] fiksa_records: {c.execute("SELECT COUNT(*) FROM fiksa_records").fetchone()[0]:,} записей')
print(f'[БД] call_history_112: {c.execute("SELECT COUNT(*) FROM call_history_112").fetchone()[0]:,} записей')
c.close()
