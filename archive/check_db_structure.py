import sqlite3

conn = sqlite3.connect('data/fiksa_database.db')
c = conn.cursor()

print("\n" + "=" * 100)
print("ПРОВЕРКА ДАННЫХ В ТАБЛИЦЕ fiksa_records")
print("=" * 100)

c.execute('SELECT card_number, full_name, phone FROM fiksa_records WHERE full_name != "" LIMIT 10')
rows = c.fetchall()

print(f"\n{'card_number':<30} | {'full_name':<40} | {'phone':<20}")
print("-" * 100)
for r in rows:
    print(f"{r[0]:<30} | {r[1]:<40} | {r[2]:<20}")

print("\n" + "=" * 100)
print("ПРОВЕРКА ДАННЫХ В ТАБЛИЦЕ applications")
print("=" * 100)

c.execute('SELECT application_number, phone FROM applications LIMIT 10')
rows = c.fetchall()

print(f"\n{'application_number':<40} | {'phone':<20}")
print("-" * 100)
for r in rows:
    print(f"{r[0]:<40} | {r[1]:<20}")

conn.close()
