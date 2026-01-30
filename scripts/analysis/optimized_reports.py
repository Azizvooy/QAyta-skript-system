# -*- coding: utf-8 -*-
import sqlite3, pandas as pd
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent.parent.parent
DB = BASE_DIR / 'data' / 'fiksa_database.db'
OUT = BASE_DIR / 'output' / 'reports'

conn = sqlite3.connect(DB)

# Дашборд
df1 = pd.read_sql_query('''
    SELECT COUNT(*) t, COUNT(DISTINCT operator_name) ops,
    ROUND(COUNT(CASE WHEN status LIKE '%Положительн%' THEN 1 END)*100.0/COUNT(*),1) sat
    FROM fiksa_records WHERE status IS NOT NULL
''', conn)
OUT.mkdir(parents=True, exist_ok=True)
df1.to_excel(OUT / f'ДАШБОРД_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx', index=False)
print(f'[OK] Дашборд: {df1.iloc[0,0]} звонков, {df1.iloc[0,2]}% удовл.')

# Операторы
df2 = pd.read_sql_query('''
    SELECT operator_name "Оператор", COUNT(*) "Всего",
    COUNT(CASE WHEN status LIKE '%Положительн%' THEN 1 END) "Положит.",
    COUNT(CASE WHEN status LIKE '%Отрицательн%' THEN 1 END) "Отрицат.",
    ROUND(COUNT(CASE WHEN status LIKE '%Положительн%' THEN 1 END)*100.0/COUNT(*),1) "Рейтинг"
    FROM fiksa_records WHERE status IS NOT NULL GROUP BY operator_name ORDER BY COUNT(*) DESC
''', conn)
df2.to_excel(OUT / f'ОПЕРАТОРЫ_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx', index=False)
print(f'[OK] Операторы: {len(df2)} операторов')

# Фидбэки
df3 = pd.read_sql_query('''
    SELECT ch.service_name "Служба", COUNT(*) "Всего",
    COUNT(CASE WHEN fr.status LIKE '%Положительн%' THEN 1 END) "Положит."
    FROM call_history_112 ch LEFT JOIN fiksa_records fr ON ch.card_number = fr.card_number
    WHERE ch.service_name IS NOT NULL GROUP BY ch.service_name ORDER BY COUNT(*) DESC
''', conn)
df3.to_excel(OUT / f'ФИДБЭКИ_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx', index=False)
print(f'[OK] Фидбэки: {len(df3)} служб')

conn.close()
print('[OK] Все отчеты созданы!')
