import pandas as pd

df = pd.read_csv('reports/СОПОСТАВЛЕНИЕ_ПО_ИНЦИДЕНТАМ_2026-01-27_11-42.csv')

print(f'Всего строк: {len(df):,}')
print(f'\nУникальных инцидентов из Sheets: {df["Инцидент_Sheets"].nunique():,}')
print(f'Уникальных инцидентов из 112: {df["Инцидент_112"].nunique():,}')

print(f'\nСлужб в данных:')
print(df["Служба_112"].value_counts())

print(f'\nПример одного инцидента:')
inc = df["Инцидент_Sheets"].iloc[0]
print(f'Инцидент: {inc}')
sample = df[df["Инцидент_Sheets"]==inc][["Инцидент_Sheets", "Служба_112", "Положительно", "Документ"]].head(10)
print(sample)

print(f'\n\nПроверка: сколько раз один инцидент повторяется?')
counts = df.groupby('Инцидент_Sheets').size()
print(f'Мин повторений: {counts.min()}')
print(f'Макс повторений: {counts.max()}')
print(f'Средне: {counts.mean():.1f}')

print(f'\nПример инцидента с максимум повторений:')
max_inc = counts.idxmax()
print(f'Инцидент {max_inc} встречается {counts.max()} раз')
sample2 = df[df["Инцидент_Sheets"]==max_inc][["Инцидент_Sheets", "Служба_112", "Документ"]].head(20)
print(sample2)
