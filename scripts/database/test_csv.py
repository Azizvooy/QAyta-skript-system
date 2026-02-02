import pandas as pd

df = pd.read_csv('exported_sheets/Narziyeva Gavxar Atxamjanovna/Narziyeva Gavxar Atxamjanovna 12.2025.csv')
print('Total rows:', len(df))
print('\nColumns:', list(df.columns))
print('\nFirst 3 rows:')
for i in range(min(3, len(df))):
    print(f'\n--- Row {i} ---')
    for col in df.columns:
        val = df.iloc[i][col]
        if pd.notna(val):
            print(f'  {col:25s}: {val}')
