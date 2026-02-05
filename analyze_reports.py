import pandas as pd
from pathlib import Path

report_dir = Path('reports/службы_детально')

# Анализируем каждый файл службы
for excel_file in sorted(report_dir.glob('СЛУЖБА_*.xlsx')):
    print(f'\n{"="*80}')
    print(f'📊 {excel_file.name}')
    print("="*80)
    
    try:
        # Читаем все листы
        xl = pd.ExcelFile(excel_file)
        print(f'\n📋 Листы в файле ({len(xl.sheet_names)}):')
        
        for sheet in xl.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            print(f'  • {sheet:<30} - {len(df):>7,} строк x {len(df.columns):>2} колонок')
            
            # Показываем первые колонки
            if len(df) > 0:
                cols = list(df.columns)[:5]
                print(f'    Колонки: {", ".join(str(c) for c in cols)}...')
        
    except Exception as e:
        print(f'  ❌ Ошибка: {e}')

# Анализируем общий отчет
general_file = report_dir / 'ОБЩИЙ_ОТЧЕТ_ВСЕ_СЛУЖБЫ_2026-02-04_07-39-29.xlsx'
if general_file.exists():
    print(f'\n{"="*80}')
    print(f'📊 ОБЩИЙ ОТЧЕТ')
    print("="*80)
    
    try:
        xl = pd.ExcelFile(general_file)
        print(f'\n📋 Листы в файле ({len(xl.sheet_names)}):')
        
        for sheet in xl.sheet_names:
            df = pd.read_excel(general_file, sheet_name=sheet, nrows=5)
            print(f'  • {sheet:<30}')
            if len(df) > 0:
                print(f'    Колонки: {", ".join(str(c) for c in df.columns[:5])}...')
    except Exception as e:
        print(f'  ❌ Ошибка: {e}')
