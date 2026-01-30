"""
Создание детальной отчётности по службам
С разделением на положительные, отрицательные и жалобы
"""
import pandas as pd
import glob
import os
from datetime import datetime

def load_latest_matching():
    """Загрузка последнего файла сопоставления"""
    files = glob.glob('reports/СОПОСТАВЛЕНИЕ_ПО_ИНЦИДЕНТАМ_*.csv')
    if not files:
        raise FileNotFoundError("Не найден файл сопоставления")
    
    latest = max(files, key=os.path.getctime)
    print(f"Загружаю: {latest}")
    
    df = pd.read_csv(latest, low_memory=False)
    print(f"Загружено записей: {len(df):,}")
    
    return df

def create_service_reports(df):
    """Создание отчётов по каждой службе"""
    
    # Получаем список служб
    services = sorted(df['Служба_112'].dropna().unique())
    print(f"\nНайдено служб: {services}")
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    
    # Создаём папку для отчётов
    output_dir = 'reports/службы_детально'
    os.makedirs(output_dir, exist_ok=True)
    
    for service in services:
        print(f"\n{'='*80}")
        print(f"СЛУЖБА {service}")
        print('='*80)
        
        # Фильтруем данные по службе
        df_service = df[df['Служба_112'] == service].copy()
        
        print(f"Всего записей: {len(df_service):,}")
        
        # Определяем положительные и отрицательные
        df_service['Категория'] = 'Не определено'
        
        # Положительные: есть "Положительно" или "положительный" в статусе
        mask_positive = (
            (df_service['Положительно'].astype(str).str.contains('Положительн', case=False, na=False)) |
            (df_service['Статус_Sheets'].astype(str).str.contains('Положительн', case=False, na=False))
        )
        df_service.loc[mask_positive, 'Категория'] = 'Положительно'
        
        # Отрицательные/Жалобы: есть жалоба или "отрицательный" в статусе
        mask_negative = (
            (df_service['Есть_жалоба'] == True) |
            (df_service['Жалоба'].notna()) |
            (df_service['Статус_Sheets'].astype(str).str.contains('Отрицательн', case=False, na=False))
        )
        df_service.loc[mask_negative, 'Категория'] = 'Отрицательно'
        
        # Статистика
        positive_count = (df_service['Категория'] == 'Положительно').sum()
        negative_count = (df_service['Категория'] == 'Отрицательно').sum()
        neutral_count = (df_service['Категория'] == 'Не определено').sum()
        
        print(f"Положительно: {positive_count:,} ({positive_count/len(df_service)*100:.1f}%)")
        print(f"Отрицательно/Жалобы: {negative_count:,} ({negative_count/len(df_service)*100:.1f}%)")
        print(f"Не определено: {neutral_count:,} ({neutral_count/len(df_service)*100:.1f}%)")
        
        # Создаём Excel файл с несколькими листами
        service_num = str(service).replace('.0', '')
        excel_file = f'{output_dir}/СЛУЖБА_{service_num}_ДЕТАЛЬНО_{timestamp}.xlsx'
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            
            # ЛИСТ 1: СВОДКА
            summary_data = {
                'Показатель': [
                    'Всего записей',
                    'Положительно',
                    'Отрицательно/Жалобы',
                    'Не определено',
                    '',
                    'Процент положительных',
                    'Процент отрицательных',
                    '',
                    'Уникальных инцидентов',
                    'Уникальных карт',
                    'Уникальных телефонов'
                ],
                'Значение': [
                    len(df_service),
                    positive_count,
                    negative_count,
                    neutral_count,
                    '',
                    f"{positive_count/len(df_service)*100:.1f}%",
                    f"{negative_count/len(df_service)*100:.1f}%",
                    '',
                    df_service['Инцидент_112'].nunique(),
                    df_service['Карта_112'].nunique(),
                    df_service['Телефон_112'].nunique()
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='СВОДКА', index=False)
            
            # ЛИСТ 2: ПОЛОЖИТЕЛЬНЫЕ
            df_positive = df_service[df_service['Категория'] == 'Положительно'].copy()
            if len(df_positive) > 0:
                # Ограничиваем до 100,000 строк для Excel
                if len(df_positive) > 100000:
                    df_positive = df_positive.head(100000)
                    print(f"  ВНИМАНИЕ: Положительных ограничено до 100,000 строк")
                
                df_positive.to_excel(writer, sheet_name='ПОЛОЖИТЕЛЬНЫЕ', index=False)
                print(f"  Лист ПОЛОЖИТЕЛЬНЫЕ: {len(df_positive):,} записей")
            
            # ЛИСТ 3: ОТРИЦАТЕЛЬНЫЕ/ЖАЛОБЫ
            df_negative = df_service[df_service['Категория'] == 'Отрицательно'].copy()
            if len(df_negative) > 0:
                if len(df_negative) > 100000:
                    df_negative = df_negative.head(100000)
                    print(f"  ВНИМАНИЕ: Отрицательных ограничено до 100,000 строк")
                
                df_negative.to_excel(writer, sheet_name='ЖАЛОБЫ_ОТРИЦАТЕЛЬНЫЕ', index=False)
                print(f"  Лист ЖАЛОБЫ: {len(df_negative):,} записей")
            
            # ЛИСТ 4: ВСЕ ДАННЫЕ (детальный)
            if len(df_service) <= 100000:
                df_service.to_excel(writer, sheet_name='ВСЕ_ДАННЫЕ', index=False)
                print(f"  Лист ВСЕ_ДАННЫЕ: {len(df_service):,} записей")
            else:
                df_service.head(100000).to_excel(writer, sheet_name='ВСЕ_ДАННЫЕ', index=False)
                print(f"  Лист ВСЕ_ДАННЫЕ: ограничено до 100,000 из {len(df_service):,}")
        
        print(f"\nСоздан файл: {excel_file}")
        
        # Также создаём CSV файл с жалобами для полного доступа
        if len(df_negative) > 0:
            csv_file = f'{output_dir}/СЛУЖБА_{service_num}_ЖАЛОБЫ_{timestamp}.csv'
            df_negative.to_csv(csv_file, index=False, encoding='utf-8-sig')
            print(f"Создан CSV с жалобами: {csv_file}")
    
    return output_dir

def create_consolidated_summary(df, output_dir):
    """Создание сводного отчёта по всем службам"""
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    txt_file = f'{output_dir}/СВОДНЫЙ_ОТЧЁТ_ПО_СЛУЖБАМ_{timestamp}.txt'
    
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("СВОДНЫЙ ОТЧЁТ: ПОЛОЖИТЕЛЬНЫЕ И ОТРИЦАТЕЛЬНЫЕ ПО СЛУЖБАМ\n")
        f.write("="*80 + "\n")
        f.write(f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Всего записей: {len(df):,}\n\n")
        
        services = sorted(df['Служба_112'].dropna().unique())
        
        for service in services:
            df_service = df[df['Служба_112'] == service]
            
            # Определяем категории
            positive = (
                (df_service['Положительно'].astype(str).str.contains('Положительн', case=False, na=False)) |
                (df_service['Статус_Sheets'].astype(str).str.contains('Положительн', case=False, na=False))
            ).sum()
            
            negative = (
                (df_service['Есть_жалоба'] == True) |
                (df_service['Жалоба'].notna()) |
                (df_service['Статус_Sheets'].astype(str).str.contains('Отрицательн', case=False, na=False))
            ).sum()
            
            neutral = len(df_service) - positive - negative
            
            service_num = str(service).replace('.0', '')
            
            f.write("="*80 + "\n")
            f.write(f"СЛУЖБА {service_num}\n")
            f.write("-"*80 + "\n")
            f.write(f"Всего записей:           {len(df_service):>10,}\n")
            f.write(f"Положительно:            {positive:>10,} ({positive/len(df_service)*100:5.1f}%)\n")
            f.write(f"Отрицательно/Жалобы:     {negative:>10,} ({negative/len(df_service)*100:5.1f}%)\n")
            f.write(f"Не определено:           {neutral:>10,} ({neutral/len(df_service)*100:5.1f}%)\n")
            f.write("\n")
            f.write(f"Уникальных инцидентов:   {df_service['Инцидент_112'].nunique():>10,}\n")
            f.write(f"Уникальных карт:         {df_service['Карта_112'].nunique():>10,}\n")
            f.write("\n")
    
    print(f"\nСоздан сводный отчёт: {txt_file}")
    return txt_file

def main():
    """Основная функция"""
    print("="*80)
    print("СОЗДАНИЕ ДЕТАЛЬНОЙ ОТЧЁТНОСТИ ПО СЛУЖБАМ")
    print("="*80)
    
    # Загрузка данных
    df = load_latest_matching()
    
    # Создание отчётов по службам
    output_dir = create_service_reports(df)
    
    # Создание сводного отчёта
    summary_file = create_consolidated_summary(df, output_dir)
    
    print("\n" + "="*80)
    print("ГОТОВО!")
    print("="*80)
    print(f"Папка с отчётами: {output_dir}")
    print(f"Создано файлов по службам: {len(df['Служба_112'].dropna().unique())}")

if __name__ == '__main__':
    main()
