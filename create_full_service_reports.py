"""
Создание детальной отчётности по службам с анализом по регионам
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

def create_service_reports_with_regions(df):
    """Создание отчётов по каждой службе с анализом по регионам"""
    
    services = sorted(df['Служба_112'].dropna().unique())
    print(f"\nНайдено служб: {services}")
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    
    output_dir = 'reports/службы_детально'
    os.makedirs(output_dir, exist_ok=True)
    
    for service in services:
        print(f"\n{'='*80}")
        print(f"СЛУЖБА {service}")
        print('='*80)
        
        df_service = df[df['Служба_112'] == service].copy()
        
        # ВАЖНО: Удаляем дубликаты по инциденту из Sheets для правильного подсчёта
        # Один инцидент может быть в нескольких службах, но считаем его один раз
        df_service_unique = df_service.drop_duplicates(subset=['Инцидент_Sheets'], keep='first')
        
        print(f"Всего записей (детальных): {len(df_service):,}")
        print(f"Уникальных инцидентов: {len(df_service_unique):,}")
        
        # Определяем категории НА УНИКАЛЬНЫХ инцидентах
        df_service_unique['Категория'] = 'Не определено'
        
        mask_positive = (
            (df_service_unique['Положительно'].astype(str).str.contains('Положительн', case=False, na=False)) |
            (df_service_unique['Статус_Sheets'].astype(str).str.contains('Положительн', case=False, na=False))
        )
        df_service_unique.loc[mask_positive, 'Категория'] = 'Положительно'
        
        mask_negative = (
            (df_service_unique['Есть_жалоба'] == True) |
            (df_service_unique['Жалоба'].notna()) |
            (df_service_unique['Статус_Sheets'].astype(str).str.contains('Отрицательн', case=False, na=False))
        )
        df_service_unique.loc[mask_negative, 'Категория'] = 'Отрицательно'
        
        positive_count = (df_service_unique['Категория'] == 'Положительно').sum()
        negative_count = (df_service_unique['Категория'] == 'Отрицательно').sum()
        
        print(f"Положительно: {positive_count:,} ({positive_count/len(df_service_unique)*100:.1f}%)")
        print(f"Отрицательно/Жалобы: {negative_count:,} ({negative_count/len(df_service_unique)*100:.1f}%)")
        
        # Также применяем категорию к полным данным для листов
        df_service['Категория'] = 'Не определено'
        df_service.loc[mask_positive.reindex(df_service.index, fill_value=False), 'Категория'] = 'Положительно'
        df_service.loc[mask_negative.reindex(df_service.index, fill_value=False), 'Категория'] = 'Отрицательно'
        
        # Создаём Excel файл
        service_num = str(service).replace('.0', '')
        excel_file = f'{output_dir}/СЛУЖБА_{service_num}_ПОЛНЫЙ_АНАЛИЗ_{timestamp}.xlsx'
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            
            # ЛИСТ 1: СВОДКА
            summary_data = {
                'Показатель': [
                    'Уникальных инцидентов (основа)',
                    'Детальных записей (с службами)',
                    '',
                    'Положительно',
                    'Отрицательно/Жалобы',
                    '',
                    'Процент положительных',
                    'Процент отрицательных',
                    '',
                    'Уникальных карт 112',
                    'Уникальных телефонов',
                    'Уникальных регионов',
                    'Уникальных районов'
                ],
                'Значение': [
                    len(df_service_unique),
                    len(df_service),
                    '',
                    positive_count,
                    negative_count,
                    '',
                    f"{positive_count/len(df_service_unique)*100:.1f}%",
                    f"{negative_count/len(df_service_unique)*100:.1f}%",
                    '',
                    df_service['Карта_112'].nunique(),
                    df_service['Телефон_112'].nunique(),
                    df_service['Регион'].nunique(),
                    df_service['Район_112'].nunique()
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='СВОДКА', index=False)
            print(f"  ✓ Лист СВОДКА")
            
            # ЛИСТ 2: АНАЛИЗ ПО РЕГИОНАМ
            region_analysis = df_service.groupby('Регион').agg({
                'Инцидент_112': 'count',
                'Категория': lambda x: (x == 'Положительно').sum()
            }).reset_index()
            region_analysis.columns = ['Регион', 'Всего записей', 'Положительно']
            region_analysis['Отрицательно'] = df_service.groupby('Регион').apply(
                lambda x: (x['Категория'] == 'Отрицательно').sum()
            ).values
            region_analysis['% Положительных'] = (
                region_analysis['Положительно'] / region_analysis['Всего записей'] * 100
            ).round(1)
            region_analysis = region_analysis.sort_values('Всего записей', ascending=False)
            region_analysis.to_excel(writer, sheet_name='ПО_РЕГИОНАМ', index=False)
            print(f"  ✓ Лист ПО_РЕГИОНАМ ({len(region_analysis)} регионов)")
            
            # ЛИСТ 3: АНАЛИЗ ПО РАЙОНАМ (ТОП-100)
            district_analysis = df_service.groupby('Район_112').agg({
                'Инцидент_112': 'count',
                'Категория': lambda x: (x == 'Положительно').sum()
            }).reset_index()
            district_analysis.columns = ['Район', 'Всего записей', 'Положительно']
            district_analysis['Отрицательно'] = df_service.groupby('Район_112').apply(
                lambda x: (x['Категория'] == 'Отрицательно').sum()
            ).values
            district_analysis['% Положительных'] = (
                district_analysis['Положительно'] / district_analysis['Всего записей'] * 100
            ).round(1)
            district_analysis = district_analysis.sort_values('Всего записей', ascending=False).head(100)
            district_analysis.to_excel(writer, sheet_name='ПО_РАЙОНАМ_ТОП100', index=False)
            print(f"  ✓ Лист ПО_РАЙОНАМ (топ-100)")
            
            # ЛИСТ 4: АНАЛИЗ ПО ОПЕРАТОРАМ (ТОП-50)
            operator_analysis = df_service.groupby('Оператор_112').agg({
                'Инцидент_112': 'count',
                'Категория': lambda x: (x == 'Положительно').sum()
            }).reset_index()
            operator_analysis.columns = ['Оператор', 'Всего записей', 'Положительно']
            operator_analysis['Отрицательно'] = df_service.groupby('Оператор_112').apply(
                lambda x: (x['Категория'] == 'Отрицательно').sum()
            ).values
            operator_analysis['% Положительных'] = (
                operator_analysis['Положительно'] / operator_analysis['Всего записей'] * 100
            ).round(1)
            operator_analysis = operator_analysis.sort_values('Всего записей', ascending=False).head(50)
            operator_analysis.to_excel(writer, sheet_name='ПО_ОПЕРАТОРАМ_ТОП50', index=False)
            print(f"  ✓ Лист ПО_ОПЕРАТОРАМ (топ-50)")
            
            # ЛИСТ 5: АНАЛИЗ ПО СТАТУСАМ
            status_analysis = df_service.groupby('Статус_112').agg({
                'Инцидент_112': 'count',
                'Категория': lambda x: (x == 'Положительно').sum()
            }).reset_index()
            status_analysis.columns = ['Статус', 'Всего записей', 'Положительно']
            status_analysis['Отрицательно'] = df_service.groupby('Статус_112').apply(
                lambda x: (x['Категория'] == 'Отрицательно').sum()
            ).values
            status_analysis['% от общего'] = (
                status_analysis['Всего записей'] / len(df_service) * 100
            ).round(1)
            status_analysis = status_analysis.sort_values('Всего записей', ascending=False)
            status_analysis.to_excel(writer, sheet_name='ПО_СТАТУСАМ', index=False)
            print(f"  ✓ Лист ПО_СТАТУСАМ ({len(status_analysis)} статусов)")
            
            # ЛИСТ 6: АНАЛИЗ ПО АГЕНТАМ
            agent_analysis = df_service.groupby('Документ').agg({
                'Инцидент_112': 'count',
                'Категория': lambda x: (x == 'Положительно').sum()
            }).reset_index()
            agent_analysis.columns = ['Агент', 'Всего записей', 'Положительно']
            agent_analysis['Отрицательно'] = df_service.groupby('Документ').apply(
                lambda x: (x['Категория'] == 'Отрицательно').sum()
            ).values
            agent_analysis['% Положительных'] = (
                agent_analysis['Положительно'] / agent_analysis['Всего записей'] * 100
            ).round(1)
            agent_analysis = agent_analysis.sort_values('Всего записей', ascending=False)
            agent_analysis.to_excel(writer, sheet_name='ПО_АГЕНТАМ', index=False)
            print(f"  ✓ Лист ПО_АГЕНТАМ ({len(agent_analysis)} агентов)")
            
            # ЛИСТ 7: ПОЛОЖИТЕЛЬНЫЕ (до 50K)
            df_positive = df_service[df_service['Категория'] == 'Положительно'].copy()
            if len(df_positive) > 0:
                df_pos_sample = df_positive.head(50000)
                df_pos_sample.to_excel(writer, sheet_name='ПОЛОЖИТЕЛЬНЫЕ', index=False)
                print(f"  ✓ Лист ПОЛОЖИТЕЛЬНЫЕ ({len(df_pos_sample):,} из {len(df_positive):,})")
            
            # ЛИСТ 8: ОТРИЦАТЕЛЬНЫЕ/ЖАЛОБЫ (все)
            df_negative = df_service[df_service['Категория'] == 'Отрицательно'].copy()
            if len(df_negative) > 0:
                if len(df_negative) <= 100000:
                    df_negative.to_excel(writer, sheet_name='ЖАЛОБЫ_ОТРИЦАТЕЛЬНЫЕ', index=False)
                    print(f"  ✓ Лист ЖАЛОБЫ ({len(df_negative):,})")
                else:
                    df_negative.head(100000).to_excel(writer, sheet_name='ЖАЛОБЫ_ОТРИЦАТЕЛЬНЫЕ', index=False)
                    print(f"  ✓ Лист ЖАЛОБЫ ({100000:,} из {len(df_negative):,})")
        
        print(f"\n✅ Создан: {excel_file}")
        
        # CSV с полным списком жалоб
        if len(df_negative) > 0:
            csv_file = f'{output_dir}/СЛУЖБА_{service_num}_ЖАЛОБЫ_{timestamp}.csv'
            df_negative.to_csv(csv_file, index=False, encoding='utf-8-sig')
            print(f"✅ Создан CSV с жалобами: {csv_file}")
    
    return output_dir

def create_consolidated_reports(df, output_dir):
    """Создание общих сводных отчётов"""
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    
    # СВОДНЫЙ ТЕКСТОВЫЙ ОТЧЁТ
    txt_file = f'{output_dir}/СВОДНЫЙ_ОТЧЁТ_{timestamp}.txt'
    
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("СВОДНЫЙ ОТЧЁТ ПО ВСЕМ СЛУЖБАМ\n")
        f.write("="*80 + "\n")
        f.write(f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Всего записей: {len(df):,}\n\n")
        
        services = sorted(df['Служба_112'].dropna().unique())
        
        for service in services:
            df_s = df[df['Служба_112'] == service]
            
            positive = (
                (df_s['Положительно'].astype(str).str.contains('Положительн', case=False, na=False)) |
                (df_s['Статус_Sheets'].astype(str).str.contains('Положительн', case=False, na=False))
            ).sum()
            
            negative = (
                (df_s['Есть_жалоба'] == True) |
                (df_s['Жалоба'].notna()) |
                (df_s['Статус_Sheets'].astype(str).str.contains('Отрицательн', case=False, na=False))
            ).sum()
            
            service_num = str(service).replace('.0', '')
            
            f.write("="*80 + "\n")
            f.write(f"СЛУЖБА {service_num}\n")
            f.write("-"*80 + "\n")
            f.write(f"Всего:                   {len(df_s):>10,}\n")
            f.write(f"Положительно:            {positive:>10,} ({positive/len(df_s)*100:5.1f}%)\n")
            f.write(f"Отрицательно/Жалобы:     {negative:>10,} ({negative/len(df_s)*100:5.1f}%)\n")
            f.write(f"Уникальных инцидентов:   {df_s['Инцидент_112'].nunique():>10,}\n")
            f.write(f"Уникальных регионов:     {df_s['Регион'].nunique():>10,}\n")
            f.write(f"Уникальных районов:      {df_s['Район_112'].nunique():>10,}\n")
            f.write("\n")
    
    print(f"\n✅ Создан сводный отчёт: {txt_file}")
    
    # СВОДНЫЙ EXCEL ПО РЕГИОНАМ
    excel_file = f'{output_dir}/СВОДКА_ПО_РЕГИОНАМ_ВСЕ_СЛУЖБЫ_{timestamp}.xlsx'
    
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        for service in sorted(df['Служба_112'].dropna().unique()):
            df_s = df[df['Служба_112'] == service]
            
            region_stats = df_s.groupby('Регион').agg({
                'Инцидент_112': 'count'
            }).reset_index()
            region_stats.columns = ['Регион', 'Всего']
            region_stats = region_stats.sort_values('Всего', ascending=False)
            
            service_num = str(service).replace('.0', '')
            region_stats.to_excel(writer, sheet_name=f'Служба_{service_num}', index=False)
    
    print(f"✅ Создан сводный Excel по регионам: {excel_file}")
    
    return txt_file, excel_file

def main():
    """Основная функция"""
    print("="*80)
    print("СОЗДАНИЕ ПОЛНОЙ ОТЧЁТНОСТИ ПО СЛУЖБАМ С АНАЛИЗОМ ПО РЕГИОНАМ")
    print("="*80)
    
    df = load_latest_matching()
    
    output_dir = create_service_reports_with_regions(df)
    
    create_consolidated_reports(df, output_dir)
    
    print("\n" + "="*80)
    print("✅ ГОТОВО!")
    print("="*80)
    print(f"Папка: {output_dir}")

if __name__ == '__main__':
    main()
