"""
Создание детальной отчётности по службам с анализом по регионам
"""
import pandas as pd
import os
from datetime import datetime
from pathlib import Path

def load_latest_matching():
    """Загрузка данных из ALL_DATA_FIXED.csv"""
    data_file = Path('data') / 'ALL_DATA_FIXED.csv'
    
    # Если не в data/, проверяем корень
    if not data_file.exists():
        data_file = Path('ALL_DATA_FIXED.csv')
    
    if not data_file.exists():
        raise FileNotFoundError(f"Не найден файл данных: {data_file}")
    
    print(f"Загружаю: {data_file}")
    
    df = pd.read_csv(data_file, encoding='utf-8-sig', low_memory=False)
    print(f"Загружено записей: {len(df):,}")
    
    return df

def create_service_reports_with_regions(df):
    """Создание отчётов по каждой службе с анализом по регионам"""
    
    # Получаем список служб - проверяем разные варианты названий колонки
    if 'Служба_112' in df.columns:
        service_col = 'Служба_112'
    elif 'Служба' in df.columns:
        service_col = 'Служба'
    else:
        raise ValueError("Не найдена колонка со службой (Служба_112 или Служба)")
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M')
    
    output_dir = 'reports/службы_детально'
    os.makedirs(output_dir, exist_ok=True)
    
    for service in services:
        print(f"\n{'='*80}")
        print(f"СЛУЖБА {service}")
        print('='*80)
        
        df_service = df[df[service_col] == service].copy()
        
        # Уникальные записи по номеру карты
        card_col = 'Номер карты' if 'Номер карты' in df_service.columns else 'Карта_112'
        df_service_unique = df_service.drop_duplicates(subset=[card_col], keep='first')
        
        print(f"Всего записей (детальных): {len(df_service):,}")
        print(f"Уникальных инцидентов: {len(df_service_unique):,}")
        
        # Определяем категории НА УНИКАЛЬНЫХ записях
        df_service_unique['Категория'] = 'Не определено'
        
        # Определяем какая колонка со статусом есть
        status_col = 'Статус' if 'Статус' in df_service_unique.columns else 'Статус_Sheets'
        comment_col = 'Комментарий' if 'Комментарий' in df_service_unique.columns else 'Жалоба'
        
        mask_positive = df_service_unique[status_col].astype(str).str.contains('Положительн|qanoatlantir|қаноатлантир', case=False, na=False, regex=True)
        df_service_unique.loc[mask_positive, 'Категория'] = 'Положительно'
        
        mask_negative = (
            df_service_unique[status_col].astype(str).str.contains('Отрицательн|qanoatlantirilmadi', case=False, na=False, regex=True) |
            (df_service_unique[comment_col].notna() & (df_service_unique[comment_col].astype(str).str.strip() != ''))
        )
        df_service_unique.loc[mask_negative, 'Категория'] = 'Отрицательно'
        
        positive_count = (df_service_unique['Категория'] == 'Положительно').sum()
        negative_count = (df_service_unique['Категория'] == 'Отрицательно').sum()
        
        print(f"Положительно: {positive_count:,} ({positive_count/len(df_service_unique)*100:.1f}%)")
        print(f"Отрицательно/Жалобы: {negative_count:,} ({negative_count/len(df_service_unique)*100:.1f}%)")
        
        # Также применяем категорию к полным данным для листов
        df_service['Категория'] = 'Не определено'
        mask_positive_all = df_service[status_col].astype(str).str.contains('Положительн|qanoatlantir|қаноатлантир', case=False, na=False, regex=True)
        mask_negative_all = (
            df_service[status_col].astype(str).str.contains('Отрицательн|qanoatlantirilmadi', case=False, na=False, regex=True) |
            (df_service[comment_col].notna() & (df_service[comment_col].astype(str).str.strip() != ''))
        )
        df_service.loc[mask_positive_all, 'Категория'] = 'Положительно'
        df_service.loc[mask_negative_all, 'Категория'] = 'Отрицательно'
        
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
                    df_service[card_col].nunique(),
                    df_service['Номер телефона'].nunique() if 'Номер телефона' in df_service.columns else df_service.get('Телефон_112', pd.Series([0])).nunique(),
                    df_service['Регион'].nunique() if 'Регион' in df_service.columns else 0,
                    df_service['Район'].nunique() if 'Район' in df_service.columns else df_service.get('Район_112', pd.Series([0])).nunique()
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='СВОДКА', index=False)
            print(f"  ✓ Лист СВОДКА")
            
            # ЛИСТ 2: АНАЛИЗ ПО РЕГИОНАМ
            region_col = 'Регион' if 'Регион' in df_service.columns else 'Область'
            if region_col in df_service.columns:
                region_analysis = df_service.groupby(region_col).agg({
                    card_col: 'count',
                    'Категория': lambda x: (x == 'Положительно').sum()
                }).reset_index()
                region_analysis.columns = ['Регион', 'Всего записей', 'Положительно']
                region_analysis['Отрицательно'] = df_service.groupby(region_col).apply(
                    lambda x: (x['Категория'] == 'Отрицательно').sum()
                ).values
                region_analysis['% Положительных'] = (
                    region_analysis['Положительно'] / region_analysis['Всего записей'] * 100
                ).round(1)
                region_analysis = region_analysis.sort_values('Всего записей', ascending=False)
                region_analysis.to_excel(writer, sheet_name='ПО_РЕГИОНАМ', index=False)
                print(f"  ✓ Лист ПО_РЕГИОНАМ ({len(region_analysis)} регионов)")
            else:
                print(f"  ⚠ Колонка региона не найдена, пропускаем анализ по регионам")
            
            # ЛИСТ 3: АНАЛИЗ ПО РАЙОНАМ (ТОП-100)
            district_col = 'Район' if 'Район' in df_service.columns else 'Район_112'
            if district_col in df_service.columns:
                district_analysis = df_service.groupby(district_col).agg({
                    card_col: 'count',
                    'Категория': lambda x: (x == 'Положительно').sum()
                }).reset_index()
                district_analysis.columns = ['Район', 'Всего записей', 'Положительно']
                district_analysis['Отрицательно'] = df_service.groupby(district_col).apply(
                    lambda x: (x['Категория'] == 'Отрицательно').sum()
                ).values
                district_analysis['% Положительных'] = (
                    district_analysis['Положительно'] / district_analysis['Всего записей'] * 100
                ).round(1)
                district_analysis = district_analysis.sort_values('Всего записей', ascending=False).head(100)
                district_analysis.to_excel(writer, sheet_name='ПО_РАЙОНАМ_ТОП100', index=False)
                print(f"  ✓ Лист ПО_РАЙОНАМ (топ-100)")
            else:
                print(f"  ⚠ Колонка района не найдена, пропускаем анализ по районам")
            
            # ЛИСТ 4: АНАЛИЗ ПО ОПЕРАТОРАМ (ТОП-50)
            operator_col = 'Оператор' if 'Оператор' in df_service.columns else 'Оператор_112'
            if operator_col in df_service.columns:
                operator_analysis = df_service.groupby(operator_col).agg({
                    card_col: 'count',
                    'Категория': lambda x: (x == 'Положительно').sum()
                }).reset_index()
                operator_analysis.columns = ['Оператор', 'Всего записей', 'Положительно']
                operator_analysis['Отрицательно'] = df_service.groupby(operator_col).apply(
                    lambda x: (x['Категория'] == 'Отрицательно').sum()
                ).values
                operator_analysis['% Положительных'] = (
                    operator_analysis['Положительно'] / operator_analysis['Всего записей'] * 100
                ).round(1)
                operator_analysis = operator_analysis.sort_values('Всего записей', ascending=False).head(50)
                operator_analysis.to_excel(writer, sheet_name='ПО_ОПЕРАТОРАМ_ТОП50', index=False)
                print(f"  ✓ Лист ПО_ОПЕРАТОРАМ (топ-50)")
            else:
                print(f"  ⚠ Колонка оператора не найдена, пропускаем анализ по операторам")
            
            # ЛИСТ 5: АНАЛИЗ ПО СТАТУСАМ
            if status_col in df_service.columns:
                status_analysis = df_service.groupby(status_col).agg({
                    card_col: 'count',
                    'Категория': lambda x: (x == 'Положительно').sum()
                }).reset_index()
                status_analysis.columns = ['Статус', 'Всего записей', 'Положительно']
                status_analysis['Отрицательно'] = df_service.groupby(status_col).apply(
                    lambda x: (x['Категория'] == 'Отрицательно').sum()
                ).values
                status_analysis['% от общего'] = (
                    status_analysis['Всего записей'] / len(df_service) * 100
                ).round(1)
                status_analysis = status_analysis.sort_values('Всего записей', ascending=False)
                status_analysis.to_excel(writer, sheet_name='ПО_СТАТУСАМ', index=False)
                print(f"  ✓ Лист ПО_СТАТУСАМ ({len(status_analysis)} статусов)")
            else:
                print(f"  ⚠ Колонка статуса не найдена, пропускаем анализ по статусам")
            
            # ЛИСТ 6: АНАЛИЗ ПО АГЕНТАМ
            agent_col = 'Документ' if 'Документ' in df_service.columns else 'Оператор фиксировавший'
            if agent_col in df_service.columns:
                agent_analysis = df_service.groupby(agent_col).agg({
                    card_col: 'count',
                    'Категория': lambda x: (x == 'Положительно').sum()
                }).reset_index()
                agent_analysis.columns = ['Агент', 'Всего записей', 'Положительно']
                agent_analysis['Отрицательно'] = df_service.groupby(agent_col).apply(
                    lambda x: (x['Категория'] == 'Отрицательно').sum()
                ).values
                agent_analysis['% Положительных'] = (
                    agent_analysis['Положительно'] / agent_analysis['Всего записей'] * 100
                ).round(1)
                agent_analysis = agent_analysis.sort_values('Всего записей', ascending=False)
                agent_analysis.to_excel(writer, sheet_name='ПО_АГЕНТАМ', index=False)
                print(f"  ✓ Лист ПО_АГЕНТАМ ({len(agent_analysis)} агентов)")
            else:
                print(f"  ⚠ Колонка агента не найдена, пропускаем анализ по агентам")
            
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
