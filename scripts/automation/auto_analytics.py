"""
Автоматическая аналитика данных с генерацией отчетов
"""

import sqlite3
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
from collections import Counter

# Пути
BASE_DIR = Path(__file__).parent.parent.parent
DB_PATH = BASE_DIR / "data" / "fiksa_database.db"
OUTPUT_DIR = BASE_DIR / "output" / "reports"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_db_connection():
    """Подключение к базе данных"""
    return sqlite3.connect(DB_PATH)

def generate_daily_report():
    """Генерация общего отчета по всем данным в базе"""
    conn = get_db_connection()
    
    # Получаем все данные (не только последнюю дату)
    query = "SELECT call_date FROM fixations GROUP BY call_date ORDER BY call_date DESC LIMIT 1"
    result = pd.read_sql_query(query, conn)
    last_date = result['call_date'][0] if not result.empty else None
    
    if not last_date:
        print("⚠️  Нет данных в базе")
        return
    
    print(f"📅 Генерация отчета за {last_date}")
    
    # Общая статистика по ВСЕМ данным
    total_query = """
    SELECT 
        COUNT(*) as total_records,
        COUNT(DISTINCT operator_id) as total_operators,
        COUNT(DISTINCT card_number) as unique_cards
    FROM fixations
    """
    total_stats = pd.read_sql_query(total_query, conn)
    
    # Статистика по операторам (все данные) из представления
    operator_query = """
    SELECT 
        operator_name,
        COUNT(*) as records,
        COUNT(DISTINCT card_number) as unique_cards,
        COUNT(CASE WHEN fixation_status = 'положительный' THEN 1 END) as positive,
        COUNT(CASE WHEN fixation_status = 'отрицательный' THEN 1 END) as negative,
        COUNT(CASE WHEN fixation_status = 'тишина' THEN 1 END) as silence,
        COUNT(CASE WHEN fixation_status = 'соединение прервано' THEN 1 END) as interrupted,
        COUNT(CASE WHEN fixation_status = 'перезвонить' THEN 1 END) as callback,
        COUNT(CASE WHEN fixation_status = 'не в зоне' THEN 1 END) as no_zone,
        COUNT(CASE WHEN fixation_status = 'недоступен' THEN 1 END) as unavailable
    FROM v_fixations_full
    GROUP BY operator_name
    ORDER BY records DESC
    """
    operator_stats = pd.read_sql_query(operator_query, conn)
    
    # Статистика по статусам (все данные)
    status_query = """
    SELECT 
        status,
        COUNT(*) as count,
        ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM fixations), 2) as percentage
    FROM fixations
    WHERE status != ''
    GROUP BY status
    ORDER BY count DESC
    """
    status_stats = pd.read_sql_query(status_query, conn)
    
    # Создание HTML отчета
    report_path = OUTPUT_DIR / f"daily_report_{last_date}.html"
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Ежедневный отчет по фиксации - {last_date}</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 20px;
                background-color: #f5f5f5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background-color: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            h1 {{
                color: #2c3e50;
                border-bottom: 3px solid #3498db;
                padding-bottom: 10px;
            }}
            h2 {{
                color: #34495e;
                margin-top: 30px;
                border-left: 4px solid #3498db;
                padding-left: 10px;
            }}
            .stats-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 20px;
                margin: 20px 0;
            }}
            .stat-card {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .stat-card h3 {{
                margin: 0;
                font-size: 14px;
                opacity: 0.9;
            }}
            .stat-card .value {{
                font-size: 36px;
                font-weight: bold;
                margin: 10px 0;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 20px;
            }}
            th, td {{
                padding: 12px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }}
            th {{
                background-color: #3498db;
                color: white;
                font-weight: bold;
            }}
            tr:hover {{
                background-color: #f5f5f5;
            }}
            .positive {{ color: #27ae60; font-weight: bold; }}
            .negative {{ color: #e74c3c; font-weight: bold; }}
            .footer {{
                margin-top: 30px;
                text-align: center;
                color: #7f8c8d;
                font-size: 12px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 Ежедневный отчет по фиксации</h1>
            <p><strong>Дата:</strong> {last_date}</p>
            <p><strong>Время создания:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            
            <h2>📈 Общая статистика</h2>
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>Всего записей</h3>
                    <div class="value">{total_stats['total_records'][0]:,}</div>
                </div>
                <div class="stat-card">
                    <h3>Активных операторов</h3>
                    <div class="value">{total_stats['total_operators'][0]}</div>
                </div>
                <div class="stat-card">
                    <h3>Уникальных карт</h3>
                    <div class="value">{total_stats['unique_cards'][0]:,}</div>
                </div>
            </div>
            
            <h2>👥 Статистика по операторам</h2>
            {operator_stats.to_html(index=False, classes='table', border=0)}
            
            <h2>📋 Распределение по статусам</h2>
            {status_stats.to_html(index=False, classes='table', border=0)}
            
            <div class="footer">
                Отчет сгенерирован автоматически системой сбора данных FIKSA<br>
                © {datetime.now().year}
            </div>
        </div>
    </body>
    </html>
    """
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"✅ HTML отчет сохранен: {report_path}")
    
    # Excel отчет
    excel_path = OUTPUT_DIR / f"daily_report_{last_date}.xlsx"
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        total_stats.to_excel(writer, sheet_name='Общая статистика', index=False)
        operator_stats.to_excel(writer, sheet_name='По операторам', index=False)
        status_stats.to_excel(writer, sheet_name='По статусам', index=False)
    
    print(f"✅ Excel отчет сохранен: {excel_path}")
    
    conn.close()
    
    return report_path, excel_path

def generate_weekly_trends():
    """Генерация трендов за неделю"""
    conn = get_db_connection()
    
    # Последние 7 дней
    query = """
    SELECT 
        collection_date,
        COUNT(*) as total_records,
        COUNT(DISTINCT operator_name) as operators,
        COUNT(CASE WHEN status = 'положительный' THEN 1 END) as positive,
        COUNT(CASE WHEN status = 'отрицательный' THEN 1 END) as negative
    FROM fiksa_records
    WHERE collection_date >= date('now', '-7 days')
    GROUP BY collection_date
    ORDER BY collection_date
    """
    trends = pd.read_sql_query(query, conn)
    
    if not trends.empty:
        trends_path = OUTPUT_DIR / f"weekly_trends_{datetime.now().strftime('%Y%m%d')}.csv"
        trends.to_csv(trends_path, index=False, encoding='utf-8-sig')
        print(f"✅ Недельные тренды сохранены: {trends_path}")
    
    conn.close()

def main():
    """Главная функция"""
    print("\n" + "=" * 80)
    print("📊 АВТОМАТИЧЕСКАЯ АНАЛИТИКА И ОТЧЕТЫ")
    print("=" * 80)
    
    try:
        # Ежедневный отчет
        report_path, excel_path = generate_daily_report()
        
        # Недельные тренды
        generate_weekly_trends()
        
        print("\n" + "=" * 80)
        print("✅ АНАЛИТИКА ЗАВЕРШЕНА УСПЕШНО")
        print("=" * 80)
        
    except Exception as e:
        print(f"\n❌ Ошибка при генерации отчетов: {e}")
        raise

if __name__ == "__main__":
    main()
