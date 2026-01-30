"""
Обработка загруженных файлов с заявками (адреса, номера)
Следит за папкой incoming_data/applications/ и обрабатывает новые файлы
"""

import sqlite3
from pathlib import Path
import pandas as pd
from datetime import datetime
import shutil

# Пути
BASE_DIR = Path(__file__).parent.parent.parent
INCOMING_DIR = BASE_DIR / "incoming_data" / "applications"
PROCESSED_DIR = BASE_DIR / "incoming_data" / "processed"
DB_PATH = BASE_DIR / "data" / "fiksa_database.db"

# Создание папок
INCOMING_DIR.mkdir(parents=True, exist_ok=True)
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

def get_db_connection():
    """Подключение к базе данных"""
    return sqlite3.connect(DB_PATH)

def create_applications_table():
    """Создание таблицы для заявок если её нет"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS applications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_date TEXT NOT NULL,
        application_number TEXT,
        card_number TEXT,
        full_name TEXT,
        phone TEXT,
        address TEXT,
        status TEXT,
        notes TEXT,
        source_file TEXT
    )
    """)
    
    conn.commit()
    conn.close()

def process_excel_file(file_path):
    """Обработка Excel файла с заявками"""
    print(f"  📄 Обработка: {file_path.name}")
    
    try:
        # Чтение Excel
        df = pd.read_excel(file_path)
        
        print(f"  📋 Колонки в файле: {list(df.columns)}")
        
        # Определение колонок (гибкая настройка)
        column_mapping = {}
        for col in df.columns:
            col_lower = str(col).lower()
            
            # Номер заявки (первая колонка - "Номер Карты" это на самом деле номер заявки)
            if 'номер' in col_lower and 'карт' in col_lower:
                column_mapping['application_number'] = col
            elif col_lower in ['номер карты', 'карта', 'card']:
                column_mapping['application_number'] = col
            
            # Дата и время
            elif 'дата' in col_lower or 'время' in col_lower or 'date' in col_lower:
                column_mapping['date_time'] = col
            
            # Область
            elif 'область' in col_lower or 'region' in col_lower:
                column_mapping['region'] = col
            
            # Район
            elif 'район' in col_lower or 'district' in col_lower:
                column_mapping['district'] = col
            
            # Адрес
            elif 'адрес' in col_lower or 'address' in col_lower:
                column_mapping['address'] = col
            
            # Телефон
            elif 'телефон' in col_lower and 'номер' not in col_lower:
                column_mapping['phone'] = col
            elif col_lower in ['телефон', 'phone']:
                column_mapping['phone'] = col
            
            # Остальные поля
            elif 'фио' in col_lower or 'имя' in col_lower:
                column_mapping['full_name'] = col
            elif 'статус' in col_lower:
                column_mapping['status'] = col
            elif 'примечан' in col_lower or 'коммент' in col_lower:
                column_mapping['notes'] = col
        
        # Если не нашли ключевые колонки, пробуем по индексу
        if not column_mapping.get('application_number') and len(df.columns) >= 1:
            column_mapping['application_number'] = df.columns[0]
        if not column_mapping.get('date_time') and len(df.columns) >= 2:
            column_mapping['date_time'] = df.columns[1]
        if not column_mapping.get('region') and len(df.columns) >= 3:
            column_mapping['region'] = df.columns[2]
        if not column_mapping.get('district') and len(df.columns) >= 4:
            column_mapping['district'] = df.columns[3]
        if not column_mapping.get('address') and len(df.columns) >= 5:
            column_mapping['address'] = df.columns[4]
        if not column_mapping.get('phone') and len(df.columns) >= 6:
            column_mapping['phone'] = df.columns[5]
        
        print(f"  🔍 Найденные колонки: {column_mapping}")
        
        if not column_mapping.get('application_number'):
            print(f"  ⚠️  Не удалось определить колонку с номером заявки")
            return 0
        
        # Подключение к БД
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Импорт данных
        imported_count = 0
        import_date = datetime.now().strftime('%Y-%m-%d')
        
        for _, row in df.iterrows():
            # Формируем полный адрес из области, района и адреса
            region = str(row.get(column_mapping.get('region', ''), '')).strip()
            district = str(row.get(column_mapping.get('district', ''), '')).strip()
            address = str(row.get(column_mapping.get('address', ''), '')).strip()
            
            full_address = ', '.join(filter(None, [region, district, address]))
            
            # Телефон - добавляем префикс +998 если нужно
            phone = str(row.get(column_mapping.get('phone', ''), '')).strip()
            if phone and not phone.startswith('+'):
                if len(phone) == 9:
                    phone = f'+998{phone}'
            
            values = {
                'import_date': import_date,
                'application_number': str(row.get(column_mapping.get('application_number', ''), '')).strip(),
                'card_number': '',  # Номера карты нет в этих данных
                'full_name': str(row.get(column_mapping.get('full_name', ''), '')).strip(),
                'phone': phone,
                'address': full_address,
                'status': str(row.get(column_mapping.get('status', ''), '')).strip(),
                'notes': f"Дата: {row.get(column_mapping.get('date_time', ''), '')}",
                'source_file': file_path.name
            }
            
            # Пропускаем пустые строки
            if not values['application_number'] and not values['phone']:
                continue
            
            cursor.execute("""
                INSERT INTO applications 
                (import_date, application_number, card_number, full_name, phone, address, status, notes, source_file)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                values['import_date'],
                str(values['application_number']),
                str(values['card_number']),
                str(values['full_name']),
                str(values['phone']),
                str(values['address']),
                str(values['status']),
                str(values['notes']),
                values['source_file']
            ))
            
            imported_count += 1
        
        conn.commit()
        conn.close()
        
        print(f"  ✅ Импортировано записей: {imported_count}")
        return imported_count
        
    except Exception as e:
        print(f"  ❌ Ошибка обработки файла: {e}")
        return 0

def process_csv_file(file_path):
    """Обработка CSV файла с заявками"""
    print(f"  📄 Обработка: {file_path.name}")
    
    try:
        # Попытка с разными кодировками
        for encoding in ['utf-8', 'cp1251', 'windows-1251']:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        
        # Передаем в обработчик Excel
        return process_excel_file(file_path)
        
    except Exception as e:
        print(f"  ❌ Ошибка обработки файла: {e}")
        return 0

def scan_and_process():
    """Сканирование папки и обработка новых файлов"""
    print("\n" + "=" * 80)
    print("📂 ОБРАБОТКА ЗАЯВОК")
    print("=" * 80)
    print(f"📁 Папка для загрузки: {INCOMING_DIR}")
    
    # Создание таблицы
    create_applications_table()
    
    # Поиск файлов
    excel_files = list(INCOMING_DIR.glob("*.xlsx")) + list(INCOMING_DIR.glob("*.xls"))
    csv_files = list(INCOMING_DIR.glob("*.csv"))
    
    all_files = excel_files + csv_files
    
    if not all_files:
        print("  ℹ️  Нет новых файлов для обработки")
        return
    
    print(f"\n📋 Найдено файлов: {len(all_files)}")
    
    total_imported = 0
    
    for file_path in all_files:
        if file_path.suffix.lower() in ['.xlsx', '.xls']:
            imported = process_excel_file(file_path)
        elif file_path.suffix.lower() == '.csv':
            imported = process_csv_file(file_path)
        else:
            continue
        
        total_imported += imported
        
        # Перемещение обработанного файла
        if imported > 0:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
            destination = PROCESSED_DIR / new_name
            
            shutil.move(str(file_path), str(destination))
            print(f"  📦 Файл перемещен в: processed/{new_name}")
    
    print("\n" + "=" * 80)
    print(f"✅ ОБРАБОТКА ЗАВЕРШЕНА. Всего импортировано: {total_imported}")
    print("=" * 80)

def main():
    """Главная функция"""
    try:
        scan_and_process()
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        raise

if __name__ == "__main__":
    main()
