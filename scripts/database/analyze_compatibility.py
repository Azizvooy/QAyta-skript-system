#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
АНАЛИЗ И ОЧИСТКА ПРОЕКТА
=============================================================================
Находит устаревшие файлы и файлы несовместимые с новой БД
=============================================================================
"""

from pathlib import Path
import sys

BASE_DIR = Path(__file__).parent

# Устаревшие файлы (используют старую структуру БД)
DEPRECATED_FILES = {
    'scripts/data_collection': [
        'daily_db_collector.py',  # Старый коллектор, заменен на db_import.py
        'improved_collector.py',   # Старый коллектор
    ],
    'scripts/automation': [
        'import_call_history.py',  # Старый импорт, заменен на db_import.py
    ],
    'scripts/analysis': [
        'match_test.py',           # Тестовый файл
        'match_applications.py',   # Старое сопоставление
        'analyze_calls.py',        # Старый анализ
        'analyze_all_no_filter.py',  # Старый анализ
        'analyze_2025.py',         # Старый анализ
        'analyze_data.py',         # Старый анализ
        'analyze_final.py',        # Старый анализ
        'analyze_fixed_data.py',   # Старый анализ
        'analyze_new_data.py',     # Старый анализ
        'analyze_pdf_report.py',   # Старый анализ
        'analyze_unique.py',       # Старый анализ
        'data_filter.py',          # Старый фильтр
        'show_data.py',            # Тестовый скрипт
    ],
    'scripts/data_processing': [
        'check_completeness.py',   # Проверки - не нужны с новой БД
        'check_data.py',           # Проверки
        'check_dates.py',          # Проверки
        'check_operators.py',      # Проверки
        'filter_2025.py',          # Фильтр
        'find_problems.py',        # Поиск проблем
        'fix_and_process_data.py', # Обработка
        'fix_operators.py',        # Исправление
        'replace_operator_names.py',  # Замена
        'verify_data.py',          # Проверка
    ],
    'scripts/reports': [
        'create_detailed_reports.py',  # Старый отчет, заменен на db_reports.py
    ],
}

# Файлы для обновления (используют старые таблицы)
FILES_TO_UPDATE = {
    'scripts/telegram/interactive_bot.py': 'Использует fiksa_records и call_history_112',
    'scripts/telegram/working_bot.py': 'Использует fiksa_records',
    'scripts/analysis/service_reports.py': 'Использует старые таблицы',
    'scripts/analysis/analytics_reports.py': 'Использует старые таблицы',
    'scripts/analysis/auto_report.py': 'Использует старые таблицы',
    'scripts/analysis/address_report.py': 'Использует старые таблицы',
    'scripts/analysis/advanced_address_report.py': 'Использует старые таблицы',
    'scripts/automation/auto_analytics.py': 'Использует fiksa_records',
    'scripts/automation/master_service.py': 'Использует старые таблицы',
}

# Дублирующиеся отчеты (оставляем только новые из БД)
DUPLICATE_REPORTS = [
    'scripts/reports/create_final_report.py',     # Дубликат
    'scripts/reports/create_final_word_report.py',  # Дубликат
    'scripts/reports/create_full_report.py',      # Дубликат
    'scripts/reports/create_qayta_report.py',     # Оставляем - уникальный формат
    'scripts/reports/create_word_report.py',      # Дубликат
]

def analyze_project():
    """Анализ проекта"""
    print('=' * 80)
    print('АНАЛИЗ ПРОЕКТА')
    print('=' * 80)
    
    total_deprecated = 0
    total_to_update = 0
    total_duplicates = 0
    
    # Подсчет устаревших файлов
    print('\n📦 УСТАРЕВШИЕ ФАЙЛЫ (используют старую структуру БД):')
    print('-' * 80)
    
    for folder, files in DEPRECATED_FILES.items():
        folder_path = BASE_DIR / folder
        print(f'\n📁 {folder}/')
        for file in files:
            file_path = folder_path / file
            if file_path.exists():
                size = file_path.stat().st_size / 1024
                print(f'   ❌ {file} ({size:.1f} KB)')
                total_deprecated += 1
            else:
                print(f'   ⚠️  {file} (не найден)')
    
    # Файлы для обновления
    print('\n\n🔄 ФАЙЛЫ ДЛЯ ОБНОВЛЕНИЯ (совместимость с новой БД):')
    print('-' * 80)
    
    for file_path, reason in FILES_TO_UPDATE.items():
        full_path = BASE_DIR / file_path
        if full_path.exists():
            size = full_path.stat().st_size / 1024
            print(f'   🔧 {file_path}')
            print(f'      Причина: {reason} ({size:.1f} KB)')
            total_to_update += 1
    
    # Дублирующиеся отчеты
    print('\n\n📋 ДУБЛИРУЮЩИЕСЯ ОТЧЕТЫ (заменены на db_reports.py):')
    print('-' * 80)
    
    for file_path in DUPLICATE_REPORTS:
        full_path = BASE_DIR / file_path
        if full_path.exists():
            size = full_path.stat().st_size / 1024
            print(f'   🗑️  {file_path} ({size:.1f} KB)')
            total_duplicates += 1
    
    # Итоги
    print('\n\n' + '=' * 80)
    print('ИТОГИ АНАЛИЗА:')
    print('=' * 80)
    print(f'❌ Устаревших файлов: {total_deprecated}')
    print(f'🔧 Файлов для обновления: {total_to_update}')
    print(f'🗑️  Дублирующихся файлов: {total_duplicates}')
    print(f'\n📊 ВСЕГО проблемных файлов: {total_deprecated + total_to_update + total_duplicates}')
    
    return total_deprecated, total_to_update, total_duplicates

def move_to_archive():
    """Переместить устаревшие файлы в архив"""
    archive_dir = BASE_DIR / 'archive' / 'old_structure'
    archive_dir.mkdir(parents=True, exist_ok=True)
    
    moved_count = 0
    
    print('\n\n' + '=' * 80)
    print('ПЕРЕМЕЩЕНИЕ В АРХИВ')
    print('=' * 80)
    
    # Перемещение устаревших файлов
    for folder, files in DEPRECATED_FILES.items():
        folder_path = BASE_DIR / folder
        
        for file in files:
            src = folder_path / file
            if src.exists():
                # Создаем структуру в архиве
                dest_folder = archive_dir / folder
                dest_folder.mkdir(parents=True, exist_ok=True)
                dest = dest_folder / file
                
                # Перемещаем
                src.rename(dest)
                print(f'✅ {folder}/{file} → archive/old_structure/{folder}/')
                moved_count += 1
    
    # Перемещение дублирующихся отчетов (кроме create_qayta_report.py)
    for file_path in DUPLICATE_REPORTS:
        if 'qayta' in file_path.lower():
            continue  # Оставляем qayta отчет
        
        src = BASE_DIR / file_path
        if src.exists():
            folder = Path(file_path).parent
            dest_folder = archive_dir / folder
            dest_folder.mkdir(parents=True, exist_ok=True)
            dest = dest_folder / Path(file_path).name
            
            src.rename(dest)
            print(f'✅ {file_path} → archive/old_structure/{folder}/')
            moved_count += 1
    
    print(f'\n📦 Перемещено файлов: {moved_count}')
    return moved_count

def create_compatibility_report():
    """Создать отчет о совместимости"""
    report_file = BASE_DIR / 'СОВМЕСТИМОСТЬ_БД.md'
    
    content = """# 🔄 ОТЧЕТ О СОВМЕСТИМОСТИ С НОВОЙ БД

## 📅 Дата: {date}

## ✅ СОВМЕСТИМЫЕ ФАЙЛЫ (работают с новой БД)

### База данных
- ✅ `scripts/database/db_schema.py` - Схема новой БД
- ✅ `scripts/database/db_import.py` - Импорт в новую БД
- ✅ `scripts/database/db_reports.py` - Отчеты из новой БД
- ✅ `scripts/database/db_manager.py` - Менеджер БД

### Отчеты
- ✅ `scripts/reports/create_qayta_report.py` - Уникальный формат QAYTA

### Обработка данных
- ✅ `scripts/formatting/` - Все скрипты форматирования

## 🔧 ТРЕБУЮТ ОБНОВЛЕНИЯ

### Telegram боты
- 🔧 `scripts/telegram/interactive_bot.py`
  - Использует: `fiksa_records`, `call_history_112`
  - Нужно: перейти на `v_fixations_full`, `v_applications_full`

- 🔧 `scripts/telegram/working_bot.py`
  - Использует: `fiksa_records`
  - Нужно: перейти на `v_fixations_full`

### Аналитика
- 🔧 `scripts/analysis/service_reports.py`
  - Использует: старые таблицы
  - Нужно: использовать представления (views)

- 🔧 `scripts/analysis/analytics_reports.py`
  - Использует: старые таблицы
  - Нужно: использовать представления (views)

- 🔧 `scripts/analysis/auto_report.py`
  - Использует: старые таблицы
  - Нужно: обновить запросы

- 🔧 `scripts/analysis/address_report.py`
  - Использует: старые таблицы
  - Нужно: обновить запросы

- 🔧 `scripts/analysis/advanced_address_report.py`
  - Использует: старые таблицы
  - Нужно: обновить запросы

### Автоматизация
- 🔧 `scripts/automation/auto_analytics.py`
  - Использует: `fiksa_records`
  - Нужно: перейти на `fixations`

- 🔧 `scripts/automation/master_service.py`
  - Использует: старые таблицы
  - Нужно: обновить под новую структуру

## ❌ АРХИВИРОВАННЫЕ (устаревшие)

### Коллекторы данных
- ❌ `scripts/data_collection/daily_db_collector.py` → archive/
- ❌ `scripts/data_collection/improved_collector.py` → archive/

### Импорт
- ❌ `scripts/automation/import_call_history.py` → archive/

### Анализ (старые версии)
- ❌ Все файлы `analyze_*.py` → archive/
- ❌ `match_applications.py`, `match_test.py` → archive/
- ❌ `data_filter.py`, `show_data.py` → archive/

### Обработка данных (проверки)
- ❌ Все файлы `check_*.py` → archive/
- ❌ Все файлы `fix_*.py` → archive/
- ❌ Все файлы `filter_*.py` → archive/
- ❌ `verify_data.py` → archive/

### Дублирующиеся отчеты
- ❌ `create_final_report.py` → archive/
- ❌ `create_final_word_report.py` → archive/
- ❌ `create_full_report.py` → archive/
- ❌ `create_word_report.py` → archive/
- ❌ `create_detailed_reports.py` → archive/

## 📊 НОВАЯ СТРУКТУРА БД

### Таблицы
1. `operators` - Операторы (справочник)
2. `services` - Службы (справочник)
3. `regions` - Регионы (справочник)
4. `applications` - Заявки (основная)
5. `fixations` - Фиксации/обзвоны (основная)
6. `operation_logs` - Логи операций
7. `daily_statistics` - Статистика

### Представления (Views)
1. `v_applications_full` - Полная информация о заявках
2. `v_fixations_full` - Полная информация о фиксациях
3. `v_operator_stats` - Статистика по операторам

## 🔄 КАК ОБНОВИТЬ ФАЙЛЫ

### Замена таблиц:
```python
# СТАРОЕ:
SELECT * FROM fiksa_records
SELECT * FROM call_history_112

# НОВОЕ:
SELECT * FROM v_fixations_full
SELECT * FROM v_applications_full
```

### Замена связей:
```python
# СТАРОЕ:
LEFT JOIN fiksa_records f ON f.full_name = ch.incident_number

# НОВОЕ:
LEFT JOIN fixations f ON f.application_id = a.application_id
```

### Замена операторов:
```python
# СТАРОЕ:
WHERE operator_name = 'Иванов'

# НОВОЕ:
WHERE o.operator_id = (SELECT operator_id FROM operators WHERE operator_name = 'Иванов')
```

## 📝 РЕКОМЕНДАЦИИ

1. **Использовать новые файлы:**
   - `db_schema.py` для создания БД
   - `db_import.py` для импорта данных
   - `db_reports.py` для создания отчетов
   - `db_manager.py` для управления

2. **Обновить боты:**
   - Перейти на представления (views)
   - Убрать прямые запросы к старым таблицам

3. **Архивировать старое:**
   - Все файлы перемещены в `archive/old_structure/`
   - При необходимости можно восстановить

4. **Тестировать:**
   - Проверить работу ботов после обновления
   - Проверить создание отчетов
   - Проверить логи

---

**Статус:** В процессе миграции
**Приоритет:** Высокий
""".format(date=__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M'))
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f'\n📄 Отчет сохранен: {report_file}')

def main():
    """Главная функция"""
    print('\n')
    
    # Анализ
    deprecated, to_update, duplicates = analyze_project()
    
    # Спрашиваем подтверждение
    print('\n' + '=' * 80)
    choice = input('\n❓ Переместить устаревшие файлы в архив? (y/n): ').lower().strip()
    
    if choice == 'y':
        moved = move_to_archive()
        print(f'\n✅ Архивировано файлов: {moved}')
    else:
        print('\n❌ Перемещение отменено')
    
    # Создаем отчет
    create_compatibility_report()
    
    print('\n' + '=' * 80)
    print('✅ АНАЛИЗ ЗАВЕРШЕН')
    print('=' * 80)
    print('\n📌 Следующие шаги:')
    print('   1. Проверьте файл СОВМЕСТИМОСТЬ_БД.md')
    print('   2. Обновите файлы из списка "ТРЕБУЮТ ОБНОВЛЕНИЯ"')
    print('   3. Протестируйте систему с новой БД')
    print('=' * 80 + '\n')

if __name__ == '__main__':
    main()
