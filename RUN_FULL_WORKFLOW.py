#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=================================================================
ПОЛНЫЙ WORKFLOW: ИМПОРТ → ОБРАБОТКА → ОТЧЕТЫ
=================================================================
Этот скрипт выполняет полный цикл обработки данных:
1. Импорт данных из Google Sheets в PostgreSQL
2. Обработка и очистка данных в БД  
3. Генерация отчетов с отрицательными отзывами в Excel
4. Организация результатов в папках по месяцам и категориям

Использование:
    python RUN_FULL_WORKFLOW.py
    
Результаты сохраняются в папке: reports/{МЕСЯЦ_ГОД}/
=================================================================
"""

import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

# Определяем рабочую директорию
BASE_DIR = Path(__file__).parent
SCRIPTS_DIR = BASE_DIR / 'scripts'

def print_section(title):
    """Вывести раздел"""
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}\n")

def run_command(script_path, description):
    """Запустить Python скрипт"""
    print(f"⏱️  {description}...")
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(BASE_DIR),
            capture_output=False,
            text=True
        )
        if result.returncode == 0:
            print(f"✅ {description} - УСПЕШНО\n")
            return True
        else:
            print(f"❌ {description} - ОШИБКА\n")
            return False
    except Exception as e:
        print(f"❌ Ошибка при выполнении {description}: {e}\n")
        return False

def main():
    """Главная функция"""
    start_time = datetime.now()
    
    print_section("QAYTA СИСТЕМА ОБРАБОТКИ ДАННЫХ")
    print(f"Дата/время запуска: {start_time.strftime('%d.%m.%Y %H:%M:%S')}")
    
    # 1. ИМПОРТ ДАННЫХ
    print_section("ШАГ 1: ИМПОРТ ДАННЫХ ИЗ GOOGLE SHEETS")
    
    import_script = SCRIPTS_DIR / 'import' / 'import_from_sheets.py'
    if import_script.exists():
        if not run_command(import_script, "Импорт всех данных из Google Sheets"):
            print("⚠️  Импорт завершен с ошибкой, продолжаем обработку...")
    else:
        print(f"⚠️  Скрипт не найден: {import_script}")
    
    # 2. ОБРАБОТКА ДАННЫХ
    print_section("ШАГ 2: ОБРАБОТКА ДАННЫХ В POSTGRESQL")
    
    process_script = SCRIPTS_DIR / 'processing' / 'process_data.py'
    if process_script.exists():
        if not run_command(process_script, "Обработка и создание таблиц в БД"):
            print("⚠️  Обработка завершена с ошибкой, продолжаем генерацию отчетов...")
    else:
        print(f"⚠️  Скрипт не найден: {process_script}")
    
    # 3. ГЕНЕРАЦИЯ ОТЧЕТОВ
    print_section("ШАГ 3: ГЕНЕРАЦИЯ ОТЧЕТОВ С ОТРИЦАТЕЛЬНЫМИ ОТЗЫВАМИ")
    
    reports_script = SCRIPTS_DIR / 'reports' / 'generate_reports.py'
    if reports_script.exists():
        if not run_command(reports_script, "Генерация Excel отчетов"):
            print("⚠️  Генерация отчетов завершена с ошибкой")
    else:
        print(f"⚠️  Скрипт не найден: {reports_script}")
    
    # Финальная информация
    elapsed = datetime.now() - start_time
    print_section("РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ")
    
    print(f"⏱️  Общее время: {elapsed}")
    print(f"📁 Отчеты сохранены в: reports/{datetime.now().strftime('%Y-%m')}/")
    print(f"📊 БД таблицы: detailed_reports, negative_complaints")
    print("\n✅ WORKFLOW ЗАВЕРШЕН\n")

if __name__ == '__main__':
    main()
