"""
Тестовый запуск автоматизации
"""
import sys
from pathlib import Path

# Добавляем путь к проекту
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from scripts.automation.master_service import show_current_stats, hourly_tasks

print('=' * 80)
print('ТЕСТ АВТОМАТИЗАЦИИ')
print('=' * 80)

# 1. Показываем текущую статистику
print('\n1. ТЕКУЩАЯ СТАТИСТИКА:')
show_current_stats()

# 2. Запускаем быстрый сбор (как при hourly)
print('\n2. БЫСТРАЯ СИНХРОНИЗАЦИЯ:')
hourly_tasks()

print('\n' + '=' * 80)
print('[OK] ТЕСТ ЗАВЕРШЕН')
print('=' * 80)
