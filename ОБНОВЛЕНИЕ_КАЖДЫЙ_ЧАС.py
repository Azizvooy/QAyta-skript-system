#!/usr/bin/env python3
"""
Ежечасное обновление данных из Google Sheets
Запускается автоматически каждый час даже если ПК спал

Работает в фоне и логирует все действия
"""
import time
import subprocess
import sys
import logging
from pathlib import Path
from datetime import datetime

# Директории
BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / 'scripts'
COLLECTOR_SCRIPT = SCRIPTS_DIR / 'data_collection' / 'improved_collector.py'
ANALYTICS_SCRIPT = SCRIPTS_DIR / 'analytics' / 'analytics_reports.py'
LOGS_DIR = BASE_DIR / 'logs'
LOG_FILE = LOGS_DIR / 'update_hourly.log'

# Создаем папку логов
LOGS_DIR.mkdir(exist_ok=True)

# Конфигурация логирования
logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def run_command(script_path, name):
    """Запустить скрипт и получить результат"""
    logger.info(f"Начинаю {name}...")
    
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            timeout=600  # 10 минут максимум
        )
        
        if result.returncode == 0:
            logger.info(f"✅ {name} успешно выполнено")
            
            # Парсим вывод для получения статистики
            if result.stdout:
                lines = result.stdout.split('\n')
                for line in lines[-10:]:  # Последние 10 строк
                    if line.strip():
                        logger.info(f"   {line.strip()}")
            return True
        else:
            logger.error(f"❌ Ошибка при {name}")
            if result.stderr:
                logger.error(f"   {result.stderr[:500]}")
            return False
            
    except subprocess.TimeoutExpired:
        logger.error(f"⏱ Превышено время ожидания для {name} (10 минут)")
        return False
    except Exception as e:
        logger.error(f"❌ Ошибка: {str(e)}", exc_info=True)
        return False

def update_data():
    """Полное обновление данных"""
    logger.info("=" * 80)
    logger.info("НАЧАЛО ЕЖЕЧАСНОГО ОБНОВЛЕНИЯ")
    logger.info("=" * 80)
    
    # 1. Обновляем данные из Google Sheets
    if COLLECTOR_SCRIPT.exists():
        success1 = run_command(COLLECTOR_SCRIPT, "обновление данных из Google Sheets")
    else:
        logger.warning(f"Скрипт не найден: {COLLECTOR_SCRIPT}")
        success1 = False
    
    time.sleep(2)
    
    # 2. Генерируем отчеты
    if ANALYTICS_SCRIPT.exists():
        success2 = run_command(ANALYTICS_SCRIPT, "генерация отчетов")
    else:
        logger.warning(f"Скрипт не найден: {ANALYTICS_SCRIPT}")
        success2 = False
    
    logger.info("=" * 80)
    if success1 or success2:
        logger.info("✅ ОБНОВЛЕНИЕ ЗАВЕРШЕНО УСПЕШНО")
    else:
        logger.warning("⚠️  ОБНОВЛЕНИЕ ЗАВЕРШЕНО С ОШИБКАМИ")
    logger.info("=" * 80)
    logger.info("")

def main():
    """Основной цикл - обновление каждый час"""
    logger.info("🟢 Сервис автоматического обновления запущен")
    logger.info("   Бот будет обновлять данные каждый час")
    logger.info("   Даже если компьютер в режиме сна, обновится при пробуждении")
    logger.info("   Логи: " + str(LOG_FILE))
    logger.info("")
    
    # Первое обновление сразу при запуске
    update_data()
    
    # Далее каждый час
    while True:
        try:
            logger.info(f"⏰ Следующее обновление через 1 час ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
            # Ждем 1 час (3600 секунд)
            time.sleep(3600)
            update_data()
        except KeyboardInterrupt:
            logger.info("\n⛔ Сервис остановлен пользователем")
            break
        except Exception as e:
            logger.error(f"❌ Неожиданная ошибка: {e}", exc_info=True)
            logger.info("⏰ Повторю попытку через 1 минуту...")
            time.sleep(60)

def test_mode():
    """Режим тестирования - обновление один раз"""
    logger.info("📝 РЕЖИМ ТЕСТИРОВАНИЯ - обновление один раз")
    logger.info("")
    update_data()
    logger.info("✅ Тестирование завершено")
    sys.exit(0)

if __name__ == '__main__':
    # Если передан аргумент 'test' - запустить один раз для тестирования
    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        test_mode()
    else:
        # Нормальный режим - фоновое обновление каждый час
        try:
            main()
        except Exception as e:
            logger.error(f"КРИТИЧЕСКАЯ ОШИБКА: {e}", exc_info=True)
            sys.exit(1)
