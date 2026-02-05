"""
Фоновый сервис для автоматического сбора данных и генерации отчетов
Запускается один раз и работает в фоне
"""

import schedule
import time
import subprocess
import sys
from pathlib import Path
from datetime import datetime
import logging

# Настройка логирования
log_dir = Path(__file__).parent.parent.parent / "output" / "logs"
log_dir.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / f"service_{datetime.now().strftime('%Y%m%d')}.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Пути к скриптам
BASE_DIR = Path(__file__).parent.parent.parent
COLLECTOR_SCRIPT = BASE_DIR / "scripts" / "data_collection" / "sheets_to_db_collector.py"
ANALYTICS_SCRIPT = BASE_DIR / "scripts" / "automation" / "auto_analytics.py"
APPLICATIONS_SCRIPT = BASE_DIR / "scripts" / "automation" / "process_applications.py"

def run_data_collection():
    """Запуск сбора данных"""
    logger.info("=" * 80)
    logger.info("🚀 ЗАПУСК СБОРА ДАННЫХ")
    logger.info("=" * 80)
    try:
        result = subprocess.run(
            [sys.executable, str(COLLECTOR_SCRIPT)],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            logger.info("✅ Сбор данных завершен успешно")
            logger.info(result.stdout)
        else:
            logger.error(f"❌ Ошибка при сборе данных: {result.stderr}")
            
    except Exception as e:
        logger.error(f"❌ Критическая ошибка при сборе данных: {e}")

def run_analytics():
    """Запуск аналитики и генерации отчетов"""
    logger.info("=" * 80)
    logger.info("📊 ЗАПУСК АНАЛИТИКИ И ОТЧЕТОВ")
    logger.info("=" * 80)
    try:
        result = subprocess.run(
            [sys.executable, str(ANALYTICS_SCRIPT)],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            logger.info("✅ Аналитика завершена успешно")
            logger.info(result.stdout)
        else:
            logger.error(f"❌ Ошибка при генерации отчетов: {result.stderr}")
            
    except Exception as e:
        logger.error(f"❌ Критическая ошибка при генерации отчетов: {e}")

def process_applications():
    """Обработка загруженных файлов заявок"""
    logger.info("=" * 80)
    logger.info("📂 ОБРАБОТКА ЗАЯВОК")
    logger.info("=" * 80)
    try:
        result = subprocess.run(
            [sys.executable, str(APPLICATIONS_SCRIPT)],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            logger.info("✅ Обработка заявок завершена")
            logger.info(result.stdout)
        else:
            logger.error(f"❌ Ошибка при обработке заявок: {result.stderr}")
            
    except Exception as e:
        logger.error(f"❌ Критическая ошибка при обработке заявок: {e}")

def daily_job():
    """Ежедневная задача: сбор данных + аналитика"""
    logger.info("\n" + "=" * 80)
    logger.info(f"📅 ЕЖЕДНЕВНАЯ ЗАДАЧА - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 80 + "\n")
    
    # 1. Сбор данных
    run_data_collection()
    
    # Пауза между задачами
    time.sleep(5)
    
    # 2. Генерация отчетов
    run_analytics()

def hourly_job():
    """Ежечасная задача: обработка новых заявок"""
    process_applications()

def main():
    """Главная функция фонового сервиса"""
    logger.info("\n" + "=" * 80)
    logger.info("🤖 ФОНОВЫЙ СЕРВИС ЗАПУЩЕН")
    logger.info("=" * 80)
    logger.info(f"⏰ Время запуска: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("")
    logger.info("📋 РАСПИСАНИЕ:")
    logger.info("  • Сбор данных и отчеты: Ежедневно в 09:00")
    logger.info("  • Обработка заявок: Каждый час")
    logger.info("")
    logger.info("❌ Для остановки нажмите Ctrl+C")
    logger.info("=" * 80 + "\n")
    
    # Настройка расписания
    schedule.every().day.at("09:00").do(daily_job)  # Ежедневно в 9:00
    schedule.every().hour.do(hourly_job)  # Каждый час
    
    # Запустить сразу при старте
    logger.info("🚀 Запуск первоначального сбора данных...")
    daily_job()
    
    # Бесконечный цикл
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Проверка каждую минуту
            
    except KeyboardInterrupt:
        logger.info("\n" + "=" * 80)
        logger.info("⛔ ФОНОВЫЙ СЕРВИС ОСТАНОВЛЕН ПОЛЬЗОВАТЕЛЕМ")
        logger.info("=" * 80)
        sys.exit(0)

if __name__ == "__main__":
    main()
