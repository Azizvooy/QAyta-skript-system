"""
Скрипт автоматического обновления данных каждый час
Запускается в фоновом режиме и обновляет данные из Google Sheets
"""

import schedule
import time
import subprocess
import logging
from datetime import datetime
import os
import sys

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data/hourly_update.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def update_data():
    """Обновление данных из Google Sheets"""
    try:
        logger.info("=" * 60)
        logger.info("🔄 Начало обновления данных")
        logger.info("=" * 60)
        
        # Запуск НОВОГО скрипта обновления (для новой структуры БД)
        result = subprocess.run(
            [sys.executable, 'scripts/data_collection/sheets_to_db_collector.py'],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            logger.info("✅ Данные успешно обновлены!")
            logger.info(f"Вывод: {result.stdout[:200]}")
        else:
            logger.error(f"❌ Ошибка при обновлении данных!")
            logger.error(f"Код ошибки: {result.returncode}")
            logger.error(f"Вывод: {result.stderr[:500]}")
            
    except Exception as e:
        logger.error(f"❌ Критическая ошибка: {e}", exc_info=True)
    
    logger.info(f"⏰ Следующее обновление в {datetime.now().hour + 1}:00")
    logger.info("-" * 60)

def main():
    """Основная функция запуска планировщика"""
    logger.info("=" * 60)
    logger.info("🚀 ЗАПУСК АВТОМАТИЧЕСКОГО ОБНОВЛЕНИЯ")
    logger.info("=" * 60)
    logger.info("⏰ Расписание: каждый час")
    logger.info(f"📂 Рабочая директория: {os.getcwd()}")
    logger.info("=" * 60)
    
    # Сразу выполняем первое обновление
    logger.info("🔄 Выполнение первичного обновления...")
    update_data()
    
    # Планируем обновления каждый час
    schedule.every().hour.at(":00").do(update_data)
    
    logger.info("✅ Планировщик запущен и работает")
    logger.info("💡 Для остановки нажмите Ctrl+C")
    logger.info("=" * 60)
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Проверяем каждую минуту
            
    except KeyboardInterrupt:
        logger.info("\n" + "=" * 60)
        logger.info("🛑 Планировщик остановлен пользователем")
        logger.info("=" * 60)
    except Exception as e:
        logger.error(f"❌ Критическая ошибка: {e}", exc_info=True)

if __name__ == "__main__":
    main()
