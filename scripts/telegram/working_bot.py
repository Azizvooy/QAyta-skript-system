"""
Рабочий Telegram бот с автоматическим обновлением данных и отчетами
"""
import asyncio
import sqlite3
import subprocess
import sys
import logging
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler, 
    MessageHandler, filters, ContextTypes
)

# Пути
BASE_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(BASE_DIR))

DB_PATH = BASE_DIR / 'data' / 'fiksa_database.db'
REPORTS_DIR = BASE_DIR / 'reports'
ANALYTICS_DIR = REPORTS_DIR / 'analytics'
SERVICES_DIR = REPORTS_DIR / 'services'

# Создаем директории
for dir_path in [REPORTS_DIR, ANALYTICS_DIR, SERVICES_DIR]:
    dir_path.mkdir(exist_ok=True, parents=True)

# Логи
LOG_DIR = BASE_DIR / 'logs'
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / 'telegram_bot.log'

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Telegram
TOKEN = '8141079204:AAErNrjLqTu4Vj1_7VS2kjGFKcR3lU9L9N4'
CHAT_ID = 2012682567

# =============================================================================
# БАЗА ДАННЫХ
# =============================================================================

def get_db():
    """Подключение к БД"""
    return sqlite3.connect(DB_PATH)

def get_stats():
    """Статистика по БД"""
    conn = get_db()
    stats = {}
    
    try:
        # Фиксации
        df = pd.read_sql_query('SELECT COUNT(*) as cnt FROM fixations', conn)
        stats['fiksa'] = df['cnt'].iloc[0]
        
        # Заявки (раньше call_history_112)
        df = pd.read_sql_query('SELECT COUNT(*) as cnt FROM applications', conn)
        stats['calls'] = df['cnt'].iloc[0]
        
        # Заявки (то же самое)
        stats['apps'] = stats['calls']
        
        # Последнее обновление
        df = pd.read_sql_query('SELECT MAX(created_at) as last_dt FROM fixations', conn)
        last_dt = df['last_dt'].iloc[0]
        stats['last_update'] = last_dt if last_dt else 'Нет данных'
        
        # Сегодняшние данные
        df = pd.read_sql_query(
            'SELECT COUNT(*) as cnt FROM fixations WHERE DATE(created_at) = DATE("now")',
            conn
        )
        stats['today'] = df['cnt'].iloc[0]
        
    except Exception as e:
        logger.error(f"Ошибка получения статистики: {e}")
        stats = {'fiksa': 0, 'calls': 0, 'apps': 0, 'last_update': 'Ошибка', 'today': 0}
    finally:
        conn.close()
    
    return stats

# =============================================================================
# ОБНОВЛЕНИЕ ДАННЫХ
# =============================================================================

async def update_fiksa_data(update: Update = None, send_message=True):
    """Обновление данных из Google Sheets"""
    logger.info("Начало обновления данных FIKSA")
    
    collector_script = BASE_DIR / 'scripts' / 'data_collection' / 'improved_collector.py'
    
    if not collector_script.exists():
        msg = f"❌ Скрипт сбора не найден: {collector_script}"
        logger.error(msg)
        if update and send_message:
            await update.effective_message.reply_text(msg)
        return False
    
    if update and send_message:
        await update.effective_message.reply_text("🔄 Начинаю обновление данных из FIKSA...")
    
    try:
        # Запускаем сбор данных
        result = subprocess.run(
            [sys.executable, str(collector_script)],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            timeout=600  # 10 минут максимум
        )
        
        if result.returncode == 0:
            # Парсим вывод для получения статистики
            lines = result.stdout.split('\n')
            collected = 0
            for line in lines:
                if 'Собрано записей:' in line:
                    try:
                        collected = int(line.split(':')[1].strip())
                    except:
                        pass
            
            msg = f"✅ Данные обновлены!\n📊 Собрано записей: {collected}"
            logger.info(msg)
            
            if update and send_message:
                await update.effective_message.reply_text(msg)
            return True
        else:
            msg = f"❌ Ошибка обновления:\n{result.stderr[:500]}"
            logger.error(msg)
            
            if update and send_message:
                await update.effective_message.reply_text(msg)
            return False
            
    except subprocess.TimeoutExpired:
        msg = "⏱ Превышено время ожидания (10 минут)"
        logger.error(msg)
        if update and send_message:
            await update.effective_message.reply_text(msg)
        return False
    except Exception as e:
        msg = f"❌ Ошибка: {str(e)}"
        logger.error(msg, exc_info=True)
        if update and send_message:
            await update.effective_message.reply_text(msg)
        return False

# =============================================================================
# ГЕНЕРАЦИЯ ОТЧЕТОВ
# =============================================================================

async def generate_reports(update: Update = None, send_message=True):
    """Генерация всех отчетов"""
    logger.info("Начало генерации отчетов")
    
    analytics_script = BASE_DIR / 'scripts' / 'analytics' / 'analytics_reports.py'
    
    if not analytics_script.exists():
        msg = f"❌ Скрипт аналитики не найден: {analytics_script}"
        logger.error(msg)
        if update and send_message:
            await update.effective_message.reply_text(msg)
        return False
    
    if update and send_message:
        await update.effective_message.reply_text("📊 Генерирую отчеты...")
    
    try:
        result = subprocess.run(
            [sys.executable, str(analytics_script)],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            timeout=600
        )
        
        if result.returncode == 0:
            msg = "✅ Отчеты созданы!"
            logger.info(msg)
            
            if update and send_message:
                await update.effective_message.reply_text(msg)
            return True
        else:
            msg = f"❌ Ошибка генерации:\n{result.stderr[:500]}"
            logger.error(msg)
            
            if update and send_message:
                await update.effective_message.reply_text(msg)
            return False
            
    except Exception as e:
        msg = f"❌ Ошибка: {str(e)}"
        logger.error(msg, exc_info=True)
        if update and send_message:
            await update.effective_message.reply_text(msg)
        return False

async def full_update(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Полное обновление: данные + отчеты"""
    await update.effective_message.reply_text("🚀 Запускаю полное обновление...")
    
    # 1. Обновляем данные
    success = await update_fiksa_data(update, send_message=True)
    
    if success:
        # 2. Генерируем отчеты
        await asyncio.sleep(2)
        await generate_reports(update, send_message=True)
        
        # 3. Показываем статистику
        await asyncio.sleep(1)
        await show_stats(update, None)
    else:
        await update.effective_message.reply_text("❌ Обновление прервано из-за ошибок")

# =============================================================================
# ОТПРАВКА ОТЧЕТОВ
# =============================================================================

async def send_operator_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отправка отчета по операторам"""
    await update.effective_message.reply_text("📊 Готовлю отчет по операторам...")
    
    file_path = ANALYTICS_DIR / 'operator_stats.xlsx'
    
    if not file_path.exists():
        await generate_reports(update, send_message=False)
        await asyncio.sleep(2)
    
    if file_path.exists():
        await update.effective_message.reply_document(
            document=open(file_path, 'rb'),
            filename='Операторы_Статистика.xlsx',
            caption=f"📊 Статистика операторов\n🕐 {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )
    else:
        await update.effective_message.reply_text("❌ Файл не найден. Попробуйте /update")

async def send_feedback_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отправка отчета по фидбэкам"""
    await update.effective_message.reply_text("📋 Готовлю отчет по фидбэкам...")
    
    file_path = ANALYTICS_DIR / 'service_feedback.xlsx'
    
    if not file_path.exists():
        await generate_reports(update, send_message=False)
        await asyncio.sleep(2)
    
    if file_path.exists():
        await update.effective_message.reply_document(
            document=open(file_path, 'rb'),
            filename='Фидбэки_Служб.xlsx',
            caption=f"🚨 Фидбэки служб 102/103/104\n🕐 {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )
    else:
        await update.effective_message.reply_text("❌ Файл не найден. Попробуйте /update")

async def send_service_report(update: Update, context: ContextTypes.DEFAULT_TYPE, service_num: int):
    """Отправка отчета по конкретной службе"""
    await update.effective_message.reply_text(f"📞 Готовлю отчет по службе {service_num}...")
    
    file_path = SERVICES_DIR / f'service_{service_num}_detailed.xlsx'
    
    if not file_path.exists():
        await generate_reports(update, send_message=False)
        await asyncio.sleep(2)
    
    if file_path.exists():
        await update.effective_message.reply_document(
            document=open(file_path, 'rb'),
            filename=f'Служба_{service_num}_Детально.xlsx',
            caption=f"📞 Детальный отчет службы {service_num}\n🕐 {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )
    else:
        await update.effective_message.reply_text("❌ Файл не найден. Попробуйте /update")

# =============================================================================
# КОМАНДЫ БОТА
# =============================================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Главное меню"""
    keyboard = [
        [InlineKeyboardButton("📊 Статистика", callback_data='stats')],
        [InlineKeyboardButton("👥 Операторы", callback_data='operators')],
        [InlineKeyboardButton("🚨 Фидбэки", callback_data='feedback')],
        [InlineKeyboardButton("📞 Служба 102", callback_data='service_102'),
         InlineKeyboardButton("📞 Служба 103", callback_data='service_103')],
        [InlineKeyboardButton("📞 Служба 104", callback_data='service_104')],
        [InlineKeyboardButton("🔄 Обновить данные", callback_data='update_data')],
        [InlineKeyboardButton("📋 Обновить + Отчеты", callback_data='full_update')],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    text = """
🤖 <b>TELEGRAM БОТ FIKSA</b>

Выберите действие:

📊 <b>Статистика</b> - текущее состояние БД
👥 <b>Операторы</b> - отчет по операторам
🚨 <b>Фидбэки</b> - отчет по службам
📞 <b>Служба 102/103/104</b> - детальные отчеты

🔄 <b>Обновить данные</b> - собрать из Google Sheets
📋 <b>Обновить + Отчеты</b> - полный цикл

<i>Бот работает 24/7</i>
"""
    
    await update.effective_message.reply_text(
        text,
        parse_mode='HTML',
        reply_markup=reply_markup
    )

async def show_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показать статистику"""
    stats = get_stats()
    
    text = f"""
📊 <b>СТАТИСТИКА БД</b>

📁 FIKSA записи: <b>{stats['fiksa']:,}</b>
📞 Звонки 112: <b>{stats['calls']:,}</b>
📋 Заявки: <b>{stats['apps']:,}</b>

📅 Сегодня добавлено: <b>{stats['today']}</b>
🕐 Последнее обновление: <code>{stats['last_update']}</code>

<i>Обновлено: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</i>
"""
    
    keyboard = [[InlineKeyboardButton("🔙 Назад", callback_data='start')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.effective_message.reply_text(
        text,
        parse_mode='HTML',
        reply_markup=reply_markup
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Помощь"""
    text = """
📖 <b>СПРАВКА</b>

<b>Команды:</b>
/start - Главное меню
/stats - Статистика БД
/update - Обновить данные
/reports - Генерировать отчеты
/full - Полное обновление
/operators - Отчет по операторам
/feedback - Отчет по фидбэкам
/service102 - Служба 102
/service103 - Служба 103
/service104 - Служба 104

<b>Процесс работы:</b>
1️⃣ Обновить данные (/update)
2️⃣ Сгенерировать отчеты (/reports)
3️⃣ Получить нужный отчет

Или используйте /full для всего сразу
"""
    
    await update.effective_message.reply_text(text, parse_mode='HTML')

# =============================================================================
# ОБРАБОТЧИКИ
# =============================================================================

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка нажатий кнопок"""
    query = update.callback_query
    await query.answer()
    
    handlers = {
        'start': start,
        'stats': show_stats,
        'operators': send_operator_stats,
        'feedback': send_feedback_report,
        'service_102': lambda u, c: send_service_report(u, c, 102),
        'service_103': lambda u, c: send_service_report(u, c, 103),
        'service_104': lambda u, c: send_service_report(u, c, 104),
        'update_data': update_fiksa_data,
        'full_update': full_update,
    }
    
    handler = handlers.get(query.data)
    if handler:
        await handler(update, context)
    else:
        await query.message.reply_text(f"Неизвестная команда: {query.data}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка текстовых сообщений"""
    text = update.message.text.lower()
    
    if any(word in text for word in ['стат', 'данные', 'инфо']):
        await show_stats(update, context)
    elif any(word in text for word in ['обнов', 'загруз', 'собр']):
        await full_update(update)
    elif any(word in text for word in ['оператор', 'сотрудник']):
        await send_operator_stats(update, context)
    elif any(word in text for word in ['фидбэк', 'отзыв', 'служб']):
        await send_feedback_report(update, context)
    elif '102' in text:
        await send_service_report(update, context, 102)
    elif '103' in text:
        await send_service_report(update, context, 103)
    elif '104' in text:
        await send_service_report(update, context, 104)
    else:
        await start(update, context)

# =============================================================================
# ЗАПУСК
# =============================================================================

def main():
    """Запуск бота"""
    startup_msg = f"""
{'=' * 80}
🤖 TELEGRAM БОТ FIKSA
{'=' * 80}
Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
База данных: {DB_PATH}
Лог-файл: {LOG_FILE}
Token: {TOKEN[:20]}...
Chat ID: {CHAT_ID}
{'=' * 80}
"""
    
    print(startup_msg)
    logger.info("=== ЗАПУСК БОТА ===")
    logger.info(f"База: {DB_PATH}")
    logger.info(f"Отчеты: {REPORTS_DIR}")
    
    try:
        app = Application.builder().token(TOKEN).build()
        logger.info("Приложение создано")
        
        # Команды
        app.add_handler(CommandHandler('start', start))
        app.add_handler(CommandHandler('help', help_command))
        app.add_handler(CommandHandler('stats', show_stats))
        app.add_handler(CommandHandler('update', update_fiksa_data))
        app.add_handler(CommandHandler('reports', generate_reports))
        app.add_handler(CommandHandler('full', full_update))
        app.add_handler(CommandHandler('operators', send_operator_stats))
        app.add_handler(CommandHandler('feedback', send_feedback_report))
        app.add_handler(CommandHandler('service102', lambda u, c: send_service_report(u, c, 102)))
        app.add_handler(CommandHandler('service103', lambda u, c: send_service_report(u, c, 103)))
        app.add_handler(CommandHandler('service104', lambda u, c: send_service_report(u, c, 104)))
        
        # Кнопки и текст
        app.add_handler(CallbackQueryHandler(button_handler))
        app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
        
        logger.info("Обработчики зарегистрированы")
        logger.info("=== БОТ ГОТОВ ===")
        print("\n✅ БОТ РАБОТАЕТ\n")
        
        # Запуск
        app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)
        
    except KeyboardInterrupt:
        logger.info("=== ОСТАНОВКА (Ctrl+C) ===")
        print("\n\n⛔ Бот остановлен")
    except Exception as e:
        logger.error(f"=== ОШИБКА === {e}", exc_info=True)
        print(f"\n\n❌ ОШИБКА: {e}")
        raise
    finally:
        logger.info("=== ЗАВЕРШЕНИЕ ===")
        print("\n👋 Бот завершил работу\n")

if __name__ == '__main__':
    main()
