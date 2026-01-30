"""
Интерактивный Telegram бот для управления отчетами
Отправляйте команды и получайте нужные отчеты
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

# Создаем директорию для логов
LOG_DIR = BASE_DIR / 'logs'
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / 'telegram_bot.log'

# Настройка логирования (файл + консоль)
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Telegram токен
TOKEN = '8141079204:AAErNrjLqTu4Vj1_7VS2kjGFKcR3lU9L9N4'
CHAT_ID = '2012682567'

# =============================================================================
# БАЗА ДАННЫХ
# =============================================================================

def get_db():
    """Подключение к базе"""
    return sqlite3.connect(DB_PATH)

def get_stats():
    """Текущая статистика"""
    conn = get_db()
    c = conn.cursor()
    
    stats = {}
    
    # Количество записей в новых таблицах
    c.execute('SELECT COUNT(*) FROM fixations')
    stats['fiksa'] = c.fetchone()[0]
    
    c.execute('SELECT COUNT(*) FROM applications')
    stats['calls'] = c.fetchone()[0]
    
    stats['apps'] = stats['calls']  # То же самое
    
    # Сегодняшние данные из статистики
    c.execute('SELECT COUNT(*) FROM daily_statistics WHERE stat_date = date("now")')
    stats['operators_today'] = c.fetchone()[0]
    
    # Последнее обновление фиксаций
    c.execute('SELECT MAX(created_at) FROM fixations')
    last_update = c.fetchone()[0]
    stats['last_update'] = last_update if last_update else 'Нет данных'
    
    conn.close()
    return stats

# =============================================================================
# КОМАНДЫ
# =============================================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Команда /start - главное меню"""
    keyboard = [
        [InlineKeyboardButton("📊 Статистика", callback_data='stats')],
        [InlineKeyboardButton("👥 Операторы", callback_data='operators')],
        [InlineKeyboardButton("🚨 Фидбэки служб", callback_data='feedback')],
        [InlineKeyboardButton("📞 Отчет по службе", callback_data='service_select')],
        [InlineKeyboardButton("📋 Все отчеты", callback_data='all_reports')],
        [InlineKeyboardButton("🔄 Обновить данные", callback_data='update_data')],
        [InlineKeyboardButton("ℹ️ Помощь", callback_data='help')],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    text = """
🤖 <b>Бот управления отчетами FIKSA</b>

Выберите действие:
• <b>Статистика</b> - текущее состояние базы
• <b>Операторы</b> - статистика по операторам
• <b>Фидбэки служб</b> - анализ ответов служб
• <b>Отчет по службе</b> - выбрать конкретную службу
• <b>Все отчеты</b> - генерация всех отчетов
• <b>Обновить данные</b> - синхронизация с Google Sheets
• <b>Помощь</b> - список команд

Или отправьте текстовый запрос!
    """
    
    if update.callback_query:
        await update.callback_query.edit_message_text(text, parse_mode='HTML', reply_markup=reply_markup)
    else:
        await update.message.reply_text(text, parse_mode='HTML', reply_markup=reply_markup)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Команда /help - помощь"""
    text = """
📖 <b>Доступные команды:</b>

<b>Основные:</b>
/start - Главное меню
/stats - Текущая статистика
/operators - Статистика операторов
/feedback - Фидбэки служб
/update - Обновить данные FIKSA

<b>Отчеты:</b>
/service_102 - Отчет по службе 102
/service_103 - Отчет по службе 103
/service_104 - Отчет по службе 104
/all_reports - Все отчеты по службам (61 файл)

<b>Текстовые запросы:</b>
Можете написать:
• "статистика за сегодня"
• "сколько звонков по 102"
• "топ операторов"
• "проблемные случаи"
• "обновить данные"

Бот поймет и выполнит!
    """
    await update.message.reply_text(text, parse_mode='HTML')

# =============================================================================
# СТАТИСТИКА
# =============================================================================

async def show_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показать статистику"""
    query = update.callback_query
    await query.answer()
    
    stats = get_stats()
    
    text = f"""
📊 <b>ТЕКУЩАЯ СТАТИСТИКА</b>

<b>База данных:</b>
• FIKSA записей: {stats['fiksa']:,}
• История 112: {stats['calls']:,}
• Заявки: {stats['apps']:,}

<b>Сегодня:</b>
• Операторов: {stats['operators_today']}

<b>Последнее обновление:</b>
{stats['last_update']}

<i>Обновлено: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</i>
    """
    
    keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data='start')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await query.edit_message_text(text, parse_mode='HTML', reply_markup=reply_markup)

# =============================================================================
# ОПЕРАТОРЫ
# =============================================================================

async def show_operators(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Статистика операторов"""
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text("⏳ Генерирую отчет по операторам...")
    
    # Запускаем генерацию отчета
    try:
        result = subprocess.run(
            [sys.executable, str(BASE_DIR / 'scripts' / 'analysis' / 'analytics_reports.py')],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        # Находим файл отчета
        report_file = REPORTS_DIR / 'analytics' / 'operator_performance_report.xlsx'
        
        if report_file.exists():
            # Читаем для краткой статистики
            df = pd.read_excel(report_file)
            
            top_5 = df.nlargest(5, 'Всего')[['Оператор', 'Всего', '% Успешных']].to_string(index=False)
            
            text = f"""
✅ <b>Отчет по операторам готов!</b>

<b>ТОП-5 операторов:</b>
<pre>{top_5}</pre>

Файл отправлен ниже 👇
            """
            
            await query.edit_message_text(text, parse_mode='HTML')
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=open(report_file, 'rb'),
                filename='Статистика_операторов.xlsx'
            )
        else:
            await query.edit_message_text("❌ Ошибка: файл отчета не найден")
            
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка: {e}")

# =============================================================================
# ФИДБЭКИ СЛУЖБ
# =============================================================================

async def show_feedback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Фидбэки служб"""
    query = update.callback_query
    await query.answer()
    
    conn = get_db()
    
    # Статистика по службам из представления
    df = pd.read_sql_query('''
        SELECT 
            service_name as service,
            COUNT(*) as total,
            SUM(CASE WHEN fixation_status LIKE '%не%' OR fixation_status LIKE '%отказ%' THEN 1 ELSE 0 END) as problems
        FROM v_fixations_full
        WHERE service_name IN ('102', '103', '104')
        GROUP BY service_name
    ''', conn)
    
    conn.close()
    
    text = "<b>🚨 ФИДБЭКИ СЛУЖБ</b>\n\n"
    
    for _, row in df.iterrows():
        service = row['service']
        total = row['total']
        problems = row['problems']
        percent = (problems / total * 100) if total > 0 else 0
        
        text += f"<b>Служба {service}:</b>\n"
        text += f"  Всего: {total:,}\n"
        text += f"  Проблемных: {problems:,} ({percent:.1f}%)\n\n"
    
    keyboard = [
        [InlineKeyboardButton("📥 Скачать полный отчет", callback_data='feedback_report')],
        [InlineKeyboardButton("◀️ Назад", callback_data='start')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await query.edit_message_text(text, parse_mode='HTML', reply_markup=reply_markup)

async def generate_feedback_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Генерация полного отчета по фидбэкам"""
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text("⏳ Генерирую отчет по фидбэкам...")
    
    try:
        subprocess.run(
            [sys.executable, str(BASE_DIR / 'scripts' / 'analysis' / 'analytics_reports.py')],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        report_file = REPORTS_DIR / 'analytics' / 'service_feedback_report.xlsx'
        
        if report_file.exists():
            await query.edit_message_text("✅ Отчет готов! Отправляю...")
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=open(report_file, 'rb'),
                filename='Фидбэки_служб.xlsx',
                caption='📊 Полный отчет по фидбэкам служб (3 листа)'
            )
        else:
            await query.edit_message_text("❌ Файл не найден")
            
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка: {e}")

# =============================================================================
# ОТЧЕТЫ ПО СЛУЖБАМ
# =============================================================================

async def service_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Выбор службы для отчета"""
    query = update.callback_query
    await query.answer()
    
    keyboard = [
        [InlineKeyboardButton("🚔 Милиция (102)", callback_data='service_102')],
        [InlineKeyboardButton("🚑 Скорая (103)", callback_data='service_103')],
        [InlineKeyboardButton("🚒 Пожарная (104)", callback_data='service_104')],
        [InlineKeyboardButton("📋 Все службы (61 файл)", callback_data='all_services')],
        [InlineKeyboardButton("◀️ Назад", callback_data='start')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await query.edit_message_text(
        "Выберите службу для генерации отчета:",
        reply_markup=reply_markup
    )

async def generate_service_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Генерация отчета по конкретной службе"""
    query = update.callback_query
    service = query.data.split('_')[1]  # service_102 -> 102
    
    await query.answer()
    await query.edit_message_text(f"⏳ Генерирую отчет по службе {service}...")
    
    try:
        # Запускаем генерацию отчетов
        subprocess.run(
            [sys.executable, str(BASE_DIR / 'scripts' / 'analysis' / 'service_reports.py')],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        # Ищем файлы службы
        services_dir = REPORTS_DIR / 'services'
        service_files = list(services_dir.glob(f'*{service}*.xlsx'))
        
        if service_files:
            await query.edit_message_text(f"✅ Найдено {len(service_files)} файлов. Отправляю...")
            
            for file in service_files[:5]:  # Максимум 5 файлов
                await context.bot.send_document(
                    chat_id=query.message.chat_id,
                    document=open(file, 'rb'),
                    filename=file.name
                )
        else:
            await query.edit_message_text(f"❌ Файлы для службы {service} не найдены")
            
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка: {e}")

async def generate_all_reports(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Генерация всех отчетов"""
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text("⏳ Генерирую ВСЕ отчеты (это может занять время)...")
    
    try:
        subprocess.run(
            [sys.executable, str(BASE_DIR / 'scripts' / 'analysis' / 'service_reports.py')],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            timeout=300  # 5 минут максимум
        )
        
        await query.edit_message_text("✅ Все отчеты сгенерированы и отправлены!")
        
    except subprocess.TimeoutExpired:
        await query.edit_message_text("⚠️ Генерация заняла слишком много времени, но процесс запущен")
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка: {e}")

# =============================================================================
# ОБНОВЛЕНИЕ ДАННЫХ
# =============================================================================

async def update_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обновление данных из Google Sheets"""
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text("⏳ Синхронизация с Google Sheets...")
    
    try:
        result = subprocess.run(
            [sys.executable, str(BASE_DIR / 'scripts' / 'data_collection' / 'improved_collector.py')],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            timeout=180
        )
        
        # Парсим вывод
        if 'ИТОГО собрано:' in result.stdout:
            lines = result.stdout.split('\n')
            for line in lines:
                if 'ИТОГО собрано:' in line:
                    count = line.split(':')[1].strip()
                    await query.edit_message_text(f"✅ Обновление завершено!\n\nСобрано записей: {count}")
                    return
        
        await query.edit_message_text("✅ Обновление завершено!")
        
    except subprocess.TimeoutExpired:
        await query.edit_message_text("⚠️ Обновление заняло много времени, процесс в фоне")
    except Exception as e:
        await query.edit_message_text(f"❌ Ошибка: {e}")

# =============================================================================
# ТЕКСТОВЫЕ ЗАПРОСЫ
# =============================================================================

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка текстовых запросов"""
    text = update.message.text.lower()
    
    # Статистика
    if any(word in text for word in ['статистика', 'сколько', 'данных', 'база']):
        stats = get_stats()
        response = f"""
📊 <b>Статистика:</b>

FIKSA: {stats['fiksa']:,} записей
История 112: {stats['calls']:,} звонков
Заявки: {stats['apps']:,}
Операторов сегодня: {stats['operators_today']}

Обновлено: {stats['last_update']}
        """
        await update.message.reply_text(response, parse_mode='HTML')
    
    # Операторы
    elif any(word in text for word in ['оператор', 'топ', 'лучш', 'работник']):
        await update.message.reply_text("⏳ Генерирую статистику операторов...")
        
        report_file = REPORTS_DIR / 'analytics' / 'operator_performance_report.xlsx'
        if report_file.exists():
            await update.message.reply_document(
                document=open(report_file, 'rb'),
                filename='Операторы.xlsx'
            )
        else:
            await update.message.reply_text("Сначала нужно обновить данные")
    
    # Обновление
    elif any(word in text for word in ['обнов', 'синхр', 'загруз']):
        await update.message.reply_text("⏳ Запускаю обновление...")
        subprocess.Popen(
            [sys.executable, str(BASE_DIR / 'scripts' / 'data_collection' / 'improved_collector.py')],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        await update.message.reply_text("✅ Обновление запущено в фоне")
    
    # Не понял
    else:
        await update.message.reply_text(
            "Не совсем понял 🤔\n\nИспользуйте /start для меню или /help для списка команд"
        )

# =============================================================================
# ОБРАБОТЧИКИ КНОПОК
# =============================================================================

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка нажатий кнопок"""
    query = update.callback_query
    
    handlers = {
        'start': start,
        'stats': show_stats,
        'operators': show_operators,
        'feedback': show_feedback,
        'feedback_report': generate_feedback_report,
        'service_select': service_select,
        'service_102': generate_service_report,
        'service_103': generate_service_report,
        'service_104': generate_service_report,
        'all_services': generate_all_reports,
        'all_reports': generate_all_reports,
        'update_data': update_data,
        'help': help_command,
    }
    
    handler = handlers.get(query.data)
    if handler:
        await handler(update, context)

# =============================================================================
# MAIN
# =============================================================================

def main():
    """Запуск бота"""
    startup_msg = f"\n{'=' * 80}\nИНТЕРАКТИВНЫЙ TELEGRAM БОТ\n{'=' * 80}\nЗапуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nЛог-файл: {LOG_FILE}\nБот готов к приему команд...\n{'=' * 80}\n"
    
    print(startup_msg)
    logger.info("=== ЗАПУСК БОТА ===")
    logger.info(f"База данных: {DB_PATH}")
    logger.info(f"Лог-файл: {LOG_FILE}")
    
    try:
        app = Application.builder().token(TOKEN).build()
        logger.info("Приложение создано")
        
        # Команды
        app.add_handler(CommandHandler('start', start))
        app.add_handler(CommandHandler('help', help_command))
        app.add_handler(CommandHandler('stats', show_stats))
        
        # Кнопки
        app.add_handler(CallbackQueryHandler(button_handler))
        
        # Текстовые сообщения
        app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
        
        logger.info("Обработчики зарегистрированы")
        logger.info("=== БОТ ГОТОВ К РАБОТЕ ===")
        
        # Запуск
        app.run_polling(allowed_updates=Update.ALL_TYPES)
        
    except KeyboardInterrupt:
        logger.info("=== ОСТАНОВКА БОТА (Ctrl+C) ===")
        print("\n\nБот остановлен пользователем")
    except Exception as e:
        logger.error(f"=== КРИТИЧЕСКАЯ ОШИБКА === {e}", exc_info=True)
        print(f"\n\n[ОШИБКА] {e}")
        raise
    finally:
        logger.info("=== ЗАВЕРШЕНИЕ РАБОТЫ БОТА ===")
        print("\n\nБот завершил работу")

if __name__ == '__main__':
    main()
