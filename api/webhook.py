"""
Webhook версия Telegram бота для развертывания на Vercel/serverless платформах
"""
import os
import logging
from pathlib import Path
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Получаем токен из переменных окружения
TOKEN = os.getenv('BOT_TOKEN', '8141079204:AAErNrjLqTu4Vj1_7VS2kjGFKcR3lU9L9N4')
WEBHOOK_URL = os.getenv('WEBHOOK_URL', '')  # Устанавливается автоматически на Vercel

# Пути (для локального использования, на Vercel будет временная папка)
BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = Path('/tmp/uploads')  # Временная папка на Vercel
UPLOADS_DIR.mkdir(exist_ok=True, parents=True)

# =============================================================================
# ОБРАБОТЧИКИ
# =============================================================================

async def start(update: Update, context):
    """Главное меню"""
    keyboard = [
        [InlineKeyboardButton("📥 Загрузить файл 112", callback_data='upload_info')],
        [InlineKeyboardButton("📋 Список файлов", callback_data='list_files')],
        [InlineKeyboardButton("📊 Статистика", callback_data='stats')],
        [InlineKeyboardButton("ℹ️ Помощь", callback_data='help')],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    text = """
🤖 <b>TELEGRAM БОТ QAyta</b>

📥 <b>Загрузить файл 112</b> - загрузить данные из системы 112
📋 <b>Список файлов</b> - просмотр загруженных файлов
📊 <b>Статистика</b> - информация о системе
ℹ️ <b>Помощь</b> - инструкции по использованию

<i>Webhook режим | Serverless</i>
"""
    
    await update.effective_message.reply_text(
        text,
        parse_mode='HTML',
        reply_markup=reply_markup
    )

async def handle_document(update: Update, context):
    """Обработка загруженных файлов"""
    document = update.message.document
    file_name = document.file_name
    
    # Проверяем расширение файла
    if not (file_name.endswith('.xlsx') or file_name.endswith('.xls') or file_name.endswith('.csv')):
        await update.message.reply_text(
            "❌ Поддерживаются только Excel (.xlsx, .xls) и CSV (.csv) файлы"
        )
        return
    
    await update.message.reply_text(f"📥 Загружаю файл: {file_name}...")
    
    try:
        # Скачиваем файл
        file = await context.bot.get_file(document.file_id)
        file_path = UPLOADS_DIR / file_name
        
        await file.download_to_drive(file_path)
        
        # Получаем размер файла
        file_size = file_path.stat().st_size / (1024 * 1024)  # В MB
        
        msg = f"""✅ <b>Файл успешно загружен!</b>

📁 Имя: <code>{file_name}</code>
📊 Размер: {file_size:.2f} MB

⚠️ <i>Примечание: На Vercel файлы хранятся временно.
Для постоянного хранения используйте облачное хранилище.</i>"""
        
        keyboard = [[InlineKeyboardButton("🔙 Главное меню", callback_data='start')]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.message.reply_text(msg, parse_mode='HTML', reply_markup=reply_markup)
        
        logger.info(f"Файл загружен: {file_name} ({file_size:.2f} MB)")
        
    except Exception as e:
        error_msg = f"❌ Ошибка при загрузке файла: {str(e)}"
        await update.message.reply_text(error_msg)
        logger.error(error_msg, exc_info=True)

async def list_files(update: Update, context):
    """Показать список файлов"""
    try:
        files = list(UPLOADS_DIR.glob('*.xlsx')) + list(UPLOADS_DIR.glob('*.xls')) + list(UPLOADS_DIR.glob('*.csv'))
        
        if not files:
            await update.effective_message.reply_text("📂 Нет загруженных файлов")
            return
        
        msg = "<b>📁 Загруженные файлы:</b>\n\n"
        for i, file in enumerate(sorted(files, key=lambda x: x.stat().st_mtime, reverse=True), 1):
            size = file.stat().st_size / (1024 * 1024)
            msg += f"{i}. <code>{file.name}</code>\n   📊 {size:.2f} MB\n\n"
        
        keyboard = [[InlineKeyboardButton("🔙 Главное меню", callback_data='start')]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.effective_message.reply_text(msg, parse_mode='HTML', reply_markup=reply_markup)
        
    except Exception as e:
        await update.effective_message.reply_text(f"❌ Ошибка: {str(e)}")
        logger.error(f"Ошибка получения списка файлов: {e}", exc_info=True)

async def show_stats(update: Update, context):
    """Показать статистику"""
    stats = f"""
📊 <b>СТАТИСТИКА СИСТЕМЫ</b>

🤖 Режим: Webhook (Serverless)
📁 Временная папка: /tmp/uploads
🌐 Платформа: Vercel

<i>Для полного функционала используйте основного бота</i>
"""
    
    keyboard = [[InlineKeyboardButton("🔙 Назад", callback_data='start')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.effective_message.reply_text(
        stats,
        parse_mode='HTML',
        reply_markup=reply_markup
    )

async def button_handler(update: Update, context):
    """Обработка нажатий кнопок"""
    query = update.callback_query
    await query.answer()
    
    # Инструкция по загрузке
    if query.data == 'upload_info':
        msg = """📥 <b>Как загрузить файл:</b>

1️⃣ Просто отправьте файл (.xlsx, .xls, .csv) в этот чат
2️⃣ Файл будет сохранён временно
3️⃣ Используйте команды для обработки

💡 <b>Поддерживаемые форматы:</b>
• Excel (.xlsx, .xls)
• CSV (.csv)"""
        
        keyboard = [[InlineKeyboardButton("🔙 Назад", callback_data='start')]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.message.reply_text(msg, parse_mode='HTML', reply_markup=reply_markup)
        return
    
    # Помощь
    if query.data == 'help':
        msg = """
📖 <b>СПРАВКА</b>

<b>Команды:</b>
/start - Главное меню
/help - Справка

<b>Как использовать:</b>
1. Отправьте файл из системы 112
2. Файл будет загружен в систему
3. Просматривайте список файлов

<b>⚠️ Важно:</b>
На Vercel файлы хранятся временно.
Для постоянной работы используйте Railway или Render.
"""
        
        keyboard = [[InlineKeyboardButton("🔙 Назад", callback_data='start')]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.message.reply_text(msg, parse_mode='HTML', reply_markup=reply_markup)
        return
    
    handlers = {
        'start': start,
        'stats': show_stats,
        'list_files': list_files,
    }
    
    handler = handlers.get(query.data)
    if handler:
        await handler(update, context)
    else:
        await query.message.reply_text(f"Неизвестная команда: {query.data}")

# =============================================================================
# ПРИЛОЖЕНИЕ
# =============================================================================

app = Application.builder().token(TOKEN).build()

# Команды
app.add_handler(CommandHandler('start', start))
app.add_handler(CommandHandler('help', start))

# Документы
app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

# Кнопки
app.add_handler(CallbackQueryHandler(button_handler))

# =============================================================================
# WEBHOOK HANDLER ДЛЯ VERCEL
# =============================================================================

async def webhook(request):
    """Обработчик webhook для Vercel"""
    try:
        update = Update.de_json(await request.json(), app.bot)
        await app.process_update(update)
        return {'statusCode': 200, 'body': 'ok'}
    except Exception as e:
        logger.error(f"Ошибка webhook: {e}", exc_info=True)
        return {'statusCode': 500, 'body': str(e)}

# Для Vercel API Routes
async def handler(request):
    """Vercel serverless function handler"""
    return await webhook(request)
