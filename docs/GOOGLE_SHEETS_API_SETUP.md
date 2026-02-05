# Настройка Google Sheets API для импорта

## Что создано:

1. **[import_from_sheets_api.py](scripts/database/import_from_sheets_api.py)** - импорт напрямую через API
2. **[ИМПОРТ_ИЗ_SHEETS_API.bat](ИМПОРТ_ИЗ_SHEETS_API.bat)** - запуск одной кнопкой

## Как получить credentials.json:

### Шаг 1: Создать проект в Google Cloud

1. Откройте: https://console.cloud.google.com/
2. Создайте новый проект или выберите существующий
3. Название проекта: `qayta-import` (любое)

### Шаг 2: Включить Google Sheets API

1. В Google Cloud Console → API и сервисы → Библиотека
2. Найти: **Google Sheets API**
3. Нажать: **Включить**

### Шаг 3: Создать Service Account

1. API и сервисы → Учетные данные
2. Создать учетные данные → Аккаунт службы
3. Название: `qayta-sheets-reader`
4. Роль: **Проект → Просмотр**
5. Нажать: **Готово**

### Шаг 4: Создать ключ

1. Нажать на созданный Service Account
2. Вкладка: **Ключи**
3. Добавить ключ → Создать новый ключ
4. Тип: **JSON**
5. Скачается файл `qayta-sheets-reader-xxxxx.json`

### Шаг 5: Установить ключ

1. Переименовать скачанный файл в `credentials.json`
2. Скопировать в папку: `config/credentials.json`

### Шаг 6: Дать доступ к таблицам

1. Открыть `credentials.json`
2. Найти строку `"client_email": "qayta-sheets-reader@..."`
3. Скопировать этот email
4. **Для КАЖДОЙ Google Sheets таблицы:**
   - Открыть таблицу
   - Поделиться → Добавить этот email
   - Права: **Читатель**
   - Отправить

## Структура info.json

Файл `docs/info.json` должен содержать:

```json
{
  "Оператор 1": [
    {
      "name": "FIKSA июнь 2025",
      "url": "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit"
    }
  ],
  "Оператор 2": [
    {
      "name": "FIKSA июль 2025",
      "url": "https://docs.google.com/spreadsheets/d/ANOTHER_ID/edit"
    }
  ]
}
```

## После настройки:

**Запуск:**
```batch
ИМПОРТ_ИЗ_SHEETS_API.bat
```

Скрипт:
1. Подключится к Google Sheets API
2. Прочитает все таблицы из info.json
3. Загрузит данные напрямую
4. Импортирует в PostgreSQL
5. Покажет статистику

## Преимущества этого метода:

- ✅ Всегда свежие данные (прямо из Google Sheets)
- ✅ Нет промежуточных CSV файлов
- ✅ Автоматическое определение структуры
- ✅ Быстрее чем скачивание CSV
- ✅ Меньше ошибок кодировки
- ✅ Автоматический retry при ошибках

## Если нет credentials.json:

Альтернативный метод - использовать существующие скрипты:

```batch
1. update_from_sheets.py  (скачать CSV)
2. Потом импортировать из CSV
```

Но импорт через API - более надежный и современный способ!

---

**Дата:** 5 февраля 2026
