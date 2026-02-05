# 🐳 PostgreSQL через Docker

## 📚 Что включено

Docker Compose конфигурация запускает:
- **PostgreSQL 16** - база данных (порт 5432)
- **pgAdmin 4** - веб-интерфейс для управления БД (порт 5050)

---

## 🚀 Быстрый Старт

### Шаг 1: Установка Docker Desktop

Если Docker не установлен:

1. **Скачать:** https://www.docker.com/products/docker-desktop/
2. **Установить** Docker Desktop для Windows
3. **Перезагрузить** компьютер
4. **Запустить** Docker Desktop (значок кита в трее должен быть зеленым)

### Шаг 2: Запуск PostgreSQL

Двойной клик: `DOCKER_ЗАПУСК.bat`

Скрипт автоматически:
- ✅ Скачает образы PostgreSQL и pgAdmin
- ✅ Создаст контейнеры
- ✅ Запустит сервисы
- ✅ Проверит подключение

**Время первого запуска:** ~2-3 минуты (скачивание образов)

### Шаг 3: Настройка схемы БД

Двойной клик: `НАСТРОЙКА_POSTGRESQL.bat`

Создаст:
- Таблицы (operators, services, regions, fixations, incidents_112)
- Индексы для быстрого поиска
- Представления для аналитики
- Функции для автокатегоризации

### Шаг 4: Импорт данных

Двойной клик: `ИМПОРТ_В_POSTGRESQL.bat`

Импортирует ~1.7 млн записей из SQLite и CSV файлов.

---

## 🔐 Учетные данные

### PostgreSQL
```
Host:     localhost
Port:     5432
User:     qayta_user
Password: qayta_password_2026
Database: qayta_data
```

### pgAdmin (веб-интерфейс)
```
URL:      http://localhost:5050
Email:    admin@qayta.uz
Password: admin
```

---

## 📊 Доступ к данным

### pgAdmin (Веб-интерфейс)

1. Открыть: http://localhost:5050
2. Войти: `admin@qayta.uz` / `admin`
3. Add New Server:
   - Name: `QAyta PostgreSQL`
   - Connection:
     - Host: `postgres` (имя контейнера)
     - Port: `5432`
     - Database: `qayta_data`
     - Username: `qayta_user`
     - Password: `qayta_password_2026`

### Python

```python
import psycopg2

conn = psycopg2.connect(
    host='localhost',
    port=5432,
    database='qayta_data',
    user='qayta_user',
    password='qayta_password_2026'
)
```

### psql через Docker

```powershell
docker exec -it qayta-postgres psql -U qayta_user -d qayta_data
```

---

## 🛠️ Управление

### Запустить контейнеры
```batch
DOCKER_ЗАПУСК.bat
```
Или:
```powershell
docker compose up -d
```

### Остановить контейнеры
```batch
DOCKER_ОСТАНОВКА.bat
```
Или:
```powershell
docker compose down
```

### Статус и логи
```batch
DOCKER_СТАТУС.bat
```
Или:
```powershell
docker compose ps
docker compose logs postgres
docker compose logs pgadmin
```

### Перезапуск
```powershell
docker compose restart
```

### Полная очистка (удалить данные)
```powershell
docker compose down -v
```
⚠️ **Это удалит все данные из PostgreSQL!**

---

## 📦 Резервное копирование

### Создать бэкап
```powershell
docker exec qayta-postgres pg_dump -U qayta_user qayta_data > backups\qayta_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').sql
```

### Восстановить из бэкапа
```powershell
docker exec -i qayta-postgres psql -U qayta_user -d qayta_data < backups\qayta_backup_20260205_143000.sql
```

---

## 🔍 Полезные команды

### Подключиться к PostgreSQL
```powershell
docker exec -it qayta-postgres psql -U qayta_user -d qayta_data
```

### Выполнить SQL
```powershell
docker exec qayta-postgres psql -U qayta_user -d qayta_data -c "SELECT COUNT(*) FROM fixations;"
```

### Просмотреть логи
```powershell
docker compose logs -f postgres
```

### Проверить состояние
```powershell
docker compose ps
```

---

## 📂 Структура

```
QAyta skript/
├── docker-compose.yml          # Конфигурация Docker
├── DOCKER_ЗАПУСК.bat          # Запуск контейнеров
├── DOCKER_ОСТАНОВКА.bat       # Остановка контейнеров
├── DOCKER_СТАТУС.bat          # Статус и логи
├── backups/                   # Папка для бэкапов
└── config/
    └── postgresql.env         # Настройки подключения
```

---

## 🐳 docker-compose.yml

```yaml
services:
  postgres:
    image: postgres:16-alpine
    ports:
      - "5432:5432"
    environment:
      POSTGRES_USER: qayta_user
      POSTGRES_PASSWORD: qayta_password_2026
      POSTGRES_DB: qayta_data
    volumes:
      - postgres_data:/var/lib/postgresql/data
      - ./backups:/backups

  pgadmin:
    image: dpage/pgadmin4
    ports:
      - "5050:80"
    environment:
      PGADMIN_DEFAULT_EMAIL: admin@qayta.uz
      PGADMIN_DEFAULT_PASSWORD: admin
```

---

## 🚨 Решение проблем

### Docker не запускается

1. Проверить что Docker Desktop запущен (иконка кита в трее)
2. Перезапустить Docker Desktop
3. Проверить Hyper-V включен (Windows Features)

### Порт 5432 уже занят

Если на компьютере уже запущен PostgreSQL:

1. Остановить локальный PostgreSQL:
   ```powershell
   net stop postgresql-x64-16
   ```

2. Или изменить порт в docker-compose.yml:
   ```yaml
   ports:
     - "5433:5432"  # Изменено на 5433
   ```

### Контейнер не запускается

```powershell
# Просмотреть логи
docker compose logs postgres

# Пересоздать контейнеры
docker compose down
docker compose up -d --force-recreate
```

### pgAdmin не открывается

1. Подождать 30 секунд после запуска
2. Проверить логи: `docker compose logs pgadmin`
3. Очистить кеш браузера
4. Попробовать другой браузер

---

## ⚡ Преимущества Docker

- ✅ **Быстрая установка** - один файл, одна команда
- ✅ **Изоляция** - не влияет на систему
- ✅ **Портативность** - одинаково работает везде
- ✅ **Легкое удаление** - `docker compose down -v`
- ✅ **Множество версий** - разные PostgreSQL в разных проектах
- ✅ **Включен pgAdmin** - сразу готов к работе

---

## 📈 Производительность

Docker контейнер с PostgreSQL обычно работает с производительностью ~90-95% от нативной установки.

Для оптимизации больших данных (>10M записей):

```yaml
environment:
  - POSTGRES_SHARED_BUFFERS=256MB
  - POSTGRES_EFFECTIVE_CACHE_SIZE=1GB
  - POSTGRES_WORK_MEM=16MB
  - POSTGRES_MAINTENANCE_WORK_MEM=128MB
```

---

## ✅ Следующие шаги

После успешного запуска Docker:

1. ✅ Открыть http://localhost:5050 (pgAdmin)
2. ✅ Запустить `НАСТРОЙКА_POSTGRESQL.bat`
3. ✅ Запустить `ИМПОРТ_В_POSTGRESQL.bat`
4. 📊 Начать анализ данных!

---

**Дата создания:** 5 февраля 2026  
**Версия Docker:** Compose V2  
**PostgreSQL:** 16 Alpine
