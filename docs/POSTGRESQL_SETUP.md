# PostgreSQL Setup Guide / Руководство по Настройке PostgreSQL

## 📚 Введение

PostgreSQL - мощная реляционная база данных для эффективного хранения и анализа больших объемов данных.

**Преимущества перед SQLite:**
- ✅ Лучшая производительность на больших объемах (миллионы записей)
- ✅ Параллельный доступ многих пользователей
- ✅ Продвинутые аналитические функции
- ✅ Полнотекстовый поиск
- ✅ JSON поддержка для гибких структур
- ✅ Репликация и резервное копирование

---

## 🚀 Быстрый Старт

### Вариант 1: Автоматическая Установка (Рекомендуется)

```batch
1. Двойной клик: НАСТРОЙКА_POSTGRESQL.bat
2. Дождитесь завершения
3. Двойной клик: ИМПОРТ_В_POSTGRESQL.bat
```

### Вариант 2: Ручная Установка

#### Шаг 1: Установка PostgreSQL

**Windows:**
1. Скачать: https://www.postgresql.org/download/windows/
2. Запустить установщик
3. Пароль для postgres: `postgres` (запомните!)
4. Порт: `5432` (по умолчанию)

**С помощью Chocolatey:**
```powershell
choco install postgresql
```

#### Шаг 2: Создание Пользователя и БД

Откройте **psql** или **pgAdmin** и выполните:

```sql
-- Создать пользователя
CREATE USER qayta_user WITH PASSWORD 'qayta_password_2026';

-- Создать базу данных
CREATE DATABASE qayta_data OWNER qayta_user;

-- Дать права
GRANT ALL PRIVILEGES ON DATABASE qayta_data TO qayta_user;
```

#### Шаг 3: Установка Python Зависимостей

```bash
pip install -r requirements_postgresql.txt
```

#### Шаг 4: Создание Схемы

```bash
python scripts/database/create_postgresql_schema.py
```

#### Шаг 5: Импорт Данных

```bash
python scripts/database/import_to_postgresql.py
```

---

## ⚙️ Конфигурация

Файл `config/postgresql.env`:

```properties
DB_HOST=localhost
DB_PORT=5432
DB_NAME=qayta_data
DB_USER=qayta_user
DB_PASSWORD=qayta_password_2026
```

**Для удаленного сервера:**
```properties
DB_HOST=192.168.1.100
DB_PORT=5432
```

---

## 📊 Структура Базы Данных

### Таблицы

#### 1. operators (Операторы)
```sql
operator_id    SERIAL PRIMARY KEY
operator_name  VARCHAR(255) UNIQUE
phone          VARCHAR(50)
position       VARCHAR(100)
active         BOOLEAN
```

#### 2. services (Службы)
```sql
service_id     SERIAL PRIMARY KEY
service_code   VARCHAR(10) UNIQUE  -- 101, 102, 103, 104
service_name   VARCHAR(255)
active         BOOLEAN
```

Предустановленные службы:
- 101 - Пожарная служба
- 102 - Скорая медицинская помощь
- 103 - Полиция
- 104 - Аварийная газовая служба

#### 3. regions (Регионы)
```sql
region_id      SERIAL PRIMARY KEY
region_name    VARCHAR(255) UNIQUE
region_code    VARCHAR(50)
```

#### 4. fixations (Фиксации)
Основная таблица с данными:
```sql
fixation_id       BIGSERIAL PRIMARY KEY
card_number       VARCHAR(50)
operator_id       INTEGER -> operators
service_id        INTEGER -> services
region_id         INTEGER -> regions
call_date         TIMESTAMP
status            VARCHAR(255)
status_category   VARCHAR(50)  -- Автоматически: Положительно/Отрицательно/Прочее
phone             VARCHAR(50)
address           TEXT
complaint         TEXT
```

#### 5. incidents_112 (Инциденты 112)
```sql
incident_id       BIGSERIAL PRIMARY KEY
incident_number   VARCHAR(100) UNIQUE
card_number       VARCHAR(50)
service_id        INTEGER
call_time         TIMESTAMP
operator_112      VARCHAR(255)
```

---

## 📈 Аналитические Представления

### v_fixations_full
Полная информация о фиксациях с именами операторов, служб, регионов:
```sql
SELECT * FROM v_fixations_full 
WHERE service_code = '102' 
  AND status_category = 'Отрицательно';
```

### v_operator_statistics
Статистика по операторам:
```sql
SELECT 
    operator_name,
    total_fixations,
    positive_count,
    positive_percentage
FROM v_operator_statistics
ORDER BY total_fixations DESC
LIMIT 10;
```

### v_service_statistics
Статистика по службам:
```sql
SELECT * FROM v_service_statistics
ORDER BY total_fixations DESC;
```

### v_region_statistics
Статистика по регионам:
```sql
SELECT * FROM v_region_statistics
WHERE total_fixations > 1000;
```

---

## 🔍 Полезные SQL Запросы

### 1. ТОП операторов по положительным фиксациям
```sql
SELECT 
    operator_name,
    positive_count,
    positive_percentage
FROM v_operator_statistics
WHERE total_fixations > 100
ORDER BY positive_percentage DESC
LIMIT 20;
```

### 2. Жалобы по службам за период
```sql
SELECT 
    s.service_name,
    COUNT(*) as complaints_count
FROM fixations f
JOIN services s ON f.service_id = s.service_id
WHERE f.status_category = 'Отрицательно'
  AND f.call_date >= '2026-01-01'
GROUP BY s.service_name
ORDER BY complaints_count DESC;
```

### 3. Динамика по дням
```sql
SELECT 
    DATE(call_date) as date,
    COUNT(*) as total,
    COUNT(CASE WHEN status_category = 'Положительно' THEN 1 END) as positive,
    COUNT(CASE WHEN status_category = 'Отрицательно' THEN 1 END) as negative
FROM fixations
WHERE call_date >= CURRENT_DATE - INTERVAL '30 days'
GROUP BY DATE(call_date)
ORDER BY date DESC;
```

### 4. Матрица регион-служба
```sql
SELECT 
    r.region_name,
    s.service_code,
    COUNT(*) as count
FROM fixations f
JOIN regions r ON f.region_id = r.region_id
JOIN services s ON f.service_id = s.service_id
GROUP BY r.region_name, s.service_code
ORDER BY r.region_name, s.service_code;
```

---

## 🛠️ Подключение к БД

### psql (Command Line)
```bash
psql -h localhost -U qayta_user -d qayta_data
```

### Python (psycopg2)
```python
import psycopg2

conn = psycopg2.connect(
    host='localhost',
    port=5432,
    database='qayta_data',
    user='qayta_user',
    password='qayta_password_2026'
)

cursor = conn.cursor()
cursor.execute("SELECT * FROM v_operator_statistics LIMIT 10")
results = cursor.fetchall()
```

### Python (pandas)
```python
import pandas as pd
from sqlalchemy import create_engine

engine = create_engine('postgresql://qayta_user:qayta_password_2026@localhost:5432/qayta_data')

df = pd.read_sql("SELECT * FROM v_fixations_full", engine)
```

### pgAdmin
1. Открыть pgAdmin
2. Add New Server
3. Host: `localhost`, Port: `5432`
4. Database: `qayta_data`
5. Username: `qayta_user`

### DBeaver
1. New Database Connection
2. PostgreSQL
3. Host: `localhost:5432`
4. Database: `qayta_data`
5. Username/Password из конфига

---

## 🔐 Безопасность

### Изменение Пароля

```sql
ALTER USER qayta_user WITH PASSWORD 'новый_пароль';
```

Обновить в `config/postgresql.env`:
```
DB_PASSWORD=новый_пароль
```

### Ограничение Доступа

В `pg_hba.conf`:
```
# Только локальный доступ
host    qayta_data    qayta_user    127.0.0.1/32    md5

# Доступ из сети
host    qayta_data    qayta_user    192.168.1.0/24  md5
```

---

## 📦 Резервное Копирование

### Создание Бэкапа
```bash
pg_dump -U qayta_user -d qayta_data -F c -f backup_qayta_$(date +%Y%m%d).dump
```

### Восстановление
```bash
pg_restore -U qayta_user -d qayta_data -c backup_qayta_20260205.dump
```

### Автоматический Бэкап (Windows)
Создать задачу в планировщике:
```batch
@echo off
set BACKUP_DIR=C:\backups\postgresql
set DATE=%date:~-4,4%%date:~-7,2%%date:~-10,2%
pg_dump -U qayta_user -d qayta_data -F c -f %BACKUP_DIR%\qayta_%DATE%.dump
```

---

## 🚨 Решение Проблем

### Ошибка подключения
```
could not connect to server
```
**Решение:**
1. Проверить запущен ли PostgreSQL: `services.msc` → PostgreSQL
2. Проверить порт: `netstat -an | findstr 5432`
3. Проверить пароль в `postgresql.env`

### Ошибка прав доступа
```
permission denied for table
```
**Решение:**
```sql
GRANT ALL PRIVILEGES ON ALL TABLES IN SCHEMA public TO qayta_user;
GRANT ALL PRIVILEGES ON ALL SEQUENCES IN SCHEMA public TO qayta_user;
```

### Медленные запросы
```sql
-- Создать индексы
CREATE INDEX idx_fixations_date_service ON fixations(call_date, service_id);
CREATE INDEX idx_fixations_region_status ON fixations(region_id, status_category);

-- Анализ производительности
EXPLAIN ANALYZE SELECT ...;
```

---

## 📊 Сравнение с SQLite

| Характеристика | SQLite | PostgreSQL |
|---------------|--------|------------|
| Размер БД | До 140 TB | Неограничен |
| Пользователи | 1 | Множество |
| Транзакции | Да | Да (ACID) |
| Репликация | Нет | Да |
| JSON | Базовый | Расширенный |
| Полнотекстовый поиск | Нет | Да |
| Аналитика | Ограничена | Продвинутая |

---

## 🎯 Следующие Шаги

1. ✅ Запустить `НАСТРОЙКА_POSTGRESQL.bat`
2. ✅ Запустить `ИМПОРТ_В_POSTGRESQL.bat`
3. 📊 Подключиться через pgAdmin/DBeaver
4. 🔍 Попробовать SQL запросы из раздела "Полезные запросы"
5. 📈 Создать дашборды для аналитики

---

**Документация:** https://www.postgresql.org/docs/  
**Дата создания:** 5 февраля 2026  
**Автор:** GitHub Copilot
