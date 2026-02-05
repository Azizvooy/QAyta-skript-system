# 🐘 Инструкция Установки PostgreSQL

## ⚠️ PostgreSQL НЕ ОБНАРУЖЕН

Для работы с PostgreSQL базой данных необходимо установить PostgreSQL сервер.

---

## 📥 СПОСОБ 1: Автоматическая установка через Chocolatey (10 минут)

### Шаг 1: Установка Chocolatey (если нет)

Откройте **PowerShell от имени Администратора** и выполните:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
```

### Шаг 2: Установка PostgreSQL

```powershell
choco install postgresql16 -y
```

### Шаг 3: Настройка пароля

После установки откройте **psql** и выполните:

```sql
ALTER USER postgres WITH PASSWORD 'postgres';
```

### Шаг 4: Запуск настройки

Двойной клик:
```
НАСТРОЙКА_POSTGRESQL.bat
```

---

## 📥 СПОСОБ 2: Ручная установка через установщик (15 минут)

### Шаг 1: Скачать установщик

1. Перейти: **https://www.postgresql.org/download/windows/**
2. Выбрать: **Download the installer** от EnterpriseDB
3. Скачать версию: **PostgreSQL 16.x** для Windows x86-64

### Шаг 2: Установка

1. Запустить скачанный файл (postgresql-16.x-windows-x64.exe)
2. **Installation Directory:** `C:\Program Files\PostgreSQL\16` (по умолчанию)
3. **Components:** Выбрать все:
   - ✅ PostgreSQL Server
   - ✅ pgAdmin 4
   - ✅ Stack Builder
   - ✅ Command Line Tools

4. **Data Directory:** `C:\Program Files\PostgreSQL\16\data` (по умолчанию)

5. **Password:** Придумать и запомнить пароль для postgres
   - Например: `postgres`
   - ⚠️ **ЗАПИШИТЕ ПАРОЛЬ!** Он понадобится

6. **Port:** `5432` (по умолчанию)

7. **Locale:** Russian, Russia или [Default locale]

8. Нажать **Next** и дождаться установки

### Шаг 3: Проверка

Откройте **PowerShell** и выполните:

```powershell
psql --version
```

Должно показать: `psql (PostgreSQL) 16.x`

### Шаг 4: Настройка проекта

1. Отредактировать файл: `config/postgresql.env`
   
   ```properties
   DB_HOST=localhost
   DB_PORT=5432
   DB_NAME=qayta_data
   DB_USER=postgres
   DB_PASSWORD=ваш_пароль_который_вы_установили
   ```

2. Запустить: **Двойной клик на `НАСТРОЙКА_POSTGRESQL.bat`**

---

## 📥 СПОСОБ 3: Docker (для опытных пользователей)

### Если у вас установлен Docker:

```powershell
# Запустить PostgreSQL в контейнере
docker run --name qayta-postgres -e POSTGRES_PASSWORD=qayta_password_2026 -e POSTGRES_DB=qayta_data -e POSTGRES_USER=qayta_user -p 5432:5432 -d postgres:16

# Проверить что работает
docker ps
```

Затем запустить:
```
НАСТРОЙКА_POSTGRESQL.bat
```

---

## 🔍 Проверка работы PostgreSQL

### Windows Services

1. Нажать `Win + R`
2. Ввести: `services.msc`
3. Найти: **postgresql-x64-16** или **PostgreSQL**
4. Статус должен быть: **Запущена (Running)**

### Командная строка

```powershell
# Проверка версии
psql --version

# Подключение к БД
psql -U postgres

# В psql консоли
\l      # Показать все базы данных
\q      # Выход
```

---

## 📊 После установки PostgreSQL

### 1. Запустить настройку

Двойной клик: `НАСТРОЙКА_POSTGRESQL.bat`

Скрипт автоматически:
- ✅ Установит Python зависимости
- ✅ Создаст пользователя `qayta_user`
- ✅ Создаст базу данных `qayta_data`
- ✅ Создаст таблицы и индексы
- ✅ Создаст аналитические представления

### 2. Импортировать данные

Двойной клик: `ИМПОРТ_В_POSTGRESQL.bat`

Импортирует:
- 📊 Все данные из SQLite (1.7 млн записей)
- 📁 Все CSV файлы из exported_sheets/
- 👥 Операторов
- 🏢 Службы (101, 102, 103, 104)
- 🌍 Регионы

### 3. Использование

**pgAdmin 4:**
1. Открыть pgAdmin 4 (установлен вместе с PostgreSQL)
2. Servers → PostgreSQL 16 → Databases → qayta_data

**DBeaver:**
1. New Connection → PostgreSQL
2. Host: `localhost`, Port: `5432`
3. Database: `qayta_data`, User: `qayta_user`

**Python:**
```python
import psycopg2
conn = psycopg2.connect(
    host='localhost',
    database='qayta_data',
    user='qayta_user',
    password='qayta_password_2026'
)
```

---

## 🚨 Решение проблем

### PostgreSQL не запускается

```powershell
# Остановить службу
net stop postgresql-x64-16

# Запустить заново
net start postgresql-x64-16
```

### Забыли пароль postgres

1. Открыть файл: `C:\Program Files\PostgreSQL\16\data\pg_hba.conf`
2. Найти строку: `host all all 127.0.0.1/32 scram-sha-256`
3. Заменить на: `host all all 127.0.0.1/32 trust`
4. Перезапустить PostgreSQL
5. Подключиться без пароля:
   ```powershell
   psql -U postgres
   ALTER USER postgres WITH PASSWORD 'новый_пароль';
   ```
6. Вернуть `scram-sha-256` в pg_hba.conf
7. Перезапустить PostgreSQL

### Порт 5432 занят

```powershell
# Проверить что использует порт
netstat -ano | findstr :5432

# Найти процесс по PID
tasklist | findstr <PID>
```

---

## 📚 Полезные ссылки

- **Документация PostgreSQL:** https://www.postgresql.org/docs/
- **pgAdmin документация:** https://www.pgadmin.org/docs/
- **Скачать PostgreSQL:** https://www.postgresql.org/download/windows/
- **Chocolatey:** https://chocolatey.org/install

---

## ✅ Что дальше?

После успешной установки PostgreSQL:

1. ✅ Запустить `НАСТРОЙКА_POSTGRESQL.bat`
2. ✅ Запустить `ИМПОРТ_В_POSTGRESQL.bat`
3. 📊 Открыть pgAdmin и исследовать данные
4. 📈 Использовать SQL для аналитики
5. 🔍 Создавать отчеты с помощью представлений

---

**Дата:** 5 февраля 2026  
**Версия:** 1.0
