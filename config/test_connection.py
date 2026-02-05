"""Тест подключения к Google Sheets API"""
import socket
import time
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Увеличиваем таймаут
socket.setdefaulttimeout(120)

TOKEN_FILE = 'token.json'
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

print("🔍 Тест подключения к Google Sheets API\n")

try:
    print("1️⃣ Загрузка токена...")
    if not os.path.exists(TOKEN_FILE):
        print("❌ Файл token.json не найден! Сначала нужно авторизоваться.")
        exit(1)
    
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    print("✅ Токен загружен")
    
    print("\n2️⃣ Создание сервиса Google Sheets...")
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    print("✅ Сервис создан")
    
    print("\n3️⃣ Пробный запрос к таблице...")
    print(f"   ID таблицы: {MASTER_SPREADSHEET_ID}")
    
    start = time.time()
    result = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SPREADSHEET_ID,
        range="Настройки!A1:A1"
    ).execute()
    elapsed = time.time() - start
    
    print(f"✅ Запрос выполнен за {elapsed:.2f} сек")
    print(f"   Получено данных: {result.get('values', [])}")
    print("\n🎉 ПОДКЛЮЧЕНИЕ РАБОТАЕТ!")
    
except FileNotFoundError:
    print("❌ Файл token.json не найден")
except Exception as e:
    print(f"\n❌ ОШИБКА: {type(e).__name__}")
    print(f"   Сообщение: {str(e)}")
    
    if "Timeout" in str(e) or "10060" in str(e):
        print("\n💡 РЕКОМЕНДАЦИИ:")
        print("   - Проверьте интернет-соединение")
        print("   - Временно отключите антивирус/firewall")
        print("   - Попробуйте другую сеть (мобильный интернет)")
        print("   - Проверьте, не блокирует ли корпоративная сеть googleapis.com")
