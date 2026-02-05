#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Тест подключения к Google Sheets API
"""
import sys
import io
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import os
import socket

# Прокси настройки
os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
socket.setdefaulttimeout(120)

from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / 'config'

CREDENTIALS_FILE = CONFIG_DIR / 'credentials.json'
TOKEN_FILE = CONFIG_DIR / 'token.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

print('='*80)
print('ТЕСТ ПОДКЛЮЧЕНИЯ К GOOGLE SHEETS API')
print('='*80)
print(f'Credentials: {CREDENTIALS_FILE.exists()}')
print(f'Token: {TOKEN_FILE.exists()}')
print(f'Spreadsheet ID: {MASTER_SPREADSHEET_ID}')
print()

try:
    creds = None
    if TOKEN_FILE.exists():
        print('[1] Загрузка существующего токена...')
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        print(f'    Валидный: {creds.valid}')
        print(f'    Просрочен: {creds.expired if hasattr(creds, "expired") else "N/A"}')
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print('[2] Обновление токена...')
            creds.refresh(Request())
            print('    ✅ Токен обновлен')
        else:
            print('[2] Авторизация через браузер...')
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
            print('    ✅ Авторизация выполнена')
        
        with open(TOKEN_FILE, 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
        print('    ✅ Токен сохранен')
    else:
        print('[1] Токен валидный')
    
    print()
    print('[3] Создание сервиса Google Sheets API...')
    service = build('sheets', 'v4', credentials=creds)
    print('    ✅ Сервис создан')
    
    print()
    print('[4] Запрос метаданных таблицы...')
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=MASTER_SPREADSHEET_ID).execute()
    
    title = spreadsheet.get('properties', {}).get('title', 'Неизвестно')
    sheets = spreadsheet.get('sheets', [])
    
    print(f'    ✅ Таблица: {title}')
    print(f'    ✅ Листов: {len(sheets)}')
    
    print()
    print('ЛИСТЫ В ТАБЛИЦЕ:')
    for i, sheet in enumerate(sheets, 1):
        sheet_title = sheet['properties']['title']
        row_count = sheet['properties']['gridProperties'].get('rowCount', 'N/A')
        print(f'  {i}. {sheet_title} ({row_count} строк)')
    
    print()
    print('[5] Тест чтения данных с первого листа...')
    first_sheet = sheets[0]['properties']['title']
    result = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SPREADSHEET_ID,
        range=f"'{first_sheet}'!A1:Z10"
    ).execute()
    
    values = result.get('values', [])
    print(f'    ✅ Прочитано строк: {len(values)}')
    if values:
        print(f'    Первая строка: {values[0][:5]}...')
    
    print()
    print('='*80)
    print('✅ ВСЕ ТЕСТЫ ПРОЙДЕНЫ УСПЕШНО!')
    print('='*80)
    
except HttpError as e:
    print()
    print('='*80)
    print(f'❌ HTTP ОШИБКА {e.resp.status}')
    print('='*80)
    print(f'Причина: {e.reason}')
    print(f'Детали: {e.error_details if hasattr(e, "error_details") else "N/A"}')
    print()
    if e.resp.status == 400:
        print('ОШИБКА 400: Неправильный запрос')
        print('Возможные причины:')
        print('  - Неверный формат Spreadsheet ID')
        print('  - Неправильные параметры запроса')
        print('  - Проблема с токеном авторизации')
    elif e.resp.status == 403:
        print('ОШИБКА 403: Доступ запрещен')
        print('Возможные причины:')
        print('  - Таблица не открыта для вашего Google аккаунта')
        print('  - API не включен в Google Cloud Console')
    elif e.resp.status == 404:
        print('ОШИБКА 404: Таблица не найдена')
        print('Проверьте Spreadsheet ID:', MASTER_SPREADSHEET_ID)
    
    import traceback
    traceback.print_exc()

except Exception as e:
    print()
    print('='*80)
    print('❌ ОШИБКА')
    print('='*80)
    print(f'{type(e).__name__}: {e}')
    import traceback
    traceback.print_exc()
