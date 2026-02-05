#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Просмотр структуры мастер-таблицы и списка операторов
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
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / 'config'

TOKEN_FILE = CONFIG_DIR / 'token.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
MASTER_SPREADSHEET_ID = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"

print('Чтение листа "Настройки"...\n')

creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
service = build('sheets', 'v4', credentials=creds)

# Читаем лист Настройки
result = service.spreadsheets().values().get(
    spreadsheetId=MASTER_SPREADSHEET_ID,
    range="Настройки!A1:Z100"
).execute()

values = result.get('values', [])

if values:
    headers = values[0]
    print('ЗАГОЛОВКИ:')
    for i, h in enumerate(headers):
        print(f'  {i}: {h}')
    
    print(f'\nВСЕГО СТРОК: {len(values) - 1}')
    print('\nПЕРВЫЕ 10 ОПЕРАТОРОВ:')
    print('-' * 100)
    
    for i, row in enumerate(values[1:11], 1):
        print(f'{i}. ФИО: {row[0] if len(row) > 0 else "N/A"}')
        print(f'   ID таблицы: {row[1] if len(row) > 1 else "N/A"}')
        print(f'   Статус: {row[2] if len(row) > 2 else "N/A"}')
        print()
