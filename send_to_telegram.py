#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Отправка отчета в Telegram
"""

import requests
from pathlib import Path
import sys

# Конфигурация
BOT_TOKEN = "8141079204:AAErNrjLqTu4Vj1_7VS2kjGFKcR3lU9L9N4"
CHAT_ID = "2012682567"

# Путь к отчету
BASE_DIR = Path(__file__).parent
REPORTS_DIR = BASE_DIR / 'output' / 'reports'

print('\n' + '='*80)
print('OTPRAVKA OTCHETA V TELEGRAM')
print('='*80)

try:
    # Находим последний отчет
    reports = sorted(REPORTS_DIR.glob('📊_ПОЛНЫЙ_ОТЧЕТ_*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True)
    
    if not reports:
        print('\nOshibka: Otchety ne naideny!')
        sys.exit(1)
    
    report_file = reports[0]
    file_size_mb = report_file.stat().st_size / (1024 * 1024)
    
    print(f'\nFail: {report_file.name}')
    print(f'Razmer: {file_size_mb:.1f} MB')
    
    # Проверяем размер файла (Telegram ограничение 50 МБ)
    if file_size_mb > 50:
        print(f'\nOshibka: Fail slishkom bolshoi ({file_size_mb:.1f} MB > 50 MB)')
        sys.exit(1)
    
    # Отправляем сообщение
    print(f'\n[1/2] Otpravka soobshcheniya...')
    message = f"""
📊 <b>ПОЛНЫЙ ОТЧЕТ ОБНОВЛЕН</b>

✅ Данные обновлены из Google Sheets
📅 Дата: {report_file.stat().st_mtime.strftime('%d.%m.%Y %H:%M')}
📈 Размер: {file_size_mb:.1f} МБ

Отправляю файл...
"""
    
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage'
    data = {
        'chat_id': CHAT_ID,
        'text': message,
        'parse_mode': 'HTML'
    }
    
    response = requests.post(url, data=data, timeout=30)
    
    if response.status_code == 200:
        print('  OK: Soobshchenie otpravleno')
    else:
        print(f'  Oshibka: {response.text}')
    
    # Отправляем файл
    print(f'\n[2/2] Otpravka faila ({file_size_mb:.1f} MB)...')
    print('  (Eto mozhet zanyat neskolko minut)')
    
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendDocument'
    
    with open(report_file, 'rb') as f:
        files = {'document': (report_file.name, f)}
        data = {
            'chat_id': CHAT_ID,
            'caption': f'📊 {report_file.name}'
        }
        
        response = requests.post(url, data=data, files=files, timeout=300)
    
    if response.status_code == 200:
        print('\n' + '='*80)
        print('USPESHNO OTPRAVLENO V TELEGRAM!')
        print('='*80)
        print(f'\nProveryte svoi Telegram')
        print(f'Otpravlen: {report_file.name}')
        print(f'Razmer: {file_size_mb:.1f} MB')
        print('\n' + '='*80)
    else:
        print(f'\nOshibka pri otpravke faila:')
        print(f'   Status: {response.status_code}')
        print(f'   Otvet: {response.text}')
        sys.exit(1)

except requests.exceptions.Timeout:
    print('\nOshibka: Prevysheno vremya ozhidaniya.')
    sys.exit(1)
except requests.exceptions.RequestException as e:
    print(f'\nOshibka seti: {e}')
    sys.exit(1)
except Exception as e:
    print(f'\nOshibka: {e}')
    import traceback
    traceback.print_exc()
    sys.exit(1)
