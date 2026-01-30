"""Быстрое пересохранение - если данные уже в памяти"""
import pickle
import pandas as pd
import os

# Загрузим если есть кеш
if os.path.exists('collection_cache.pkl'):
    print("📦 Загружаем кеш...")
    with open('collection_cache.pkl', 'rb') as f:
        cache = pickle.load(f)
    print(f"✅ Найдено в кеше: {len(cache.get('stats', []))} операторов")
else:
    print("❌ Кеш не найден. Запустите collect_to_excel.py")
