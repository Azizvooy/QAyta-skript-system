"""
Приложение для обработки данных по заявкам и инцидентам
Запрашивает файлы, обрабатывает, сопоставляет и создаёт отчёты
"""
import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import socket
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


class DataProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка данных 112 и Google Sheets")
        self.root.geometry("900x700")
        
        # Переменные для хранения данных
        self.selected_agents = []
        self.available_sheets = []  # Реальные названия листов из Google Sheets
        self.incident_files = []
        self.output_folder = None
        self.sheets_data = None
        
        # Google Sheets настройки
        self.base_dir = Path(__file__).parent
        self.token_file = self.base_dir / 'config' / 'token.json'
        self.credentials_file = self.base_dir / 'config' / 'credentials.json'
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.master_spreadsheet_id = "1s0nbLCo6q_KoM0jCP2v2vMxLbIHuScjigNTMSvUn0GA"
        
        # Прокси
        os.environ['HTTP_PROXY'] = 'http://10.145.62.76:3128'
        os.environ['HTTPS_PROXY'] = 'http://10.145.62.76:3128'
        socket.setdefaulttimeout(120)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Создание интерфейса"""
        # Заголовок
        title = tk.Label(self.root, text="Обработка данных о заявках и инцидентах",
                        font=("Arial", 16, "bold"), pady=20)
        title.pack()
        
        # Рамка для выбора агентов Google Sheets
        sheets_frame = ttk.LabelFrame(self.root, text="Шаг 1: Выбор операторов из Google Sheets", padding=10)
        sheets_frame.pack(fill="x", padx=20, pady=5)
        
        # Фрейм для кнопок
        btn_frame = ttk.Frame(sheets_frame)
        btn_frame.pack(fill="x", pady=5)
        
        ttk.Button(btn_frame, text="Загрузить список из Google Sheets",
                  command=self.load_sheets_list).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Выбрать операторов",
                  command=self.select_agents).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Выбрать все",
                  command=self.select_all_agents).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Очистить",
                  command=self.clear_agents).pack(side="left", padx=5)
        
        self.sheets_label = tk.Label(sheets_frame, text="Сначала загрузите список листов из Google Sheets", fg="gray")
        self.sheets_label.pack(pady=5)
        
        # Рамка для файлов 112
        incident_frame = ttk.LabelFrame(self.root, text="Шаг 2: Файлы с данными 112 (Excel)",
                                       padding=10)
        incident_frame.pack(fill="x", padx=20, pady=5)
        
        self.incident_label = tk.Label(incident_frame, text="Файлы не выбраны", fg="gray")
        self.incident_label.pack(side="left", padx=5)
        
        # Кнопки для выбора файлов
        btn_container = ttk.Frame(incident_frame)
        btn_container.pack(side="right", padx=5)
        
        ttk.Button(btn_container, text="Выбрать файлы",
                  command=self.select_incident_files).pack(side="left", padx=2)
        ttk.Button(btn_container, text="Выбрать папку",
                  command=self.select_incident_folder).pack(side="left", padx=2)
        ttk.Button(btn_container, text="Создать папку incoming",
                  command=self.create_incoming_folder).pack(side="left", padx=2)
        
        # Информация о выбранных файлах
        info_frame = ttk.LabelFrame(self.root, text="Выбранные файлы", padding=10)
        info_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        self.info_text = tk.Text(info_frame, height=10, wrap="word")
        self.info_text.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(info_frame, command=self.info_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.info_text.config(yscrollcommand=scrollbar.set)
        
        # Прогресс-бар
        self.progress_frame = ttk.Frame(self.root)
        self.progress_frame.pack(fill="x", padx=20, pady=10)
        
        self.progress_label = tk.Label(self.progress_frame, text="")
        self.progress_label.pack()
        
        self.progress = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        
        # Кнопки действий
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.process_btn = ttk.Button(button_frame, text="Начать обработку",
                                     command=self.start_processing, state="disabled")
        self.process_btn.pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="Выход", command=self.root.quit).pack(side="left", padx=5)
    
    def load_sheets_list(self):
        """Загрузить список документов и листов из Google Sheets"""
        try:
            self.sheets_label.config(text="Подключение к Google Sheets...", fg="blue")
            self.root.update()
            
            # Аутентификация
            service = self.authenticate_google()
            
            # Получаем список ID документов из листа "Настройки"
            self.update_progress("Загрузка списка документов из листа 'Настройки'...")
            
            try:
                settings_range = "Настройки!A2:B100"  # Колонка A - название, B - ID документа
                result = service.spreadsheets().values().get(
                    spreadsheetId=self.master_spreadsheet_id,
                    range=settings_range
                ).execute()
                
                values = result.get('values', [])
                
                if not values:
                    messagebox.showerror("Ошибка", "Лист 'Настройки' пуст или не найден")
                    self.sheets_label.config(text="Ошибка загрузки", fg="red")
                    return
                
                # Собираем ID документов
                spreadsheet_ids = []
                for row in values:
                    if len(row) >= 2 and row[1]:  # Есть ID во второй колонке
                        doc_id = str(row[1]).strip()
                        doc_name = str(row[0]).strip() if row[0] else f"Документ {len(spreadsheet_ids)+1}"
                        if doc_id:
                            spreadsheet_ids.append({'id': doc_id, 'name': doc_name})
                
                if not spreadsheet_ids:
                    messagebox.showerror("Ошибка", "Не найдено ID документов в листе 'Настройки'")
                    self.sheets_label.config(text="ID не найдены", fg="red")
                    return
                
                self.update_progress(f"Найдено документов: {len(spreadsheet_ids)}. Загрузка списка операторов...")
                
                # Теперь получаем листы из каждого документа
                self.available_sheets = []
                
                for idx, doc_info in enumerate(spreadsheet_ids, 1):
                    try:
                        self.update_progress(f"Загрузка листов {idx}/{len(spreadsheet_ids)}: {doc_info['name']}...")
                        
                        spreadsheet = service.spreadsheets().get(spreadsheetId=doc_info['id']).execute()
                        sheets = spreadsheet.get('sheets', [])
                        
                        for sheet in sheets:
                            title = sheet['properties']['title']
                            # Пропускаем служебные листы
                            if title not in ['Настройки', 'Статистика', 'Сводка', 'Тренды', 'Итого']:
                                # Добавляем информацию о листе и документе
                                self.available_sheets.append({
                                    'sheet_name': title,
                                    'doc_id': doc_info['id'],
                                    'doc_name': doc_info['name'],
                                    'display_name': f"{doc_info['name']} → {title}"
                                })
                    except Exception as e:
                        print(f"Ошибка при чтении документа {doc_info['name']}: {e}")
                        continue
                
                if not self.available_sheets:
                    messagebox.showwarning("Внимание", "Не найдено листов операторов в документах")
                    self.sheets_label.config(text="Листы не найдены", fg="red")
                    return
                
                self.sheets_label.config(
                    text=f"Загружено листов: {len(self.available_sheets)} из {len(spreadsheet_ids)} документов",
                    fg="green"
                )
                messagebox.showinfo("Успех", 
                    f"Загружено:\n• Документов: {len(spreadsheet_ids)}\n• Листов операторов: {len(self.available_sheets)}\n\nТеперь выберите операторов для импорта.")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось прочитать лист 'Настройки':\n\n{str(e)}")
                self.sheets_label.config(text="Ошибка чтения настроек", fg="red")
                return
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к Google Sheets:\n\n{str(e)}")
            self.sheets_label.config(text="Ошибка подключения", fg="red")
    
    def select_agents(self):
        """Диалог выбора агентов"""
        if not self.available_sheets:
            messagebox.showwarning("Внимание", "Сначала загрузите список листов из Google Sheets")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Выбор операторов")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text=f"Выберите операторов для импорта (найдено: {len(self.available_sheets)}):",
                font=("Arial", 10, "bold")).pack(pady=10)
        
        # Список агентов
        agents_frame = ttk.Frame(dialog)
        agents_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(agents_frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = tk.Listbox(agents_frame, selectmode="multiple", yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Добавляем названия в формате "Документ → Лист"
        for sheet_info in self.available_sheets:
            listbox.insert(tk.END, sheet_info['display_name'])
        
        # Восстанавливаем предыдущий выбор
        for idx, sheet_info in enumerate(self.available_sheets):
            if sheet_info in self.selected_agents:
                listbox.selection_set(idx)
        
        def confirm():
            selected_indices = listbox.curselection()
            self.selected_agents = [self.available_sheets[i] for i in selected_indices]
            self.sheets_label.config(
                text=f"Выбрано операторов: {len(self.selected_agents)}",
                fg="green" if self.selected_agents else "gray"
            )
            self.update_info()
            self.check_ready()
            dialog.destroy()
        
        ttk.Button(dialog, text="Подтвердить", command=confirm).pack(pady=10)
    
    def select_all_agents(self):
        """Выбрать все агенты"""
        if not self.available_sheets:
            messagebox.showwarning("Внимание", "Сначала загрузите список листов из Google Sheets")
            return
        
        self.selected_agents = self.available_sheets.copy()
        self.sheets_label.config(text=f"Выбрано операторов: {len(self.selected_agents)} (все)", fg="green")
        self.update_info()
        self.check_ready()
    
    def clear_agents(self):
        """Очистить выбор агентов"""
        self.selected_agents = []
        if self.available_sheets:
            self.sheets_label.config(text=f"Операторы не выбраны (доступно: {len(self.available_sheets)})", fg="gray")
        else:
            self.sheets_label.config(text="Сначала загрузите список листов", fg="gray")
        self.update_info()
        self.check_ready()
    
    def select_incident_files(self):
        """Выбор файлов с данными 112"""
        filenames = filedialog.askopenfilenames(
            title="Выберите Excel файлы с данными 112",
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        
        if filenames:
            self.incident_files = list(filenames)
            self.incident_label.config(text=f"Выбрано файлов: {len(filenames)}", fg="green")
            self.update_info()
            self.check_ready()
    
    def select_incident_folder(self):
        """Выбор папки с файлами 112 - загрузит все Excel файлы из папки"""
        folder = filedialog.askdirectory(
            title="Выберите папку с Excel файлами 112"
        )
        
        if folder:
            folder_path = Path(folder)
            # Ищем все Excel файлы в папке
            excel_files = list(folder_path.glob("*.xlsx")) + list(folder_path.glob("*.xls"))
            
            if excel_files:
                self.incident_files = [str(f) for f in excel_files]
                self.incident_label.config(
                    text=f"Из папки загружено файлов: {len(excel_files)}",
                    fg="green"
                )
                self.update_info()
                self.check_ready()
            else:
                messagebox.showwarning(
                    "Внимание",
                    f"В выбранной папке не найдено Excel файлов (.xlsx, .xls)"
                )
    
    def create_incoming_folder(self):
        """Создание папки incoming_data/112_files для загрузки файлов"""
        incoming_folder = self.base_dir / "incoming_data" / "112_files"
        incoming_folder.mkdir(parents=True, exist_ok=True)
        
        # Проверяем наличие файлов
        excel_files = list(incoming_folder.glob("*.xlsx")) + list(incoming_folder.glob("*.xls"))
        
        message = f"Папка создана:\n{incoming_folder}\n\n"
        
        if excel_files:
            message += f"В папке найдено файлов: {len(excel_files)}\n\nЗагрузить эти файлы для обработки?"
            
            if messagebox.askyesno("Папка готова", message):
                self.incident_files = [str(f) for f in excel_files]
                self.incident_label.config(
                    text=f"Загружено из incoming: {len(excel_files)} файлов",
                    fg="green"
                )
                self.update_info()
                self.check_ready()
        else:
            message += "В папке пока нет файлов.\n\nПоместите Excel файлы с данными 112 в эту папку,\nзатем снова нажмите эту кнопку для загрузки.\n\nОткрыть папку?"
            
            if messagebox.askyesno("Папка создана", message):
                os.startfile(incoming_folder)
    
    def update_info(self):
        """Обновление информации о выбранных файлах"""
        self.info_text.delete(1.0, tk.END)
        
        if self.available_sheets:
            self.info_text.insert(tk.END, f"📊 Доступно листов в Google Sheets: {len(self.available_sheets)}\n\n")
        
        if self.selected_agents:
            self.info_text.insert(tk.END, f"✅ Выбрано операторов ({len(self.selected_agents)}):\n")
            for sheet_info in self.selected_agents:
                self.info_text.insert(tk.END, f"   • {sheet_info['display_name']}\n")
            self.info_text.insert(tk.END, "\n")
        
        if self.incident_files:
            self.info_text.insert(tk.END, f"📋 Файлы 112 ({len(self.incident_files)}):\n")
            for f in self.incident_files:
                try:
                    file_path = Path(f)
                    if file_path.exists():
                        file_size = file_path.stat().st_size / 1024  # в KB
                        self.info_text.insert(tk.END, f"   • {file_path.name} ({file_size:.1f} KB)\n")
                    else:
                        self.info_text.insert(tk.END, f"   • {file_path.name} (файл не найден)\n")
                except Exception as e:
                    self.info_text.insert(tk.END, f"   • {Path(f).name}\n")
    
    def check_ready(self):
        """Проверка готовности к обработке"""
        if self.selected_agents and self.incident_files:
            self.process_btn.config(state="normal")
        else:
            self.process_btn.config(state="disabled")
    
    def start_processing(self):
        """Запуск обработки в отдельном потоке"""
        self.process_btn.config(state="disabled")
        self.progress.pack(fill="x", pady=5)
        self.progress.start(10)
        
        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()
    
    def update_progress(self, message):
        """Обновление статуса обработки"""
        self.progress_label.config(text=message)
        self.root.update()
    
    def authenticate_google(self):
        """Аутентификация в Google Sheets API"""
        creds = None
        
        if self.token_file.exists():
            creds = Credentials.from_authorized_user_file(str(self.token_file), self.scopes)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(str(self.credentials_file), self.scopes)
                creds = flow.run_local_server(port=0)
            
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        
        return build('sheets', 'v4', credentials=creds)
    
    def collect_sheets_data(self, service, selected_sheets):
        """Сбор данных из Google Sheets от выбранных листов"""
        all_data = []
        
        for sheet_info in selected_sheets:
            try:
                self.update_progress(f"Загрузка: {sheet_info['display_name']}...")
                
                # Читаем данные с листа
                range_name = f"'{sheet_info['sheet_name']}'!A1:Z10000"
                result = service.spreadsheets().values().get(
                    spreadsheetId=sheet_info['doc_id'],
                    range=range_name
                ).execute()
                
                values = result.get('values', [])
                if not values or len(values) < 2:
                    continue
                
                # Первая строка - заголовки
                headers = values[0]
                num_headers = len(headers)
                data_rows = values[1:]
                
                # Нормализация данных - приводим все строки к одинаковому количеству колонок
                normalized_rows = []
                for row in data_rows:
                    if len(row) < num_headers:
                        # Дополняем строку пустыми значениями
                        row = row + [''] * (num_headers - len(row))
                    elif len(row) > num_headers:
                        # Обрезаем лишние колонки
                        row = row[:num_headers]
                    normalized_rows.append(row)
                
                # Создаём DataFrame
                df_sheet = pd.DataFrame(normalized_rows, columns=headers)
                df_sheet['Документ'] = sheet_info['doc_name']  # Добавляем источник
                df_sheet['Лист'] = sheet_info['sheet_name']  # Добавляем название листа
                
                all_data.append(df_sheet)
                
            except Exception as e:
                print(f"Ошибка при чтении {sheet_info['display_name']}: {e}")
                continue
        
        if not all_data:
            raise Exception("Не удалось загрузить данные ни от одного оператора")
        
        # Объединяем все данные
        df_combined = pd.concat(all_data, ignore_index=True)
        
        return df_combined
    
    def process_data(self):
        """Основная логика обработки данных"""
        try:
            # Создаём папку на рабочем столе с датой
            desktop = Path.home() / "Desktop"
            date_str = datetime.now().strftime("%Y-%m-%d_%H-%M")
            self.output_folder = desktop / f"Отчёты_112_{date_str}"
            self.output_folder.mkdir(exist_ok=True)
            
            self.update_progress("Подключение к Google Sheets...")
            
            # 1. Аутентификация и загрузка данных Google Sheets
            service = self.authenticate_google()
            
            self.update_progress(f"Загрузка данных от {len(self.selected_agents)} агентов...")
            df_sheets = self.collect_sheets_data(service, self.selected_agents)
            
            self.update_progress(f"Загружено записей из Sheets: {len(df_sheets):,}")
            
            # Переименовываем Колонка_2 в Инцидент_Sheets
            if 'Колонка_2' in df_sheets.columns:
                df_sheets = df_sheets.rename(columns={'Колонка_2': 'Инцидент_Sheets'})
            
            # Нормализация инцидентов
            df_sheets['Инцидент_Sheets_norm'] = df_sheets['Инцидент_Sheets'].astype(str).str.strip().str.upper()
            
            self.update_progress(f"Загружено записей из Sheets: {len(df_sheets):,}")
            
            # Определяем наличие жалоб
            if 'Колонка_22' in df_sheets.columns:
                df_sheets['Есть_жалоба'] = df_sheets['Колонка_22'].notna() & (df_sheets['Колонка_22'].astype(str).str.strip() != '')
            else:
                df_sheets['Есть_жалоба'] = False
            
            self.update_progress(f"Загружено записей из Sheets: {len(df_sheets):,}")
            
            # 2. Загрузка данных 112
            self.update_progress("Загрузка файлов 112...")
            
            all_112_data = []
            total_rows_loaded = 0
            
            for idx, file_path in enumerate(self.incident_files, 1):
                try:
                    self.update_progress(f"Загрузка файла {idx}/{len(self.incident_files)}: {Path(file_path).name}...")
                    df_temp = pd.read_excel(file_path)
                    
                    # Проверяем, это файл журнала (есть колонка Служба)
                    if 'Служба' in df_temp.columns:
                        total_rows_loaded += len(df_temp)
                        all_112_data.append(df_temp)
                    else:
                        print(f"Пропущен файл {Path(file_path).name}: отсутствует колонка 'Служба'")
                except Exception as e:
                    print(f"Ошибка при чтении {Path(file_path).name}: {e}")
            
            if not all_112_data:
                raise Exception("Не найдено файлов журнала 112 (с колонкой 'Служба')")
            
            self.update_progress(f"Объединение {len(all_112_data)} файлов (всего строк: {total_rows_loaded:,})...")
            df_112 = pd.concat(all_112_data, ignore_index=True)
            
            # Многоступенчатое удаление дубликатов
            before_dedup = len(df_112)
            
            # Этап 1: Удаление полностью идентичных строк (построчное сравнение)
            self.update_progress("Удаление полных дубликатов между файлами...")
            df_112 = df_112.drop_duplicates()
            full_dupes_removed = before_dedup - len(df_112)
            
            # Этап 2: Удаление дубликатов по ключевым полям с сохранением первого вхождения
            self.update_progress("Удаление дубликатов по ключевым полям (Инцидент, Карта, Служба, Телефон)...")
            key_cols = ['Инцидент', 'Карта', 'Служба', 'Телефон']
            existing_cols = [col for col in key_cols if col in df_112.columns]
            
            if existing_cols:
                before_key_dedup = len(df_112)
                df_112 = df_112.drop_duplicates(subset=existing_cols, keep='first')
                key_dupes_removed = before_key_dedup - len(df_112)
            else:
                key_dupes_removed = 0
            
            # Этап 3: Удаление дубликатов по инциденту (если одна и та же заявка в разных файлах)
            # ВАЖНО: Делаем это ДО переименования колонок!
            if 'Инцидент' in df_112.columns:
                self.update_progress("Удаление дубликатов по номеру инцидента...")
                before_incident_dedup = len(df_112)
                # Сортируем по дате/времени если есть, чтобы оставить самую свежую запись
                if 'Дата' in df_112.columns or 'Время' in df_112.columns:
                    sort_cols = []
                    if 'Дата' in df_112.columns:
                        sort_cols.append('Дата')
                    if 'Время' in df_112.columns:
                        sort_cols.append('Время')
                    try:
                        df_112 = df_112.sort_values(sort_cols, ascending=False)
                    except:
                        pass  # Если сортировка не удалась, продолжаем без неё
                
                df_112 = df_112.drop_duplicates(subset=['Инцидент'], keep='first')
                incident_dupes_removed = before_incident_dedup - len(df_112)
            else:
                incident_dupes_removed = 0
            
            total_removed = before_dedup - len(df_112)
            
            dedup_info = f"""Загружено записей 112: {len(df_112):,}
            
            📊 Дедупликация:
              • Загружено из файлов: {total_rows_loaded:,}
              • Полных дубликатов удалено: {full_dupes_removed:,}
              • По ключевым полям удалено: {key_dupes_removed:,}
              • По инцидентам удалено: {incident_dupes_removed:,}
              • Всего удалено дубликатов: {total_removed:,}
              • Осталось уникальных записей: {len(df_112):,}"""
            
            self.update_progress(dedup_info)
            print("\n" + dedup_info)
            
            # Переименовываем колонки 112
            rename_map = {
                'Инцидент': 'Инцидент_112',
                'Карта': 'Карта_112',
                'Телефон': 'Телефон_112',
                'Статус': 'Статус_112',
                'Оператор': 'Оператор_112',
                'Район': 'Район_112'
            }
            df_112 = df_112.rename(columns={k: v for k, v in rename_map.items() if k in df_112.columns})
            
            # Нормализация инцидентов в 112
            if 'Инцидент_112' in df_112.columns:
                df_112['Инцидент_112_norm'] = df_112['Инцидент_112'].astype(str).str.strip().str.upper()
            
            # Переименование статусов
            status_map = {
                'Не отвечает': 'Не удалось дозвониться',
                'Не берет трубку': 'Не удалось дозвониться',
                'Сбрасывает': 'Не удалось дозвониться',
                'Недоступен': 'Не удалось дозвониться'
            }
            if 'Статус_112' in df_112.columns:
                df_112['Статус_112'] = df_112['Статус_112'].replace(status_map)
            
            # 3. Сопоставление данных
            self.update_progress("Сопоставление данных по инцидентам...")
            
            df_matched = df_sheets.merge(
                df_112,
                left_on='Инцидент_Sheets_norm',
                right_on='Инцидент_112_norm',
                how='inner'
            )
            
            self.update_progress(f"Сопоставлено записей: {len(df_matched):,}")
            
            # 4. Применение бизнес-логики
            self.update_progress("Применение логики категоризации...")
            
            df_matched['Категория'] = 'Не определено'
            
            # Для жалоб - смотрим на службу
            complaints_mask = df_matched['Есть_жалоба'] == True
            
            for service_code in df_matched['Служба'].unique():
                service_mask = df_matched['Служба'] == service_code
                
                # Отрицательно для жалоб этой службы
                df_matched.loc[complaints_mask & service_mask, 'Категория'] = 'Отрицательно'
                
                # Положительно для НЕ жалоб этой службы
                df_matched.loc[~complaints_mask & service_mask, 'Категория'] = 'Положительно'
            
            # 5. Сохранение сопоставленных данных
            matched_file = self.output_folder / f"СОПОСТАВЛЕНИЕ_{date_str}.csv"
            df_matched.to_csv(matched_file, index=False, encoding='utf-8-sig')
            
            # 6. Создание отчётов по службам
            self.update_progress("Создание отчётов по службам...")
            
            services_folder = self.output_folder / "службы_детально"
            services_folder.mkdir(exist_ok=True)
            
            service_names = {
                101: 'Пожарная',
                102: 'Скорая помощь',
                103: 'Газовая',
                104: 'Аварийная'
            }
            
            summary_data = []
            
            for service_code in sorted(df_matched['Служба'].unique()):
                df_service = df_matched[df_matched['Служба'] == service_code].copy()
                
                # Получаем уникальные инциденты для подсчёта
                df_service_unique = df_service.drop_duplicates(subset=['Инцидент_Sheets'], keep='first')
                
                service_name = service_names.get(service_code, f'Служба {service_code}')
                
                # Считаем на уникальных инцидентах
                positive_count = (df_service_unique['Категория'] == 'Положительно').sum()
                negative_count = (df_service_unique['Категория'] == 'Отрицательно').sum()
                
                # Создаём Excel файл
                excel_file = services_folder / f"СЛУЖБА_{service_code}_{service_name}_{date_str}.xlsx"
                
                with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                    # Лист 1: СВОДКА
                    summary = pd.DataFrame({
                        'Показатель': [
                            'Уникальных инцидентов',
                            'Детальных записей',
                            '',
                            'Положительно',
                            'Отрицательно',
                            '',
                            'Процент положительных',
                            'Процент отрицательных'
                        ],
                        'Значение': [
                            len(df_service_unique),
                            len(df_service),
                            '',
                            positive_count,
                            negative_count,
                            '',
                            f"{positive_count/len(df_service_unique)*100:.1f}%" if len(df_service_unique) > 0 else "0%",
                            f"{negative_count/len(df_service_unique)*100:.1f}%" if len(df_service_unique) > 0 else "0%"
                        ]
                    })
                    summary.to_excel(writer, sheet_name='СВОДКА', index=False)
                    
                    # Лист 2: Положительные
                    df_positive = df_service[df_service['Категория'] == 'Положительно']
                    if len(df_positive) > 0:
                        df_positive.to_excel(writer, sheet_name='ПОЛОЖИТЕЛЬНЫЕ', index=False)
                    
                    # Лист 3: Жалобы
                    df_negative = df_service[df_service['Категория'] == 'Отрицательно']
                    if len(df_negative) > 0:
                        df_negative.to_excel(writer, sheet_name='ЖАЛОБЫ', index=False)
                        # Также сохраняем CSV с жалобами
                        csv_file = services_folder / f"СЛУЖБА_{service_code}_ЖАЛОБЫ_{date_str}.csv"
                        df_negative.to_csv(csv_file, index=False, encoding='utf-8-sig')
                
                summary_data.append({
                    'Служба': f"{service_code} - {service_name}",
                    'Уникальных инцидентов': len(df_service_unique),
                    'Детальных записей': len(df_service),
                    'Положительно': positive_count,
                    'Отрицательно': negative_count,
                    '% Положительных': f"{positive_count/len(df_service_unique)*100:.1f}%" if len(df_service_unique) > 0 else "0%"
                })
            
            # 7. Сводный отчёт по всем службам
            summary_df = pd.DataFrame(summary_data)
            summary_file = self.output_folder / f"СВОДКА_ПО_СЛУЖБАМ_{date_str}.xlsx"
            summary_df.to_excel(summary_file, index=False)
            
            # Также сохраняем текстовый отчёт
            txt_file = self.output_folder / f"СВОДНЫЙ_ОТЧЁТ_{date_str}.txt"
            with open(txt_file, 'w', encoding='utf-8') as f:
                f.write("="*80 + "\n")
                f.write("СВОДНЫЙ ОТЧЁТ ПО ОБРАБОТКЕ ДАННЫХ\n")
                f.write("="*80 + "\n\n")
                f.write(f"Дата обработки: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write(f"Выбрано операторов: {len(self.selected_agents)}\n")
                for sheet_info in self.selected_agents:
                    f.write(f"  • {sheet_info['display_name']}\n")
                f.write(f"\nФайлов 112: {len(self.incident_files)}\n\n")
                f.write(f"Загружено из Sheets: {len(df_sheets):,} записей\n")
                f.write(f"Загружено из 112: {len(df_112):,} записей\n")
                f.write(f"Сопоставлено: {len(df_matched):,} записей\n")
                f.write(f"Уникальных инцидентов: {df_matched['Инцидент_Sheets'].nunique():,}\n\n")
                f.write("="*80 + "\n")
                f.write("ПО СЛУЖБАМ:\n")
                f.write("="*80 + "\n\n")
                f.write(summary_df.to_string(index=False))
            
            # Завершение
            self.progress.stop()
            self.progress.pack_forget()
            self.update_progress("✅ Обработка завершена!")
            
            # Показываем результат
            result_message = f"""
Обработка завершена успешно!

📊 Обработано:
  • Google Sheets: {len(df_sheets):,} записей
  • 112 данные: {len(df_112):,} записей
  • Сопоставлено: {len(df_matched):,} записей
  • Уникальных инцидентов: {df_matched['Инцидент_Sheets'].nunique():,}

📁 Результаты сохранены в:
{self.output_folder}

Созданные файлы:
  • Сопоставленные данные (CSV)
  • Отчёты по каждой службе (Excel)
  • Сводный отчёт (Excel + TXT)
  • CSV файлы с жалобами

Открыть папку с результатами?
"""
            
            if messagebox.askyesno("Готово", result_message):
                os.startfile(self.output_folder)
            
            self.process_btn.config(state="normal")
            
        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.update_progress("❌ Ошибка при обработке")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n\n{str(e)}")
            self.process_btn.config(state="normal")


def main():
    root = tk.Tk()
    app = DataProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
