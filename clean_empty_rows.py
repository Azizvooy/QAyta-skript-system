#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Удаляет пустые строки во всех листах, кроме листов с названием FIKSA.
"""

from pathlib import Path
import os
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

BASE_DIR = Path(__file__).parent
SERVICE_ACCOUNT_FILE = BASE_DIR / "config" / "service_account.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_IDS = [
    "18y_QSol_XIZiaKGdoc64-tqerxYXg1kwmO7mmxo21rQ",
    "1JG9RqC64MZbP63HCSa0avp55zXP8nEqzsiuQbbSi9Do",
    "1Ld37ljkNect6iLV0X0QztTJC7vhgPDkaRXTJGVYOtfg",
    "16AK6_7FcWeksg2KaLfuzxo6sOGxsORw9hU-lycgeGiM",
    "1h-EZltuaagu2dyFj0wf-lch2qIhXKaycyCM4MYwLa4U",
    "18WH4ocx371qXsuPIf00hvKR1Kc7u1cd2ZTOUsWS377U",
    "1iSH66bKucYRU7St8Ch_-NThYay3C-tD_p_VlLpVS7WA",
    "1toIzc0CpyQIditC9KqeMVf3II2CWAW4c96pqp1b1t6s",
    "1nrlXxHexPwEBJCyXUt6ehLSMUoC11A6tTr8B5MK1Tuw",
    "11cK2I6pQ_hrHMrbs0AhKmqLLiCbOV00R7tboG0CdrS0",
    "1-T0dKoRATQ4uWmYYli6qU7Jq9FvblD7AGW2wQZFoUB0",
    "10aFfdXjkLNlt_H0D9e4VJnxFJalGX5R9zLzAf8XpQDw",
    "1Gb0RJDqr-Z34D9dHXfPVcGJguKP46YKgPq1qgT_A4nI",
    "1-BdML7lK0fW3vrcxl8yTpgL9FpeM2NUotigf6JIN4CQ",
    "12EVYbShzbwujbqm42rkXZACrRbblo_2Ls-8XmUnID2I",
    "1YmNqdrkLeQBH5Nnq_gVdeVs3LDHArWXfHjo1BBYfpD4",
    "1mJfVK1dCSIMV4ME2lHXUyT1NXD0qaQgr_1pKqMPgCe4",
    "1MLsLwaimRSR6Gdhcm9fuatNs1B-kTgVDfUfbm6BSCjY",
    "12hSjYYlTj9DrVq4PTkI9IwTD4radEgG-Z0bYpgddXiU",
    "1_Ch5eolJHF5JQdSH7uLweBYvIZ1DomhED7pdWRdn7oI",
    "1uhvZlw1GEbMdDi5sGRJoPkd5tyvKnvob0XU0uKl99Z4",
    "1xohqcQR6vpLcmvWPNgbIPO1UCDQJmPILQ29yUpnH_A0",
    "1yDlr5nqVkoEpzPDdKRFyHTxSYJwnlyzBQ_NaJOev300",
    "1D6EIWhpH-QgjL1HvVqX54cQQl_52GJs0oAXWNQF8FKc",
    "1mP6RJtA918WUi8zq7N2RmPh4jMJfGa41UZT0mu3U4nQ",
    "1LbRFZb3830m77GKIBVipifmW6D0kCxWLP-Etku6sYts",
    "1XmQDC7hk0VYV1TQ9ETf9ZJfll5On41ZD2742H5fRmT0",
    "1j_VMVVb8CkM883y8nw2b1BHbTCCFw_43KkMZWhEK1SM",
    "1QMCAddnW5qn5OG9awAyI7Jqeo5Jzdh3mvEHuwqVBHsU",
    "1LjNHy0nsNqjeHRoAfCbGIRh_0QsLoSWSLPwRt4pik58",
    "1Ii1LlQRHtq8dqyZHCtFkCNti37fFK5ff9qbPxNHhcVw",
    "1S7oJBkx9NjsXramXxYeDq36zvN-3a--9f9KhhilfiNc",
    "1jJ8nz7lzFOgz40bN12kX1cyjkhCcquba6H9QHn8kldE",
    "1UaWDeG1pcNbGvvMrVSaRKGX6nqNdEmpQ92mueII1VRE",
    "1YuRCpm_iZkuK-eVqJ3EN5rCgEBOSgaY3qg5shaszb9A",
]


def build_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=credentials)


def request_with_retry(func, max_retries=6, base_delay=1.5):
    for attempt in range(1, max_retries + 1):
        try:
            return func()
        except HttpError as e:
            status = e.resp.status if e.resp else None
            if status in (429, 500, 502, 503, 504):
                sleep_time = base_delay * attempt
                print(f"  ⏳ Лимит/ошибка {status}. Жду {sleep_time:.1f}с и повторяю...", flush=True)
                time.sleep(sleep_time)
                continue
            raise
    return None


def find_empty_ranges(values):
    """Возвращает список (start, end) пустых строк (0-based, end exclusive)."""
    empty_rows = []
    for idx, row in enumerate(values):
        if idx == 0:
            continue  # не трогаем заголовок
        if not row or all(str(cell).strip() == "" for cell in row):
            empty_rows.append(idx)

    if not empty_rows:
        return []

    ranges = []
    start = empty_rows[0]
    prev = empty_rows[0]
    for i in empty_rows[1:]:
        if i == prev + 1:
            prev = i
            continue
        ranges.append((start, prev + 1))
        start = prev = i
    ranges.append((start, prev + 1))
    return ranges


def main():
    service = build_service()

    start_from = int(os.environ.get("START_FROM", "1"))
    for doc_index, spreadsheet_id in enumerate(SPREADSHEET_IDS, 1):
        if doc_index < start_from:
            continue
        print(f"[{doc_index}/{len(SPREADSHEET_IDS)}] Документ {spreadsheet_id[:8]}...")

        def get_meta():
            return service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        metadata = request_with_retry(get_meta)
        if not metadata:
            print("  ❌ Не удалось получить метаданные")
            continue

        sheets = metadata.get("sheets", [])
        for sheet in sheets:
            props = sheet.get("properties", {})
            title = props.get("title", "")
            sheet_id = props.get("sheetId")
            row_count = props.get("gridProperties", {}).get("rowCount", 0)

            if "fiksa" in title.lower():
                continue
            if row_count <= 1:
                continue

            range_name = f"'{title}'!B1:L{row_count}"

            def get_values():
                return service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=range_name
                ).execute()

            values_resp = request_with_retry(get_values)
            if not values_resp:
                print(f"  ⚠️  {title}: не удалось получить данные")
                continue

            values = values_resp.get("values", [])
            if not values:
                continue

            empty_ranges = find_empty_ranges(values)
            if not empty_ranges:
                continue

            # Удаляем снизу вверх
            requests = []
            for start, end in reversed(empty_ranges):
                requests.append({
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": start,
                            "endIndex": end,
                        }
                    }
                })

            print(f"  🧹 {title}: удаляю {sum(e - s for s, e in empty_ranges)} пустых строк")

            def do_update():
                return service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": requests}
                ).execute()

            try:
                request_with_retry(do_update)
            except HttpError as e:
                message = getattr(e, "_get_reason", lambda: "")()
                if "protected" in str(message).lower() or "protected" in str(e).lower():
                    print(f"  🔒 {title}: защищено, пропускаю")
                    continue
                raise
            time.sleep(0.4)

    print("Готово.")


if __name__ == "__main__":
    main()
