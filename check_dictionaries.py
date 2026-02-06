import pandas as pd
from pathlib import Path
from generate_service_reports_by_incident import load_sheets_maps

print("Загружаем словари...")
incident_status, incident_positive, incident_all_complaints, incident_service_complaints = load_sheets_maps()

incident = '01.AAD4284/26'

print(f"\nПроверка инцидента: {incident}")
print(f"В incident_status: {incident in incident_status}")
if incident in incident_status:
    print(f"  Значение: '{incident_status[incident]}'")

print(f"\nВ incident_positive: {incident in incident_positive}")
if incident in incident_positive:
    print(f"  Значение: '{incident_positive[incident]}'")

print(f"\nВ incident_all_complaints: {incident in incident_all_complaints}")
if incident in incident_all_complaints:
    print(f"  Жалобы: {incident_all_complaints[incident]}")

print(f"\nВ incident_service_complaints:")
for key in incident_service_complaints:
    if key[0] == incident:
        print(f"  Ключ {key}: {incident_service_complaints[key]}")

print(f"\nВсего записей в incident_status: {len(incident_status)}")
print(f"Всего записей в incident_positive: {len(incident_positive)}")
print(f"Всего записей в incident_all_complaints: {len(incident_all_complaints)}")
print(f"Всего записей в incident_service_complaints: {len(incident_service_complaints)}")
