from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import requests
import pandas as pd
from backend.reports.common import (
    GRAFANA_URL,
    DATASOURCE_UID,
    get_headers,
    get_time_boundaries,
    normalize_park_name,
)
from openpyxl.styles import PatternFill, Font, Alignment

PARK_BASIS = {
    "AVIAPARK": 27,
    "Atyrau": 20,
    "BAKU": 26,
    "BOGOTA-NUESTRO": 17,
    "DUBAI": 32,
    "KASPIYSK": 17,
    "MEGA": 18,
    "OMAN": 30,
    "RIVIERA": 15,
    "SAKHALIN": 26,
    "SELIGERSKAYA": 36,
    "SOCHI": 18,
    "VLADIKAVKAZ": 26,
    "VORONEZH": 42,
}

def get_interval_chunks(start_iso: str, stop_iso: str, max_chunk_days: int = 31):
    """
    Разбивает временной интервал на отрезки не более max_chunk_days дней,
    чтобы избежать таймаутов InfluxDB/Grafana при сканировании года или длинных периодов.
    """
    dt_start = datetime.fromisoformat(start_iso.replace("Z", "+00:00"))
    dt_stop = datetime.fromisoformat(stop_iso.replace("Z", "+00:00"))
    chunks = []
    curr = dt_start
    while curr < dt_stop:
        nxt = min(curr + timedelta(days=max_chunk_days), dt_stop)
        chunks.append((
            curr.strftime("%Y-%m-%dT%H:%M:%SZ"),
            nxt.strftime("%Y-%m-%dT%H:%M:%SZ")
        ))
        curr = nxt
    return chunks

def fetch_chunk_metrics(start_iso: str, stop_iso: str):
    flux_total_tasks = f"""
    from(bucket: "Analytics_AvatarBD")
      |> range(start: {start_iso}, stop: {stop_iso})
      |> filter(fn: (r) => r["_measurement"] == "AchievementCounted")
      |> filter(fn: (r) => exists r["id"])
      |> group(columns: ["park"])
      |> count(column: "_value")
      |> yield(name: "total_tasks")
    """

    flux_users = f"""
    from(bucket: "Analytics_AvatarBD")
      |> range(start: {start_iso}, stop: {stop_iso})
      |> filter(fn: (r) => r["_measurement"] == "AchievementCounted")
      |> filter(fn: (r) => exists r["id"])
      |> group(columns: ["park", "id"])
      |> limit(n: 1)
      |> group(columns: ["park"])
      |> count(column: "_value")
      |> yield(name: "unique_users")
    """

    payload = {
        "queries": [
            {"refId": "A", "datasource": {"type": "influxdb", "uid": DATASOURCE_UID}, "query": flux_total_tasks},
            {"refId": "B", "datasource": {"type": "influxdb", "uid": DATASOURCE_UID}, "query": flux_users}
        ],
        "from": "0",
        "to": "9999999999999"
    }

    response = requests.post(f"{GRAFANA_URL}/api/ds/query", headers=get_headers(), json=payload, timeout=60)
    response.raise_for_status()
    resp_json = response.json()

    chunk_tasks = {}
    for frame in resp_json.get("results", {}).get("A", {}).get("frames", []):
        park = "Unknown"
        for field in frame.get("schema", {}).get("fields", []):
            if "park" in field.get("labels", {}):
                park = field["labels"]["park"]
                break
        park = normalize_park_name(park)
        vals = frame.get("data", {}).get("values", [])
        if vals and len(vals[0]) > 0:
            chunk_tasks[park] = chunk_tasks.get(park, 0) + (vals[0][0] or 0)

    chunk_users = {}
    for frame in resp_json.get("results", {}).get("B", {}).get("frames", []):
        park = "Unknown"
        for field in frame.get("schema", {}).get("fields", []):
            if "park" in field.get("labels", {}):
                park = field["labels"]["park"]
                break
        park = normalize_park_name(park)
        vals = frame.get("data", {}).get("values", [])
        if vals and len(vals[0]) > 0:
            chunk_users[park] = chunk_users.get(park, 0) + (vals[0][0] or 0)

    return chunk_tasks, chunk_users

def generate_quest_depth(date_val: str, period_type: str, selected_parks: list, output_path: str):
    start_date, stop_date, _ = get_time_boundaries(date_val, period_type)
    chunks = get_interval_chunks(start_date, stop_date, max_chunk_days=31)

    total_tasks_map = {}
    users_map = {}

    try:
        max_workers = min(6, len(chunks))
        if max_workers <= 1:
            c_tasks, c_users = fetch_chunk_metrics(chunks[0][0], chunks[0][1])
            total_tasks_map = c_tasks
            users_map = c_users
        else:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(fetch_chunk_metrics, s, e) for s, e in chunks]
                for f in futures:
                    c_tasks, c_users = f.result()
                    for p, t in c_tasks.items():
                        total_tasks_map[p] = total_tasks_map.get(p, 0) + t
                    for p, u in c_users.items():
                        users_map[p] = users_map.get(p, 0) + u
    except Exception as e:
        raise Exception(f"Ошибка при запросе данных из Grafana: {e}")

    # Нормализуем выбранные парки
    parks_to_include = sorted([p for p in selected_parks if p in PARK_BASIS and p != "OFFICE"])
    if not parks_to_include:
        parks_to_include = sorted(list(PARK_BASIS.keys()))

    rows = []
    total_tasks_network = 0
    total_users_network = 0
    weighted_pei_sum = 0.0

    for park in parks_to_include:
        basis = PARK_BASIS.get(park, 0)
        tasks = total_tasks_map.get(park, 0)
        users = users_map.get(park, 0)

        if users > 0:
            avg_depth = tasks / users
            pei = (avg_depth / basis * 100.0) if basis > 0 else 0.0
        else:
            avg_depth = 0.0
            pei = 0.0

        total_tasks_network += tasks
        total_users_network += users
        weighted_pei_sum += pei * users

        rows.append({
            "Парк": park,
            "Базис заданий": basis,
            "Средняя глубина (заданий)": round(avg_depth, 1),
            "Индекс охвата (%)": round(pei, 1),
            "_users": users
        })

    # Итоговая строка
    if total_users_network > 0:
        network_avg_depth = round(total_tasks_network / total_users_network, 1)
        network_pei = round(weighted_pei_sum / total_users_network, 1)
    else:
        network_avg_depth = 0.0
        network_pei = 0.0

    summary_row = {
        "Парк": "ИТОГО",
        "Базис заданий": "—",
        "Средняя глубина (заданий)": network_avg_depth,
        "Индекс охвата (%)": network_pei,
        "_users": total_users_network
    }

    df_data = pd.DataFrame(rows)
    df_summary = pd.DataFrame([summary_row])
    df_final = pd.concat([df_data, df_summary], ignore_index=True)
    df_final = df_final.drop(columns=["_users"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheet_name = f"Глубина {date_val}"[:31]
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        # Стилизация в фирменном стиле Hello Park
        orange = PatternFill("solid", fgColor="FF6B00")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        center = Alignment(horizontal="center", vertical="center")
        gray = PatternFill("solid", fgColor="E0E0E0")
        bold_font = Font(bold=True)
        green = PatternFill("solid", fgColor="D5F5E3")

        # Заголовки
        for col in range(1, len(df_final.columns) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = orange
            cell.font = header_font
            cell.alignment = center

        last_row = ws.max_row

        # Данные строк
        for row in range(2, last_row):
            ws.cell(row=row, column=2).alignment = center
            ws.cell(row=row, column=3).alignment = center
            ws.cell(row=row, column=3).number_format = '0.0'
            
            cell_pei = ws.cell(row=row, column=4)
            cell_pei.fill = green
            cell_pei.alignment = center
            cell_pei.number_format = '0.0"%"'

        # Строка ИТОГО
        for col in range(1, len(df_final.columns) + 1):
            cell = ws.cell(row=last_row, column=col)
            cell.fill = gray
            cell.font = bold_font
            cell.alignment = center

        ws.cell(row=last_row, column=3).number_format = '0.0'
        ws.cell(row=last_row, column=4).number_format = '0.0"%"'

        # Ширина колонок
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 28
        ws.column_dimensions["D"].width = 24

    return output_path
