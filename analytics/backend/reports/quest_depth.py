import time
import calendar
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import requests
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from backend.reports.common import (
    GRAFANA_URL,
    DATASOURCE_UID,
    get_headers,
    get_time_boundaries,
    normalize_park_name,
)

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

MONTH_NAMES_RU = [
    "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

MONTH_SHORT_RU = [
    "", "Янв", "Фев", "Мар", "Апр", "Май", "Июн",
    "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"
]

def fetch_chunk_metrics(start_iso: str, stop_iso: str, retries: int = 3):
    """Выполняет запрос total_tasks и unique_users для заданного интервала времени"""
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

    last_err = None
    for attempt in range(retries):
        try:
            response = requests.post(f"{GRAFANA_URL}/api/ds/query", headers=get_headers(), json=payload, timeout=60)
            response.raise_for_status()
            resp_json = response.json()
            last_err = None
            break
        except Exception as e:
            last_err = e
            if attempt < retries - 1:
                time.sleep(1.5)

    if last_err is not None:
        raise last_err

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

def style_sheet(ws, is_percentage_matrix=False, is_depth_matrix=False):
    """Стилизует лист в фирменном стиле Hello Park"""
    orange = PatternFill("solid", fgColor="FF6B00")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    gray = PatternFill("solid", fgColor="E0E0E0")
    bold_font = Font(bold=True)
    green = PatternFill("solid", fgColor="D5F5E3")

    last_row = ws.max_row
    last_col = ws.max_column

    # Заголовки
    for col in range(1, last_col + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = orange
        cell.font = header_font
        cell.alignment = center

    # Данные строк
    for r in range(2, last_row):
        ws.cell(row=r, column=1).alignment = left
        ws.cell(row=r, column=2).alignment = center
        for c in range(3, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = center
            if is_percentage_matrix:
                cell.fill = green
                cell.number_format = '0.0"%"'
            elif is_depth_matrix:
                cell.number_format = '0.0'
            else:
                if c == 3:
                    cell.number_format = '0.0'
                elif c == 4:
                    cell.fill = green
                    cell.number_format = '0.0"%"'

    # Строка ИТОГО
    for col in range(1, last_col + 1):
        cell = ws.cell(row=last_row, column=col)
        cell.fill = gray
        cell.font = bold_font
        cell.alignment = center
        if is_percentage_matrix:
            if col >= 3:
                cell.number_format = '0.0"%"'
        elif is_depth_matrix:
            if col >= 3:
                cell.number_format = '0.0'
        else:
            if col == 3:
                cell.number_format = '0.0'
            elif col == 4:
                cell.number_format = '0.0"%"'

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)
    ws.column_dimensions["A"].width = 22

def build_standard_table(tasks_map: dict, users_map: dict, parks: list):
    """Формирует классическую таблицу из 4 колонок (Парк, Базис, Глубина, Индекс) + строка ИТОГО"""
    rows = []
    total_tasks = 0
    total_users = 0
    weighted_pei_sum = 0.0

    for park in parks:
        basis = PARK_BASIS.get(park, 0)
        tasks = tasks_map.get(park, 0)
        users = users_map.get(park, 0)

        if users > 0:
            avg_depth = tasks / users
            pei = (avg_depth / basis * 100.0) if basis > 0 else 0.0
        else:
            avg_depth = 0.0
            pei = 0.0

        total_tasks += tasks
        total_users += users
        weighted_pei_sum += pei * users

        rows.append({
            "Парк": park,
            "Базис заданий": basis,
            "Средняя глубина (заданий)": round(avg_depth, 1),
            "Индекс охвата (%)": round(pei, 1),
        })

    if total_users > 0:
        net_avg_depth = round(total_tasks / total_users, 1)
        net_pei = round(weighted_pei_sum / total_users, 1)
    else:
        net_avg_depth = 0.0
        net_pei = 0.0

    summary_row = {
        "Парк": "ИТОГО",
        "Базис заданий": "—",
        "Средняя глубина (заданий)": net_avg_depth,
        "Индекс охвата (%)": net_pei,
    }

    df_data = pd.DataFrame(rows)
    df_summary = pd.DataFrame([summary_row])
    return pd.concat([df_data, df_summary], ignore_index=True)

def generate_quest_depth(date_val: str, period_type: str, selected_parks: list, output_path: str):
    parks_to_include = sorted([p for p in selected_parks if p in PARK_BASIS and p != "OFFICE"])
    if not parks_to_include:
        parks_to_include = sorted(list(PARK_BASIS.keys()))

    if period_type == "year":
        year = int(date_val)

        def fetch_month_worker(m):
            s, e, _ = get_time_boundaries(f"{year}-{m:02d}", "month")
            t, u = fetch_chunk_metrics(s, e)
            return m, t, u

        with ThreadPoolExecutor(max_workers=6) as executor:
            month_results = list(executor.map(fetch_month_worker, range(1, 13)))

        month_results.sort(key=lambda x: x[0])
        # Оставляем месяцы, в которых есть данные
        active_months = [r for r in month_results if sum(r[1].values()) > 0 or sum(r[2].values()) > 0]
        if not active_months:
            active_months = month_results[:datetime.now().month]

        # Агрегация за год
        year_tasks = {}
        year_users = {}
        for m, t_map, u_map in active_months:
            for p in parks_to_include:
                year_tasks[p] = year_tasks.get(p, 0) + t_map.get(p, 0)
                year_users[p] = year_users.get(p, 0) + u_map.get(p, 0)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # 1. Лист «Динамика охвата (%)»
            pei_matrix_rows = []
            for p in parks_to_include:
                b = PARK_BASIS[p]
                row = {"Парк": p, "Базис заданий": b}
                for m, t_map, u_map in active_months:
                    col = f"{MONTH_SHORT_RU[m]} {year}"
                    t = t_map.get(p, 0)
                    u = u_map.get(p, 0)
                    val = (t / u / b * 100.0) if (u > 0 and b > 0) else 0.0
                    row[col] = round(val, 1)
                y_t = year_tasks.get(p, 0)
                y_u = year_users.get(p, 0)
                row["За год"] = round((y_t / y_u / b * 100.0) if (y_u > 0 and b > 0) else 0.0, 1)
                pei_matrix_rows.append(row)

            pei_summary = {"Парк": "ИТОГО", "Базис заданий": "—"}
            for m, t_map, u_map in active_months:
                col = f"{MONTH_SHORT_RU[m]} {year}"
                m_u = sum(u_map.get(p, 0) for p in parks_to_include)
                w_sum = sum((t_map.get(p, 0) / PARK_BASIS[p]) for p in parks_to_include if PARK_BASIS.get(p, 0) > 0)
                pei_summary[col] = round(w_sum / m_u * 100.0, 1) if m_u > 0 else 0.0

            tot_y_u = sum(year_users.get(p, 0) for p in parks_to_include)
            w_y_sum = sum((year_tasks.get(p, 0) / PARK_BASIS[p]) for p in parks_to_include if PARK_BASIS.get(p, 0) > 0)
            pei_summary["За год"] = round(w_y_sum / tot_y_u * 100.0, 1) if tot_y_u > 0 else 0.0
            pei_matrix_rows.append(pei_summary)

            df_pei = pd.DataFrame(pei_matrix_rows)
            df_pei.to_excel(writer, sheet_name="Динамика охвата (%)", index=False)
            style_sheet(writer.sheets["Динамика охвата (%)"], is_percentage_matrix=True)

            # 2. Лист «Динамика глубины»
            depth_matrix_rows = []
            for p in parks_to_include:
                b = PARK_BASIS[p]
                row = {"Парк": p, "Базис заданий": b}
                for m, t_map, u_map in active_months:
                    col = f"{MONTH_SHORT_RU[m]} {year}"
                    t = t_map.get(p, 0)
                    u = u_map.get(p, 0)
                    val = (t / u) if u > 0 else 0.0
                    row[col] = round(val, 1)
                y_t = year_tasks.get(p, 0)
                y_u = year_users.get(p, 0)
                row["За год"] = round((y_t / y_u) if y_u > 0 else 0.0, 1)
                depth_matrix_rows.append(row)

            depth_summary = {"Парк": "ИТОГО", "Базис заданий": "—"}
            for m, t_map, u_map in active_months:
                col = f"{MONTH_SHORT_RU[m]} {year}"
                m_t = sum(t_map.get(p, 0) for p in parks_to_include)
                m_u = sum(u_map.get(p, 0) for p in parks_to_include)
                depth_summary[col] = round(m_t / m_u, 1) if m_u > 0 else 0.0

            tot_y_t = sum(year_tasks.get(p, 0) for p in parks_to_include)
            depth_summary["За год"] = round(tot_y_t / tot_y_u, 1) if tot_y_u > 0 else 0.0
            depth_matrix_rows.append(depth_summary)

            df_depth = pd.DataFrame(depth_matrix_rows)
            df_depth.to_excel(writer, sheet_name="Динамика глубины", index=False)
            style_sheet(writer.sheets["Динамика глубины"], is_depth_matrix=True)

            # 3. Лист «Итого {year}»
            df_year = build_standard_table(year_tasks, year_users, parks_to_include)
            year_sheet_title = f"Итого {year}"[:31]
            df_year.to_excel(writer, sheet_name=year_sheet_title, index=False)
            style_sheet(writer.sheets[year_sheet_title])

            # 4. Листы по каждому месяцу
            for m, t_map, u_map in active_months:
                df_month = build_standard_table(t_map, u_map, parks_to_include)
                sheet_title = f"{MONTH_NAMES_RU[m]} {year}"[:31]
                df_month.to_excel(writer, sheet_name=sheet_title, index=False)
                style_sheet(writer.sheets[sheet_title])

        return output_path

    else:
        # Для периодов "month", "day" или коротких "custom"
        start_date, stop_date, _ = get_time_boundaries(date_val, period_type)
        total_tasks_map, users_map = fetch_chunk_metrics(start_date, stop_date)

        df_final = build_standard_table(total_tasks_map, users_map, parks_to_include)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            sheet_name = f"Глубина {date_val}"[:31]
            df_final.to_excel(writer, sheet_name=sheet_name, index=False)
            style_sheet(writer.sheets[sheet_name])

        return output_path
