"""
update_data.py

Reads MATA's raw exported CSV/XLSX reports out of the /data folder, computes
the same KPIs shown on the executive dashboard, and writes the result to
data.json at the repo root.

This is the piece that would eventually be replaced by a real API call
(e.g. Trapeze, Genfare, Navineo) once you have direct system access -- for
now it re-parses whatever exports get dropped into /data, which is a
completely valid and common "integration" pattern when a system doesn't
expose a live API yet.

Expected input files (place these in the /data folder, replacing them each
time you have a fresh export):
  - data/mayors_dashboard.csv          (Mayor's Dashboard FY export)
  - data/route_rankings.csv            (Fixed Route Ridership Rankings export)
  - data/stats_update.xlsx             (MATA Stats Update workbook)

Requires: openpyxl  (pip install openpyxl --break-system-packages)

Run manually with:
    python scripts/update_data.py
"""

import csv
import json
from pathlib import Path
from datetime import datetime, timezone

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data"
OUTPUT_FILE = REPO_ROOT / "data.json"

MONTHS = ["Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
          "Jan", "Feb", "Mar", "Apr", "May", "Jun"]


def clean_number(raw):
    """Turn strings like ' 2,549,260 ' or '61.2%' into a float. Returns None if blank."""
    if raw is None:
        return None
    text = raw.strip().replace(",", "").replace('"', "")
    if text == "":
        return None
    is_percent = text.endswith("%")
    text = text.rstrip("%")
    try:
        value = float(text)
    except ValueError:
        return None
    return value


def load_csv_rows(path):
    with open(path, encoding="cp1252", newline="") as f:
        return list(csv.reader(f))


def find_row(rows, label_contains):
    """Return the first row whose first cell contains the given text (case-insensitive)."""
    for row in rows:
        if row and label_contains.lower() in row[0].lower():
            return row
    return None


def parse_mayors_dashboard(path):
    rows = load_csv_rows(path)

    def monthly_and_total(label):
        row = find_row(rows, label)
        if row is None:
            return None, None
        monthly = [clean_number(v) for v in row[1:13]]
        total = clean_number(row[14]) if len(row) > 14 else None
        return monthly, total

    motorbus_ridership, motorbus_total = monthly_and_total("Ridership - MotorBus")
    trolley_ridership, trolley_total = monthly_and_total("Ridership - Steel Rail")
    paratransit_ridership, paratransit_total = monthly_and_total("Ridership - Demand Response")
    ready1_monthly, ready1_total = monthly_and_total("Ready! Zone 1")
    ready2_monthly, ready2_total = monthly_and_total("Ready! Zone 2")
    ready3_monthly, ready3_total = monthly_and_total("Ready! Zone 3")
    groove_monthly, groove_total = monthly_and_total("Groove")

    otp_fixed_row = find_row(rows, "On Time Performance- MotorBus")
    otp_para_row = find_row(rows, "On Time Performance - Demand Response")

    otp_fixed_monthly = [clean_number(v) for v in otp_fixed_row[1:13]] if otp_fixed_row else []
    otp_fixed_goal = clean_number(otp_fixed_row[13]) if otp_fixed_row else None
    otp_para_monthly = [clean_number(v) for v in otp_para_row[1:13]] if otp_para_row else []
    otp_para_goal = clean_number(otp_para_row[13]) if otp_para_row else None

    def latest(monthly):
        vals = [v for v in monthly if v is not None]
        return vals[-1] if vals else None

    totals = [t for t in [motorbus_total, trolley_total, paratransit_total,
                           ready1_total, ready2_total, ready3_total, groove_total]
              if t is not None]

    return {
        "totalRidershipYTD": sum(totals) if totals else None,
        "otpFixedRoute": latest(otp_fixed_monthly),
        "otpFixedRouteGoal": otp_fixed_goal,
        "otpParatransit": latest(otp_para_monthly),
        "otpParatransitGoal": otp_para_goal,
        "ridershipByMode": {
            "months": MONTHS,
            "motorbus": motorbus_ridership or [],
            "trolley": trolley_ridership or [],
            "paratransit": paratransit_ridership or [],
        },
        "otpTrend": {
            "months": MONTHS,
            "fixedRoute": otp_fixed_monthly,
            "paratransit": otp_para_monthly,
        },
    }


def parse_route_rankings(path, fiscal_year_label="FY26"):
    rows = load_csv_rows(path)

    # Find the header row for the target fiscal year block
    start_index = None
    for i, row in enumerate(rows):
        if row and fiscal_year_label.lower() in row[0].lower():
            start_index = i + 1
            break

    if start_index is None:
        return []

    routes = []
    for row in rows[start_index:]:
        if not row or not row[0].strip():
            break  # blank row marks the end of this fiscal year's block
        name = row[0].strip()
        total = clean_number(row[13]) if len(row) > 13 else None
        if total is not None:
            routes.append({"route": name, "riders": int(total)})

    routes.sort(key=lambda r: r["riders"], reverse=True)
    return routes[:10]


def parse_stats_update(path):
    """Reads the 'OTP by Route' and '3mon - OTP' sheets and returns
    per-route on-time performance for the most recent reporting month,
    plus a trend arrow showing direction over the last 3 months."""
    wb = openpyxl.load_workbook(path, data_only=True)
    otp_sheet = wb["OTP by Route"]

    header_row = 3  # row containing the month dates
    data_start_row = 4
    dates = [cell.value for cell in otp_sheet[header_row]]

    latest_by_route = {}
    for r in range(data_start_row, otp_sheet.max_row + 1):
        route = otp_sheet.cell(row=r, column=1).value
        if not route or str(route).strip().lower() == "total":
            continue
        route = str(route).strip()
        # scan right-to-left for the most recent populated month
        for c in range(otp_sheet.max_column, 1, -1):
            value = otp_sheet.cell(row=r, column=c).value
            if value is not None:
                latest_by_route[route] = {
                    "otp": round(value * 100, 1),
                    "asOf": dates[c - 1].strftime("%Y-%m-%d") if dates[c - 1] else None,
                }
                break

    if not latest_by_route:
        return []

    # Only keep routes reporting as of the most recent month (drops
    # discontinued routes that stopped reporting long ago)
    most_recent = max(v["asOf"] for v in latest_by_route.values() if v["asOf"])
    active_routes = {r: v for r, v in latest_by_route.items() if v["asOf"] == most_recent}

    # Merge in 3-month trend arrows where available
    if "3mon - OTP" in wb.sheetnames:
        trend_sheet = wb["3mon - OTP"]
        for r in range(2, trend_sheet.max_row + 1):
            route = trend_sheet.cell(row=r, column=1).value
            if not route:
                continue
            route = str(route).strip()
            if route in active_routes:
                arrow = trend_sheet.cell(row=r, column=6).value
                active_routes[route]["trend"] = arrow

    result = [
        {"route": route, "otp": v["otp"], "trend": v.get("trend", "")}
        for route, v in active_routes.items()
    ]
    result.sort(key=lambda r: r["otp"])
    return result


def main():
    mayors_csv = DATA_DIR / "mayors_dashboard.csv"
    rankings_csv = DATA_DIR / "route_rankings.csv"
    stats_xlsx = DATA_DIR / "stats_update.xlsx"

    if not mayors_csv.exists() or not rankings_csv.exists():
        raise FileNotFoundError(
            "Expected data/mayors_dashboard.csv and data/route_rankings.csv "
            "to exist. Drop your latest CSV exports into the /data folder."
        )

    kpi_data = parse_mayors_dashboard(mayors_csv)
    top_routes = parse_route_rankings(rankings_csv)
    otp_by_route = parse_stats_update(stats_xlsx) if stats_xlsx.exists() else []

    output = {
        "lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        **kpi_data,
        "topRoutes": top_routes,
        "otpByRoute": otp_by_route,
    }

    OUTPUT_FILE.write_text(json.dumps(output, indent=2))
    print(f"Wrote {OUTPUT_FILE} with {len(top_routes)} top routes and "
          f"{len(otp_by_route)} route-level OTP entries.")


if __name__ == "__main__":
    main()
