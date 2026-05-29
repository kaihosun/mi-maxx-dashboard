#!/usr/bin/env python3
"""
process_incoming.py
Parse and aggregate Incoming trailer data for MI Technologies MTY MAXX dashboard.
Input: Reporte UPH MI Technologies MTY MAXX  - Incoming (1).csv
Output: JSON to stdout with monthly and grand-total aggregates.
"""

import csv
import json
from datetime import datetime
from collections import defaultdict

CSV_PATH = "/Users/eduardoflores/AI_Projects/Dashboards/Especiales/Reporte UPH MI Technologies MTY MAXX  - Incoming (1).csv"

# ---------------------------------------------------------------------------
# Origen normalization
# ---------------------------------------------------------------------------
ORIGEN_MAP = {
    "Tijuana, B.C.": "Tijuana, BC",
    "Tijuana, BC": "Tijuana, BC",
    "Jhonston, NY": "Johnstown, NY",
    "SPartanburg, SC": "Spartanburg, SC",
    "Garcia, N.L.": "Garcia, NL",
    "Garcia,N.L.": "Garcia, NL",
}

SKIP_ORIGEN = {"0", ""}


def normalize_origen(raw: str):
    """Return normalized origen string, or None if it should be skipped."""
    raw = raw.strip()
    if raw in SKIP_ORIGEN:
        return None
    return ORIGEN_MAP.get(raw, raw)


TIJUANA = "Tijuana, BC"


def is_tijuana(origen: str) -> bool:
    return origen == TIJUANA


# ---------------------------------------------------------------------------
# Tiempo de Descarga parser  ->  decimal hours
# Handles "0:30", "1:30:00", "0:30:00"
# ---------------------------------------------------------------------------
def parse_tiempo(raw: str):
    raw = raw.strip()
    if not raw or raw in {"0", "#DIV/0!"}:
        return None
    parts = raw.split(":")
    try:
        if len(parts) == 2:
            h, m = int(parts[0]), int(parts[1])
            s = 0
        elif len(parts) == 3:
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
        else:
            return None
        total = h + m / 60 + s / 3600
        return total if total > 0 else None
    except (ValueError, IndexError):
        return None


# ---------------------------------------------------------------------------
# Integer field parser
# ---------------------------------------------------------------------------
def parse_int(raw: str):
    raw = raw.strip()
    if not raw or raw in {"0", "#DIV/0!"}:
        return None
    try:
        val = int(raw)
        return val if val > 0 else None
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Date parser — MM/DD/YYYY
# ---------------------------------------------------------------------------
def parse_fecha(raw: str):
    """Return 'YYYY-MM' or None."""
    raw = raw.strip()
    if not raw or raw == "0":
        return None
    try:
        dt = datetime.strptime(raw, "%m/%d/%Y")
        return dt.strftime("%Y-%m")
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Main aggregation
# ---------------------------------------------------------------------------
def main():
    # month-keyed accumulators
    months_data = defaultdict(lambda: {
        "pallets_tijuana": 0,
        "pallets_us": 0,
        "events_tijuana": 0,
        "events_us": 0,
        "total_hours": 0.0,
        "headcount_sum": 0.0,
        "headcount_count": 0,
    })

    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            fecha_raw = row.get("Fecha", "").strip()
            origen_raw = row.get("Origen", "").strip()

            month = parse_fecha(fecha_raw)
            origen = normalize_origen(origen_raw)

            # Skip rows without valid month or origen
            if not month or not origen:
                continue

            qty_tar = parse_int(row.get("Qty Tar", ""))
            tiempo = parse_tiempo(row.get("Tiempo de Descarga", ""))
            headcount = parse_int(row.get("Qty Personal", ""))

            bucket = months_data[month]

            if is_tijuana(origen):
                bucket["events_tijuana"] += 1
                if qty_tar is not None:
                    bucket["pallets_tijuana"] += qty_tar
            else:
                bucket["events_us"] += 1
                if qty_tar is not None:
                    bucket["pallets_us"] += qty_tar

            if tiempo is not None:
                bucket["total_hours"] += tiempo

            if headcount is not None:
                bucket["headcount_sum"] += headcount
                bucket["headcount_count"] += 1

    # Sort months chronologically
    sorted_months = sorted(months_data.keys())

    # Build parallel arrays
    pallets_tijuana = []
    pallets_us = []
    events_tijuana = []
    events_us = []
    total_hours = []
    avg_headcount = []

    for m in sorted_months:
        b = months_data[m]
        pallets_tijuana.append(b["pallets_tijuana"])
        pallets_us.append(b["pallets_us"])
        events_tijuana.append(b["events_tijuana"])
        events_us.append(b["events_us"])
        total_hours.append(round(b["total_hours"], 2))
        if b["headcount_count"] > 0:
            avg_headcount.append(round(b["headcount_sum"] / b["headcount_count"], 2))
        else:
            avg_headcount.append(None)

    # Grand totals
    gt_pallets_tijuana = sum(pallets_tijuana)
    gt_pallets_us = sum(pallets_us)
    gt_pallets_all = gt_pallets_tijuana + gt_pallets_us
    gt_events_tijuana = sum(events_tijuana)
    gt_events_us = sum(events_us)
    gt_events_all = gt_events_tijuana + gt_events_us

    all_hc_sums = sum(months_data[m]["headcount_sum"] for m in sorted_months)
    all_hc_counts = sum(months_data[m]["headcount_count"] for m in sorted_months)
    gt_avg_headcount = round(all_hc_sums / all_hc_counts, 2) if all_hc_counts > 0 else None

    output = {
        "months": sorted_months,
        "pallets_tijuana": pallets_tijuana,
        "pallets_us": pallets_us,
        "events_tijuana": events_tijuana,
        "events_us": events_us,
        "total_hours": total_hours,
        "avg_headcount": avg_headcount,
        "totals": {
            "pallets_tijuana": gt_pallets_tijuana,
            "pallets_us": gt_pallets_us,
            "pallets_all": gt_pallets_all,
            "events_all": gt_events_all,
            "events_tijuana": gt_events_tijuana,
            "avg_headcount_global": gt_avg_headcount,
        },
    }

    print(json.dumps(output, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
