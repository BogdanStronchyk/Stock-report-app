import os
import re
from typing import Any, Dict, Optional, Tuple
from openpyxl import load_workbook

def parse_range_cell(s: Any) -> Optional[Tuple[Optional[float], Optional[float]]]:
    if s is None or not isinstance(s, str):
        return None

    txt = s.strip()
    if not txt:
        return None

    txt = txt.replace("–", "-").replace("×", "").replace("x", "").strip()

    if any(w in txt.lower() for w in ["expanding", "stable", "contracting", "not primary", "not used", "use"]):
        return None

    mul = 1.0
    if "$" in txt:
        if "b" in txt.lower():
            mul = 1e9
        elif "m" in txt.lower():
            mul = 1e6
        txt = txt.replace("$", "").replace("B", "").replace("b", "").replace("M", "").replace("m", "").strip()

    txt = txt.replace("%", "").strip()

    m = re.match(r"^<\s*([0-9\.]+)", txt)
    if m:
        return (None, float(m.group(1)) * mul)

    m = re.match(r"^>\s*([0-9\.]+)", txt)
    if m:
        return (float(m.group(1)) * mul, None)

    m = re.match(r"^([0-9\.]+)\s*-\s*([0-9\.]+)", txt)
    if m:
        return (float(m.group(1)) * mul, float(m.group(2)) * mul)

    return None


def load_thresholds_from_excel(path: str) -> Dict[str, Dict[str, Dict[str, Any]]]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Checklist file not found: {path}")

    wb = load_workbook(path, data_only=True)

    categories = ["Valuation", "Profitability", "Balance Sheet", "Growth", "Risk"]
    thresholds: Dict[str, Dict[str, Dict[str, Any]]] = {c: {} for c in categories}

    for cat in categories:
        if cat not in wb.sheetnames:
            continue
        ws = wb[cat]
        for r in range(2, ws.max_row + 1):
            metric = ws.cell(r, 1).value
            if not metric or not isinstance(metric, str):
                continue

            green = ws.cell(r, 2).value
            yellow = ws.cell(r, 3).value
            red = ws.cell(r, 4).value
            marking = ws.cell(r, 5).value
            notes = ws.cell(r, 6).value

            thresholds[cat][metric.strip()] = {
                "Default (All)": {
                    "green_txt": green,
                    "yellow_txt": yellow,
                    "red_txt": red,
                    "marking": marking,
                    "notes": notes,
                }
            }

    if "Sector Adjustments" in wb.sheetnames:
        ws = wb["Sector Adjustments"]
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        col_map = {name: idx + 1 for idx, name in enumerate(header) if isinstance(name, str)}

        sector_cols = [
            "Default (All)",
            "Software/Tech",
            "Industrials",
            "Consumer Staples",
            "Consumer Discretionary",
            "Healthcare/Pharma",
            "Energy/Materials",
            "Financials (Banks)",
            "REITs",
            "Utilities/Telecom",
        ]

        notes_col = col_map.get("Notes (why/when)", ws.max_column)

        for r in range(2, ws.max_row + 1):
            metric = ws.cell(r, col_map.get("Metric", 1)).value
            if not metric or not isinstance(metric, str):
                continue
            metric = metric.strip()

            for sec in sector_cols:
                if sec not in col_map:
                    continue
                val = ws.cell(r, col_map[sec]).value
                if val is None or not isinstance(val, str) or not val.strip():
                    continue

                parts = [p.strip() for p in val.split("/") if p.strip()]
                if len(parts) < 3:
                    continue

                green_txt, yellow_txt, red_txt = parts[0], parts[1], parts[2]
                sec_notes = ws.cell(r, notes_col).value

                for cat in thresholds:
                    if metric in thresholds[cat]:
                        thresholds[cat][metric][sec] = {
                            "green_txt": green_txt,
                            "yellow_txt": yellow_txt,
                            "red_txt": red_txt,
                            "marking": thresholds[cat][metric]["Default (All)"].get("marking"),
                            "notes": sec_notes or thresholds[cat][metric]["Default (All)"].get("notes"),
                        }

    return thresholds


def get_threshold_set(thresholds: Dict[str, Dict[str, Dict[str, Any]]], category: str, metric: str, sector_bucket: str):
    entry = thresholds.get(category, {}).get(metric)
    if not entry:
        return None
    return entry.get(sector_bucket) or entry.get("Default (All)")
