import os
import re
from typing import Any, Dict, Optional, Tuple
from openpyxl import load_workbook

_NUM = r"[+-]?\d*\.?\d+"  # signed float

def _norm_metric(s: str) -> str:
    s = s or ""
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def _strip_parens(s: str) -> str:
    # remove parenthetical fragments to improve matching: "FCF Yield (TTM ...)" -> "FCF Yield"
    return re.sub(r"\s*\([^)]*\)", "", s).strip()

def _is_heading_row(metric: str) -> bool:
    m = _norm_metric(metric)
    if not m:
        return True
    # common non-metric helper rows in the Sector Adjustments sheet
    return any(
        x in m
        for x in [
            "how to use",
            "step",
            "sector id",
            "adjustments",
        ]
    )

def _metric_matches(adj_metric: str, threshold_metric: str) -> bool:
    a = _norm_metric(adj_metric)
    t = _norm_metric(threshold_metric)
    if not a or not t:
        return False
    if a == t:
        return True
    # try without parentheses on either side
    a2 = _norm_metric(_strip_parens(adj_metric))
    t2 = _norm_metric(_strip_parens(threshold_metric))
    if a2 and (a2 == t2):
        return True
    # containment / prefix matching (handles "FCF Yield" vs "FCF Yield (TTM ...)")
    if a2 and a2 in t:
        return True
    if a and a in t:
        return True
    if t2 and t2 in a:
        return True
    return False


def parse_range_cell(s: Any) -> Optional[Tuple[Optional[float], Optional[float]]]:
    """
    Robustly parse threshold text into (lo, hi).

    Handles messy real-world checklist strings like:
      "< 15"
      "15–25"
      "> 25 or negative"
      "> 6% | otherwise ..."
      "< 3 days"
      "3–6 days"
      "<= 10"
      ">= -20%"
      "$10B" / "$250M" / "$1.2T"
      "-35 to -50"
      "-10% to -20%"

    Returns (None, hi) for "< hi", (lo, None) for "> lo", (lo, hi) for ranges.
    """
    if s is None:
        return None
    txt = str(s).strip()
    if not txt:
        return None

    # Normalize dashes and whitespace
    txt = txt.replace("–", "-").replace("—", "-")
    txt = re.sub(r"\s+", " ", txt).strip()

    # Ignore purely qualitative guidance (keep this list tight)
    low_txt = txt.lower()
    if any(w in low_txt for w in ["expanding", "stable", "contracting"]):
        return None

    # Remove separators that interfere with float parsing
    txt = txt.replace(",", "").replace("_", "")

    # Currency multipliers (only if $ is present)
    mul = 1.0
    if "$" in txt:
        # detect suffix after a number, e.g. $10B, $250M, $1.2T
        m = re.search(rf"({_NUM})\s*([KMBT])\b", txt, flags=re.IGNORECASE)
        if m:
            suf = m.group(2).upper()
            mul = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}.get(suf, 1.0)
            # strip the suffix letter so the remaining patterns still match
            txt = re.sub(rf"({_NUM})\s*[KMBT]\b", r"\1", txt, flags=re.IGNORECASE)

        txt = txt.replace("$", "").strip()

    # Remove % (your metrics are already in percent units, not fractions)
    txt = txt.replace("%", "").strip()

    # Normalize "to" to "-"
    txt = re.sub(r"\s+to\s+", " - ", txt, flags=re.IGNORECASE)

    # 1) Comparator rule anywhere in the string (supports extra trailing words)
    m = re.search(rf"(<=|>=|<|>)\s*({_NUM})", txt)
    if m:
        op = m.group(1)
        num = float(m.group(2)) * mul
        if op in ("<", "<="):
            return (None, num)
        else:
            return (num, None)

    # 2) Range rule anywhere in the string (supports extra trailing words like "days")
    m = re.search(rf"({_NUM})\s*-\s*({_NUM})", txt)
    if m:
        a = float(m.group(1)) * mul
        b = float(m.group(2)) * mul
        lo, hi = (a, b) if a <= b else (b, a)
        return (lo, hi)

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

    # Sector adjustments sheet (optional)
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
            if _is_heading_row(metric):
                continue

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
                    for tm in list(thresholds[cat].keys()):
                        if _metric_matches(metric, tm):
                            thresholds[cat][tm][sec] = {

                            "green_txt": green_txt,
                            "yellow_txt": yellow_txt,
                            "red_txt": red_txt,
                            "marking": thresholds[cat][tm]["Default (All)"].get("marking"),
                            "notes": sec_notes,
                        }

    return thresholds


def get_threshold_set(
    thresholds: Dict[str, Dict[str, Dict[str, Any]]],
    category: str,
    metric: str,
    sector_bucket: str
) -> Dict[str, Any]:
    if category not in thresholds or metric not in thresholds[category]:
        return {}
    by_sector = thresholds[category][metric]
    if sector_bucket in by_sector:
        return by_sector[sector_bucket]
    return by_sector.get("Default (All)", {})
