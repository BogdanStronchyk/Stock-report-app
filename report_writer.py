from typing import Any, Dict, List, Optional, Tuple
import math

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from config import (
    FILL_HDR, FONT_HDR, ALIGN_CENTER, ALIGN_WRAP,
    FILL_GREEN, FILL_YELLOW, FILL_RED, FILL_GRAY
)
from checklist_loader import get_threshold_set, parse_range_cell
from scoring import score_with_threshold_txt


def autosize(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 60)


def _extract_numeric_bounds(green_txt: Any, yellow_txt: Any, red_txt: Any) -> Tuple[Optional[float], Optional[float]]:
    rules = []
    for txt in (green_txt, yellow_txt, red_txt):
        if txt is None:
            continue
        r = parse_range_cell(str(txt))
        if r:
            rules.append(r)
    if not rules:
        return (None, None)

    lows = [lo for lo, hi in rules if lo is not None]
    highs = [hi for lo, hi in rules if hi is not None]
    return (min(lows) if lows else None, max(highs) if highs else None)


def _limits_text(th: Optional[Dict[str, Any]]) -> str:
    if not th:
        return ""
    g = th.get("green_txt", "")
    y = th.get("yellow_txt", "")
    r = th.get("red_txt", "")
    parts = []
    if g not in (None, ""):
        parts.append(f"G:{g}")
    if y not in (None, ""):
        parts.append(f"Y:{y}")
    if r not in (None, ""):
        parts.append(f"R:{r}")
    return " | ".join(parts)


def _score_symbol(points: int) -> str:
    return "🟢" if points == 2 else ("🟡" if points == 1 else "🔴")


def _normalize_reversal_pack(revpack: Dict[str, Any]) -> Dict[str, Any]:
    """
    Supports BOTH:
      - old format: fund_symbols/fund_details/fund_score_pct/...
      - current reversal.py format:
          fundamental, technical, fundamental_symbols, technical_symbols,
          fundamental_score, technical_score, counts
    Returns a normalized dict with:
      fund_symbols, tech_symbols, fund_details, tech_details,
      fund_score_pct, tech_score_pct, total_score_pct,
      fund_green, fund_green_yellow, tech_green, tech_green_yellow
    """
    if not isinstance(revpack, dict):
        revpack = {}

    fundamental = revpack.get("fundamental") or {}
    technical = revpack.get("technical") or {}

    # Symbols
    fund_symbols = revpack.get("fund_symbols") or revpack.get("fundamental_symbols")
    tech_symbols = revpack.get("tech_symbols") or revpack.get("technical_symbols")

    if not fund_symbols and isinstance(fundamental, dict):
        fund_symbols = {k: _score_symbol(v[0]) for k, v in fundamental.items() if isinstance(v, (tuple, list)) and len(v) >= 1}
    if not tech_symbols and isinstance(technical, dict):
        tech_symbols = {k: _score_symbol(v[0]) for k, v in technical.items() if isinstance(v, (tuple, list)) and len(v) >= 1}

    # Details (condition -> (points, detail_text))
    fund_details = revpack.get("fund_details")
    tech_details = revpack.get("tech_details")
    if not fund_details and isinstance(fundamental, dict):
        fund_details = {k: (v[0], v[1] if len(v) > 1 else "") for k, v in fundamental.items() if isinstance(v, (tuple, list)) and len(v) >= 1}
    if not tech_details and isinstance(technical, dict):
        tech_details = {k: (v[0], v[1] if len(v) > 1 else "") for k, v in technical.items() if isinstance(v, (tuple, list)) and len(v) >= 1}

    # Scores
    fund_score = revpack.get("fund_score_pct")
    tech_score = revpack.get("tech_score_pct")
    if fund_score is None:
        fund_score = revpack.get("fundamental_score")
    if tech_score is None:
        tech_score = revpack.get("technical_score")

    # Total combined (0.6/0.4 default)
    total_score = revpack.get("total_score_pct")
    if total_score is None and fund_score is not None and tech_score is not None:
        try:
            total_score = 0.6 * float(fund_score) + 0.4 * float(tech_score)
        except Exception:
            total_score = None

    # Counts
    counts = revpack.get("counts") or {}
    fund_green = revpack.get("fund_green")
    fund_gy = revpack.get("fund_green_yellow")  # old name
    tech_green = revpack.get("tech_green")
    tech_gy = revpack.get("tech_green_yellow")  # old name

    if fund_green is None:
        fund_green = counts.get("fund_green")
    if fund_gy is None:
        fund_gy = counts.get("fund_gy")
    if tech_green is None:
        tech_green = counts.get("tech_green")
    if tech_gy is None:
        tech_gy = counts.get("tech_gy")

    # Fallback counts if still missing
    if fund_green is None and isinstance(fund_symbols, dict):
        fund_green = sum(1 for v in fund_symbols.values() if v == "🟢")
    if fund_gy is None and isinstance(fund_symbols, dict):
        fund_gy = sum(1 for v in fund_symbols.values() if v in ("🟢", "🟡"))
    if tech_green is None and isinstance(tech_symbols, dict):
        tech_green = sum(1 for v in tech_symbols.values() if v == "🟢")
    if tech_gy is None and isinstance(tech_symbols, dict):
        tech_gy = sum(1 for v in tech_symbols.values() if v in ("🟢", "🟡"))

    return {
        "fund_symbols": fund_symbols or {},
        "tech_symbols": tech_symbols or {},
        "fund_details": fund_details or {},
        "tech_details": tech_details or {},
        "fund_score_pct": fund_score,
        "tech_score_pct": tech_score,
        "total_score_pct": total_score,
        "fund_green": fund_green if fund_green is not None else 0,
        "fund_green_yellow": fund_gy if fund_gy is not None else 0,
        "tech_green": tech_green if tech_green is not None else 0,
        "tech_green_yellow": tech_gy if tech_gy is not None else 0,
    }


def _is_aberrant(metric_name: str, value: Any, low: Optional[float], high: Optional[float]) -> Tuple[bool, str]:
    """
    Detect values that are so extreme they likely indicate a bad scrape/parse.

    IMPORTANT:
    Some metrics are legitimately huge (Market Cap, Avg Daily $ Volume, EV, etc.).
    Those should NOT be nulled out, even if they are far outside checklist bounds.
    Let them score RED instead of becoming N/A.
    """
    try:
        if value is None:
            return (False, "")
        if isinstance(value, str):
            return (False, "")
        v = float(value)
        if math.isnan(v) or math.isinf(v):
            return (True, "No meaningful data to calculate this metric (NaN/Inf).")
    except Exception:
        return (False, "")

    big_metrics = [
        "Market Cap",
        "Enterprise Value",
        "Avg Daily $ Volume",
    ]
    if any(k in (metric_name or "") for k in big_metrics):
        return (False, "")

    if low is None and high is None:
        if abs(v) > 1e15:
            return (True, "No meaningful data to calculate this metric (extreme magnitude).")
        return (False, "")

    ref = max([abs(x) for x in [low, high] if x is not None] + [1.0])
    MULT = 50.0

    if low is not None and high is not None:
        span = max(abs(high - low), ref * 0.05)
        hard_low = low - MULT * span
        hard_high = high + MULT * span
    elif low is None:
        hard_low = -MULT * ref
        hard_high = (high if high is not None else ref) + MULT * ref
    else:
        hard_low = (low if low is not None else -ref) - MULT * ref
        hard_high = MULT * ref

    if v < hard_low or v > hard_high:
        return (True, "No meaningful data to calculate this metric (aberrant outlier vs checklist limits).")
    return (False, "")


def _write_reversal_block(
    ws,
    start_row: int,
    title: str,
    symbols: Dict[str, str],
    details: Dict[str, Tuple[int, str]],
    green_count: int,
    gy_count: int,
    score_pct: Optional[float]
) -> int:
    ws[f"A{start_row}"] = title
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
    start_row += 1

    ws.cell(start_row, 1).value = "Condition"
    ws.cell(start_row, 2).value = "Score"
    ws.cell(start_row, 3).value = "Details"
    ws.cell(start_row, 4).value = "Summary"
    for c in range(1, 5):
        cell = ws.cell(start_row, c)
        cell.fill = FILL_HDR
        cell.font = FONT_HDR
        cell.alignment = ALIGN_CENTER
    start_row += 1

    top = start_row

    # stable ordering
    for cond in list(symbols.keys()):
        sym = symbols.get(cond)
        ws.cell(start_row, 1).value = cond
        ws.cell(start_row, 2).value = sym
        ws.cell(start_row, 2).fill = FILL_GREEN if sym == "🟢" else (FILL_YELLOW if sym == "🟡" else FILL_RED)

        det = details.get(cond, (0, ""))[1]
        ws.cell(start_row, 3).value = det
        ws.cell(start_row, 3).alignment = ALIGN_WRAP
        start_row += 1

    score_txt = "NA" if score_pct is None else f"{float(score_pct):.1f}%"
    total_n = len(symbols) if isinstance(symbols, dict) else 0
    ws.cell(top, 4).value = f"Green: {green_count}/{total_n}\nGreen+Yellow: {gy_count}/{total_n}\nWeighted score: {score_txt}"
    ws.cell(top, 4).alignment = ALIGN_WRAP

    for r in range(top, start_row):
        for c in range(1, 5):
            ws.cell(r, c).alignment = ALIGN_WRAP

    return start_row + 1


def create_report_workbook(
    tickers: List[str],
    thresholds: Dict[str, Dict[str, Dict[str, Any]]],
    metrics_by_ticker: Dict[str, Dict[str, Any]],
    reversal_by_ticker: Dict[str, Dict[str, Any]],
    out_path: str
):
    wb = Workbook()
    wb.remove(wb.active)

    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.append([
        "Ticker", "Sector Bucket", "NUPL Regime", "Composite NUPL",
        "Fundamental Score %", "Technical Score %", "Total Combined Reversal Score %",
        "Fund Green", "Fund G+Y", "Tech Green", "Tech G+Y"
    ])
    for cell in ws_sum[1]:
        cell.fill = FILL_HDR
        cell.font = FONT_HDR
        cell.alignment = ALIGN_CENTER

    category_maps = {
        "Valuation": "Valuation",
        "Profitability": "Profitability",
        "Balance Sheet": "Balance Sheet",
        "Growth": "Growth",
        "Risk": "Risk",
    }

    for t in tickers:
        m = metrics_by_ticker[t]
        revpack_raw = reversal_by_ticker.get(t, {})
        revpack = _normalize_reversal_pack(revpack_raw)

        bucket = m.get("Sector Bucket", "Default (All)")
        metric_notes = m.get("__notes__", {}) or {}

        ws = wb.create_sheet(t)

        ws["A1"] = f"{t} — Checklist v2 (Sector-adjusted): {bucket}"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)

        ws["A2"] = "Yahoo Sector"
        ws["B2"] = m.get("Yahoo Sector")

        ws["A3"] = "Yahoo Industry"
        ws["B3"] = m.get("Yahoo Industry")

        ws["A4"] = "Price"
        ws["B4"] = m.get("Price")

        ws["A5"] = "Total Combined Reversal Score %"
        ws["B5"] = revpack.get("total_score_pct")

        ws["A6"] = "Data Notes"
        ws["B6"] = (metric_notes.get("FMP") or "")[:4000]

        ws_sum.append([
            t,
            bucket,
            m.get("NUPL Regime"),
            m.get("Composite NUPL"),
            revpack.get("fund_score_pct"),
            revpack.get("tech_score_pct"),
            revpack.get("total_score_pct"),
            revpack.get("fund_green"),
            revpack.get("fund_green_yellow"),
            revpack.get("tech_green"),
            revpack.get("tech_green_yellow"),
        ])

        row = 8

        for cat_sheet, cat_title in category_maps.items():
            ws.cell(row, 1).value = cat_title
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1

            ws.append(["Metric", "Value", "Score", "Sector Mode", "Limits used", "Notes"])
            for c in range(1, 7):
                cell = ws.cell(row, c)
                cell.fill = FILL_HDR
                cell.font = FONT_HDR
                cell.alignment = ALIGN_CENTER
            row += 1

            for metric, thset in thresholds.get(cat_sheet, {}).items():
                raw_val = m.get(metric)

                # correct signature
                th = get_threshold_set(thresholds, cat_sheet, metric, bucket)

                mode = bucket if (isinstance(thset, dict) and bucket in thset) else "Default (All)"
                base_notes = (th.get("notes") if th else "") or ""

                extra_auto_note = ""
                if th:
                    lo, hi = _extract_numeric_bounds(th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
                    is_bad, bad_msg = _is_aberrant(metric, raw_val, lo, hi)
                    if is_bad:
                        val = None
                        score = "NA"
                        fill = FILL_GRAY
                        extra_auto_note = bad_msg
                    else:
                        val = raw_val
                        score, fill = score_with_threshold_txt(val, th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
                else:
                    val = raw_val
                    score = "NA"
                    fill = FILL_GRAY

                extra_note = metric_notes.get(metric, "")

                notes = base_notes
                if extra_note:
                    notes = (notes + "\n" if notes else "") + "⚠ " + extra_note
                if extra_auto_note:
                    notes = (notes + "\n" if notes else "") + "⚠ " + extra_auto_note

                ws.cell(row, 1).value = metric
                ws.cell(row, 2).value = val
                ws.cell(row, 3).value = score
                ws.cell(row, 4).value = mode
                ws.cell(row, 5).value = _limits_text(th)
                ws.cell(row, 6).value = notes

                ws.cell(row, 3).fill = fill

                for c in range(1, 7):
                    ws.cell(row, c).alignment = ALIGN_WRAP

                row += 1

            row += 1

        # Reversal blocks (now using normalized pack)
        row = _write_reversal_block(
            ws,
            row,
            "Fundamental Turnaround Signals (7)",
            revpack.get("fund_symbols", {}),
            revpack.get("fund_details", {}),
            revpack.get("fund_green", 0),
            revpack.get("fund_green_yellow", 0),
            revpack.get("fund_score_pct"),
        )

        row = _write_reversal_block(
            ws,
            row,
            "Technical Confirmation Signals (7)",
            revpack.get("tech_symbols", {}),
            revpack.get("tech_details", {}),
            revpack.get("tech_green", 0),
            revpack.get("tech_green_yellow", 0),
            revpack.get("tech_score_pct"),
        )

        autosize(ws)

    autosize(ws_sum)
    wb.save(out_path)
