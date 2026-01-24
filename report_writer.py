
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


def _is_aberrant(value: Any, low: Optional[float], high: Optional[float]) -> Tuple[bool, str]:
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

    if low is None and high is None:
        if abs(v) > 1e9:
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
    score_pct: float
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
    for cond, sym in symbols.items():
        ws.cell(start_row, 1).value = cond
        ws.cell(start_row, 2).value = sym
        ws.cell(start_row, 2).fill = FILL_GREEN if sym == "🟢" else (FILL_YELLOW if sym == "🟡" else FILL_RED)
        ws.cell(start_row, 3).value = details.get(cond, (0, ""))[1]
        ws.cell(start_row, 3).alignment = ALIGN_WRAP
        start_row += 1

    ws.cell(top, 4).value = f"Green: {green_count}/{len(symbols)}\nGreen+Yellow: {gy_count}/{len(symbols)}\nWeighted score: {score_pct:.1f}%"
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
        "Fundamental Score %", "Technical Score %",
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
        revpack = reversal_by_ticker[t]
        bucket = m.get("Sector Bucket", "Default (All)")
        metric_notes = m.get("__notes__", {}) if isinstance(m.get("__notes__", {}), dict) else {}

        f_score = revpack.get("fundamental_score", 0.0)
        t_score = revpack.get("technical_score", 0.0)
        counts = revpack.get("counts", {})

        ws_sum.append([
            t, bucket, m.get("NUPL Regime"), m.get("Composite NUPL"),
            f_score, t_score,
            counts.get("fund_green", 0), counts.get("fund_gy", 0),
            counts.get("tech_green", 0), counts.get("tech_gy", 0),
        ])

        ws = wb.create_sheet(t)
        ws["A1"] = f"{t} — Checklist v2 (Sector-adjusted): {bucket}"
        ws.merge_cells("A1:F1")

        ws["A2"] = "Yahoo Sector"
        ws["B2"] = m.get("Yahoo Sector")
        ws["A3"] = "Yahoo Industry"
        ws["B3"] = m.get("Yahoo Industry")
        ws["A4"] = "Price"
        ws["B4"] = m.get("Price")

        row = 6

        for cat_sheet, cat_title in category_maps.items():
            ws[f"A{row}"] = cat_title
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1

            headers = ["Metric", "Value", "Score", "Sector Mode", "Limits used", "Notes"]
            for i, h in enumerate(headers, start=1):
                cell = ws.cell(row, i)
                cell.value = h
                cell.fill = FILL_HDR
                cell.font = FONT_HDR
                cell.alignment = ALIGN_CENTER
            row += 1

            for metric in thresholds.get(cat_sheet, {}).keys():
                raw_val = m.get(metric)
                th = get_threshold_set(thresholds, cat_sheet, metric, bucket)

                base_notes = ""
                limits = ""
                mode = "Default (All)"
                score = "NA"
                fill = FILL_GRAY
                extra_auto_note = ""

                if th:
                    limits = f"{th['green_txt']} | {th['yellow_txt']} | {th['red_txt']}"
                    base_notes = th.get("notes") or ""
                    mode = bucket if bucket in thresholds[cat_sheet][metric] else "Default (All)"

                    lo, hi = _extract_numeric_bounds(th["green_txt"], th["yellow_txt"], th["red_txt"])
                    is_bad, bad_msg = _is_aberrant(raw_val, lo, hi)
                    if is_bad:
                        val = None
                        score = "NA"
                        fill = FILL_GRAY
                        extra_auto_note = bad_msg
                    else:
                        val = raw_val
                        score, fill = score_with_threshold_txt(val, th["green_txt"], th["yellow_txt"], th["red_txt"])
                else:
                    val = raw_val

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
                ws.cell(row, 5).value = limits
                ws.cell(row, 6).value = notes

                ws.cell(row, 3).fill = fill

                for c in range(1, 7):
                    ws.cell(row, c).alignment = ALIGN_WRAP

                row += 1

            row += 1

        fund_sym = revpack.get("fundamental_symbols", {})
        tech_sym = revpack.get("technical_symbols", {})
        fund_det = revpack.get("fundamental", {})
        tech_det = revpack.get("technical", {})

        row = _write_reversal_block(
            ws, row,
            "Fundamental Turnaround Signals (7)",
            fund_sym, fund_det,
            counts.get("fund_green", 0), counts.get("fund_gy", 0),
            f_score
        )

        row = _write_reversal_block(
            ws, row,
            "Technical Confirmation Signals (7)",
            tech_sym, tech_det,
            counts.get("tech_green", 0), counts.get("tech_gy", 0),
            t_score
        )

        autosize(ws)

    autosize(ws_sum)
    ws_sum.freeze_panes = "A2"
    wb.save(out_path)
