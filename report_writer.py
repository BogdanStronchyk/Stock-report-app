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
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 58)


def _extract_numeric_bounds(green_txt: Any, yellow_txt: Any, red_txt: Any) -> Tuple[Optional[float], Optional[float]]:
    """Try to infer plausible numeric bounds from checklist text ranges."""
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

    lo = min(lows) if lows else None
    hi = max(highs) if highs else None
    return (lo, hi)


def _is_aberrant(value: Any, low: Optional[float], high: Optional[float]) -> Tuple[bool, str]:
    """Detect extreme outliers that almost certainly indicate a data issue.

    We use soft bounds derived from checklist limits and allow a generous multiple,
    because real-world metrics can be noisy.
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

    # If we have no bounds, don't attempt to label as aberrant
    if low is None and high is None:
        # generic sanity checks only
        if abs(v) > 1e9:
            return (True, "No meaningful data to calculate this metric (extreme magnitude).")
        return (False, "")

    # Reference magnitude for scaling
    ref_candidates = []
    if low is not None:
        ref_candidates.append(abs(low))
    if high is not None:
        ref_candidates.append(abs(high))
    ref = max(ref_candidates + [1.0])

    # Allow huge but finite multiples of the expected range
    MULT = 50.0  # generous buffer; catches -3000 vs -10 type issues
    # If we have both bounds, also use range-based buffer
    if low is not None and high is not None:
        span = abs(high - low)
        span = max(span, ref * 0.05)  # avoid zero span
        hard_low = low - MULT * span
        hard_high = high + MULT * span
    else:
        # one-sided: create a wide implied corridor around the known bound
        if low is None:
            hard_low = -MULT * ref
            hard_high = (high if high is not None else ref) + MULT * ref
        else:
            hard_low = (low if low is not None else -ref) - MULT * ref
            hard_high = MULT * ref

    if v < hard_low or v > hard_high:
        msg = "No meaningful data to calculate this metric (aberrant outlier vs checklist limits)."
        return (True, msg)

    return (False, "")


def create_report_workbook(
    tickers: List[str],
    thresholds: Dict[str, Dict[str, Dict[str, Any]]],
    metrics_by_ticker: Dict[str, Dict[str, Any]],
    reversal_by_ticker: Dict[str, Dict[str, str]],
    out_path: str
):
    wb = Workbook()
    wb.remove(wb.active)

    # Summary sheet
    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.append(["Ticker", "Sector Bucket", "NUPL Regime", "Composite NUPL", "Reversal (Green)", "Reversal (G+Y)"])
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
        rev = reversal_by_ticker[t]
        bucket = m.get("Sector Bucket", "Default (All)")
        metric_notes = m.get("__notes__", {}) if isinstance(m.get("__notes__", {}), dict) else {}

        rev_green = sum(1 for v in rev.values() if v == "🟢")
        rev_gy = sum(1 for v in rev.values() if v in ("🟢", "🟡"))

        ws_sum.append([t, bucket, m.get("NUPL Regime"), m.get("Composite NUPL"), rev_green, rev_gy])

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

        # Category blocks
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

                # Base notes from checklist file
                base_notes = ""
                limits = ""
                mode = "Default (All)"
                score = "NA"
                fill = FILL_GRAY

                if th:
                    limits = f"{th['green_txt']} | {th['yellow_txt']} | {th['red_txt']}"
                    base_notes = th.get("notes") or ""
                    mode = bucket if bucket in thresholds[cat_sheet][metric] else "Default (All)"

                    # --- Outlier sanitation (NEW) ---
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
                        extra_auto_note = ""
                else:
                    val = raw_val
                    extra_auto_note = ""

                # Extra computed note from metrics.py (e.g., not meaningful ratios)
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

        # Reversal block
        ws[f"A{row}"] = "Trend Reversal Checklist (7)"
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

        ws.cell(row, 1).value = "Condition"
        ws.cell(row, 2).value = "Score"
        ws.cell(row, 3).value = "Counts"
        for c in range(1, 4):
            cell = ws.cell(row, c)
            cell.fill = FILL_HDR
            cell.font = FONT_HDR
            cell.alignment = ALIGN_CENTER
        row += 1

        top = row
        for cond, sym in rev.items():
            ws.cell(row, 1).value = cond
            ws.cell(row, 2).value = sym
            ws.cell(row, 2).fill = FILL_GREEN if sym == "🟢" else (FILL_YELLOW if sym == "🟡" else FILL_RED)
            row += 1

        ws.cell(top, 3).value = f"Green: {rev_green}/7\nGreen+Yellow: {rev_gy}/7"
        ws.cell(top, 3).alignment = ALIGN_WRAP

        autosize(ws)

    autosize(ws_sum)
    ws_sum.freeze_panes = "A2"
    wb.save(out_path)
