
from typing import Any, Dict, List, Optional, Tuple
import math

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from config import (
    FILL_HDR, FONT_HDR, ALIGN_CENTER, ALIGN_WRAP,
    FILL_GREEN, FILL_YELLOW, FILL_RED, FILL_GRAY
)
from checklist_loader import get_threshold_set, parse_range_cell
from scoring import (
    score_with_threshold_txt,
    CATEGORY_WEIGHTS,
    compute_category_score_and_coverage,
    adjusted_from_raw_and_coverage,
    rating_to_points,
)


def autosize(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 70)


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
    if not isinstance(revpack, dict):
        revpack = {}

    fundamental = revpack.get("fundamental") or {}
    technical = revpack.get("technical") or {}

    fund_symbols = revpack.get("fund_symbols") or revpack.get("fundamental_symbols")
    tech_symbols = revpack.get("tech_symbols") or revpack.get("technical_symbols")

    if not fund_symbols and isinstance(fundamental, dict):
        fund_symbols = {k: _score_symbol(v[0]) for k, v in fundamental.items() if isinstance(v, (tuple, list)) and len(v) >= 1}
    if not tech_symbols and isinstance(technical, dict):
        tech_symbols = {k: _score_symbol(v[0]) for k, v in technical.items() if isinstance(v, (tuple, list)) and len(v) >= 1}

    fund_details = revpack.get("fund_details")
    tech_details = revpack.get("tech_details")
    if not fund_details and isinstance(fundamental, dict):
        fund_details = {k: (v[0], v[1] if len(v) > 1 else "") for k, v in fundamental.items() if isinstance(v, (tuple, list)) and len(v) >= 1}
    if not tech_details and isinstance(technical, dict):
        tech_details = {k: (v[0], v[1] if len(v) > 1 else "") for k, v in technical.items() if isinstance(v, (tuple, list)) and len(v) >= 1}

    fund_score = revpack.get("fund_score_pct")
    tech_score = revpack.get("tech_score_pct")
    if fund_score is None:
        fund_score = revpack.get("fundamental_score")
    if tech_score is None:
        tech_score = revpack.get("technical_score")

    total_score = revpack.get("total_score_pct")
    if total_score is None and fund_score is not None and tech_score is not None:
        try:
            total_score = 0.6 * float(fund_score) + 0.4 * float(tech_score)
        except Exception:
            total_score = None

    counts = revpack.get("counts") or {}
    fund_green = revpack.get("fund_green")
    fund_gy = revpack.get("fund_green_yellow")
    tech_green = revpack.get("tech_green")
    tech_gy = revpack.get("tech_green_yellow")

    if fund_green is None:
        fund_green = counts.get("fund_green")
    if fund_gy is None:
        fund_gy = counts.get("fund_gy")
    if tech_green is None:
        tech_green = counts.get("tech_green")
    if tech_gy is None:
        tech_gy = counts.get("tech_gy")

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

    big_metrics = ["Market Cap", "Enterprise Value", "Avg Daily $ Volume"]
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


def _metric_weight(category: str, metric: str) -> float:
    m = metric or ""
    cat = category or ""
    w = 1.0

    if cat == "Valuation":
        if "EV/FCF" in m or "FCF Yield" in m or "EV/EBIT" in m:
            w = 1.6
        elif "EV/EBITDA" in m or "P/E" in m:
            w = 1.3
        elif "P/S" in m or "EV/Gross Profit" in m:
            w = 1.1

    elif cat == "Profitability":
        if "ROIC" in m:
            w = 1.8
        elif "Operating Margin" in m or "FCF Margin" in m:
            w = 1.4
        elif "Gross Margin" in m or "Net Margin" in m:
            w = 1.2
        elif "CFO / Net Income" in m or "ROE" in m:
            w = 1.0

    elif cat == "Balance Sheet":
        if "Net Debt / EBITDA" in m or "Interest Coverage" in m:
            w = 1.7
        elif "Net Debt / FCF" in m or "FCF / Interest" in m:
            w = 1.4
        elif "Current Ratio" in m or "Quick Ratio" in m:
            w = 1.0
        elif "Cash / Total Assets" in m:
            w = 0.8

    elif cat == "Growth":
        if "FCF per Share CAGR" in m or "Revenue per Share CAGR" in m:
            w = 1.4
        elif "Revenue CAGR" in m:
            w = 1.2
        else:
            w = 0.9

    elif cat == "Risk":
        if "Max Drawdown" in m or "Realized Volatility" in m:
            w = 1.4
        elif "Worst Weekly Return" in m:
            w = 1.3
        elif "Avg Daily $ Volume" in m:
            w = 1.2
        elif "Beta" in m:
            w = 1.0
        elif "Short Interest" in m or "Days to Cover" in m:
            w = 0.9
        elif "Market Cap" in m:
            w = 0.7

    return float(w)


def _apply_category_caps(category: str, raw_score: Optional[float], ratings_by_metric: Dict[str, str]) -> Optional[float]:
    if raw_score is None:
        return None

    if category == "Balance Sheet":
        red_flag = any(("Net Debt / EBITDA" in k or "Interest Coverage" in k) and (v or "").upper() == "RED"
                       for k, v in ratings_by_metric.items())
        if red_flag:
            return float(min(raw_score, 60.0))

    if category == "Risk":
        illiquid_red = any("Avg Daily $ Volume" in k and (v or "").upper() == "RED" for k, v in ratings_by_metric.items())
        if illiquid_red:
            return float(min(raw_score, 65.0))

    return float(raw_score)


def _weighted_blend(values_by_category: Dict[str, Optional[float]]) -> Optional[float]:
    wsum = 0.0
    acc = 0.0
    for cat, w in CATEGORY_WEIGHTS.items():
        v = values_by_category.get(cat)
        if v is None:
            continue
        wsum += float(w)
        acc += float(w) * float(v)
    return None if wsum == 0 else (acc / wsum)


def _write_reversal_block(ws, start_row: int, title: str, symbols: Dict[str, str], details: Dict[str, Tuple[int, str]],
                          green_count: int, gy_count: int, score_pct: Optional[float]) -> int:
    ws[f"A{start_row}"] = title
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=9)
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
    for cond in list(symbols.keys()):
        sym = symbols.get(cond)
        ws.cell(start_row, 1).value = cond
        ws.cell(start_row, 2).value = sym
        ws.cell(start_row, 2).fill = FILL_GREEN if sym == "🟢" else (FILL_YELLOW if sym == "🟡" else FILL_RED)
        ws.cell(start_row, 3).value = details.get(cond, (0, ""))[1]
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


def _add_cheat_sheet(wb: Workbook):
    ws = wb.create_sheet("Cheat Sheet", 1)
    ws["A1"] = "How to interpret Fundamental Checklist scoring"
    ws.merge_cells("A1:I1")
    ws["A2"] = "Ratings → Points"
    ws["A3"] = "GREEN=2, YELLOW=1, RED=0. NA is excluded from scoring denominators but reduces coverage."
    ws["A5"] = "Coverage %"
    ws["A6"] = "Coverage is computed per category and overall, as % of the category's metric weights that were scorable (not NA)."
    ws["A8"] = "Raw Fundamental Checklist Score %"
    ws["A9"] = "Raw score uses ONLY scorable metrics. High raw score with low coverage means 'good on limited data'."
    ws["A11"] = "Adjusted Fundamental Checklist Score %"
    ws["A12"] = "Adjusted score discounts raw scores by coverage (category-weighted). Use this as a conservative headline number."
    ws["A14"] = "Suggested bands (Adjusted Score)"
    ws["A15"] = "0–20 weak/unknown | 20–40 fragile | 40–60 acceptable | 60–75 strong | 75–100 exceptional"
    for r in range(1, 16):
        ws[f"A{r}"].alignment = ALIGN_WRAP
    ws.column_dimensions["A"].width = 95


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
        "Fund Green", "Fund G+Y", "Tech Green", "Tech G+Y",
        "Valuation %", "Valuation Cov %",
        "Profitability %", "Profitability Cov %",
        "Balance Sheet %", "Balance Sheet Cov %",
        "Growth %", "Growth Cov %",
        "Risk %", "Risk Cov %",
        "Fund Checklist % (Raw)", "Data Coverage % (Total)",
        "Fund Checklist % (Adjusted)",
    ])
    for cell in ws_sum[1]:
        cell.fill = FILL_HDR
        cell.font = FONT_HDR
        cell.alignment = ALIGN_CENTER

    category_maps = {"Valuation": "Valuation", "Profitability": "Profitability", "Balance Sheet": "Balance Sheet", "Growth": "Growth", "Risk": "Risk"}
    _add_cheat_sheet(wb)

    for t in tickers:
        m = metrics_by_ticker[t]
        revpack = _normalize_reversal_pack(reversal_by_ticker.get(t, {}))

        bucket = m.get("Sector Bucket", "Default (All)")
        metric_notes = m.get("__notes__", {}) or {}

        ws = wb.create_sheet(t)

        ws["A1"] = f"{t} — Checklist v2 (Sector-adjusted): {bucket}"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)

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

        category_ratings: Dict[str, Dict[str, str]] = {c: {} for c in category_maps}
        category_weights: Dict[str, Dict[str, float]] = {c: {} for c in category_maps}

        row = 8

        for cat_sheet, cat_title in category_maps.items():
            ws.cell(row, 1).value = cat_title
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
            row += 1

            ws.append(["Metric", "Value", "Rating", "Points", "Weight", "Weighted Pts", "Sector Mode", "Limits used", "Notes"])
            for c in range(1, 10):
                cell = ws.cell(row, c)
                cell.fill = FILL_HDR
                cell.font = FONT_HDR
                cell.alignment = ALIGN_CENTER
            row += 1

            for metric, thset in thresholds.get(cat_sheet, {}).items():
                raw_val = m.get(metric)
                th = get_threshold_set(thresholds, cat_sheet, metric, bucket)
                mode = bucket if (isinstance(thset, dict) and bucket in thset) else "Default (All)"
                base_notes = (th.get("notes") if th else "") or ""

                extra_auto_note = ""
                if th:
                    lo, hi = _extract_numeric_bounds(th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
                    is_bad, bad_msg = _is_aberrant(metric, raw_val, lo, hi)
                    if is_bad:
                        val = None
                        rating = "NA"
                        fill = FILL_GRAY
                        extra_auto_note = bad_msg
                    else:
                        val = raw_val
                        rating, fill = score_with_threshold_txt(val, th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
                else:
                    val = raw_val
                    rating = "NA"
                    fill = FILL_GRAY

                extra_note = metric_notes.get(metric, "")
                notes = base_notes
                if extra_note:
                    notes = (notes + "\n" if notes else "") + "⚠ " + extra_note
                if extra_auto_note:
                    notes = (notes + "\n" if notes else "") + "⚠ " + extra_auto_note

                w = _metric_weight(cat_title, metric)
                pts = rating_to_points(rating)
                wpts = (pts * w) if rating != "NA" else None

                category_ratings[cat_title][metric] = rating
                category_weights[cat_title][metric] = w

                ws.cell(row, 1).value = metric
                ws.cell(row, 2).value = val
                ws.cell(row, 3).value = rating
                ws.cell(row, 4).value = None if rating == "NA" else pts
                ws.cell(row, 5).value = w
                ws.cell(row, 6).value = wpts
                ws.cell(row, 7).value = mode
                ws.cell(row, 8).value = _limits_text(th)
                ws.cell(row, 9).value = notes

                ws.cell(row, 3).fill = fill

                for c in range(1, 10):
                    ws.cell(row, c).alignment = ALIGN_WRAP

                row += 1

            row += 1

        cat_scores_raw: Dict[str, Optional[float]] = {}
        cat_coverages: Dict[str, float] = {}
        cat_scores_adj: Dict[str, Optional[float]] = {}

        for cat in category_maps.values():
            raw, cov = compute_category_score_and_coverage(category_ratings[cat], category_weights[cat])
            raw = _apply_category_caps(cat, raw, category_ratings[cat])
            cat_scores_raw[cat] = raw
            cat_coverages[cat] = cov
            cat_scores_adj[cat] = adjusted_from_raw_and_coverage(raw, cov)

        fund_checklist_raw = _weighted_blend(cat_scores_raw)
        total_coverage = _weighted_blend({k: float(v) for k, v in cat_coverages.items()})
        fund_checklist_adj = _weighted_blend(cat_scores_adj)

        ws["D2"] = "Valuation %"; ws["E2"] = cat_scores_raw["Valuation"]; ws["F2"] = "Cov %"; ws["G2"] = cat_coverages["Valuation"]
        ws["D3"] = "Profitability %"; ws["E3"] = cat_scores_raw["Profitability"]; ws["F3"] = "Cov %"; ws["G3"] = cat_coverages["Profitability"]
        ws["D4"] = "Balance Sheet %"; ws["E4"] = cat_scores_raw["Balance Sheet"]; ws["F4"] = "Cov %"; ws["G4"] = cat_coverages["Balance Sheet"]
        ws["D5"] = "Growth %"; ws["E5"] = cat_scores_raw["Growth"]; ws["F5"] = "Cov %"; ws["G5"] = cat_coverages["Growth"]
        ws["D6"] = "Risk %"; ws["E6"] = cat_scores_raw["Risk"]; ws["F6"] = "Cov %"; ws["G6"] = cat_coverages["Risk"]
        ws["D7"] = "Fund Checklist % (Raw)"; ws["E7"] = fund_checklist_raw
        ws["F7"] = "Total Coverage %"; ws["G7"] = total_coverage
        ws["H7"] = "Adjusted %"; ws["I7"] = fund_checklist_adj

        for r in range(2, 8):
            for c in ["D", "F", "H"]:
                ws[f"{c}{r}"].font = FONT_HDR
                ws[f"{c}{r}"].fill = FILL_HDR
                ws[f"{c}{r}"].alignment = ALIGN_CENTER
            for c in ["E", "G", "I"]:
                ws[f"{c}{r}"].alignment = ALIGN_CENTER

        ws_sum.append([
            t, bucket, m.get("NUPL Regime"), m.get("Composite NUPL"),
            revpack.get("fund_score_pct"), revpack.get("tech_score_pct"), revpack.get("total_score_pct"),
            revpack.get("fund_green"), revpack.get("fund_green_yellow"),
            revpack.get("tech_green"), revpack.get("tech_green_yellow"),
            cat_scores_raw["Valuation"], cat_coverages["Valuation"],
            cat_scores_raw["Profitability"], cat_coverages["Profitability"],
            cat_scores_raw["Balance Sheet"], cat_coverages["Balance Sheet"],
            cat_scores_raw["Growth"], cat_coverages["Growth"],
            cat_scores_raw["Risk"], cat_coverages["Risk"],
            fund_checklist_raw, total_coverage, fund_checklist_adj,
        ])

        row = _write_reversal_block(ws, row, "Fundamental Turnaround Signals (7)",
                                   revpack.get("fund_symbols", {}), revpack.get("fund_details", {}),
                                   revpack.get("fund_green", 0), revpack.get("fund_green_yellow", 0),
                                   revpack.get("fund_score_pct"))

        row = _write_reversal_block(ws, row, "Technical Confirmation Signals (7)",
                                   revpack.get("tech_symbols", {}), revpack.get("tech_details", {}),
                                   revpack.get("tech_green", 0), revpack.get("tech_green_yellow", 0),
                                   revpack.get("tech_score_pct"))

        autosize(ws)

    autosize(ws_sum)
    wb.save(out_path)
