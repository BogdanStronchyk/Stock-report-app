from typing import Any, Dict, List, Optional, Tuple
import math
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from config import FILL_HDR, FONT_HDR, ALIGN_CENTER, ALIGN_WRAP, FILL_GREEN, FILL_YELLOW, FILL_RED, FILL_GRAY
from checklist_loader import get_threshold_set, parse_range_cell
from scoring import score_with_threshold_txt, compute_category_score_and_coverage, adjusted_from_raw_and_coverage

WHITELIST = {
    "P/E (TTM, positive EPS)", "EV/EBIT", "FCF Yield (TTM FCF / Market Cap)",
    "Gross Margin %", "Operating Margin %", "ROIC % (standardized)",
    "Net Debt / EBITDA", "Interest Coverage (EBIT / Interest)",
    "Revenue per Share CAGR (5Y)", "FCF per Share CAGR (5Y)",
    "Market Cap", "Max Drawdown (3–5Y)"
}


def _apply_metric_value_format(cell, metric: str, val: Any):
    if val is None or not isinstance(val, (int, float)): return
    if any(x in metric.lower() for x in ["%", "yield", "margin", "drawdown", "roic", "cagr"]):
        cell.number_format = '0.00"%"'
    else:
        cell.number_format = "0.00"


def band_fill(score: Optional[float]):
    if score is None: return FILL_GRAY
    return FILL_GREEN if score >= 60.0 else (FILL_YELLOW if score >= 40.0 else FILL_RED)


def reversal_fill(score: Optional[float]):
    if score is None: return FILL_GRAY
    return FILL_GREEN if score >= 70.0 else (FILL_YELLOW if score >= 50.0 else FILL_RED)


def _limits_text(th: Optional[Dict[str, Any]]) -> str:
    if not th: return ""
    return " | ".join(
        [f"{p}:{th.get(k)}" for k, p in [("green_txt", "G"), ("yellow_txt", "Y"), ("red_txt", "R")] if th.get(k)])


def _metric_weight(metric: str) -> float:
    if metric in ["EV/EBIT", "FCF Yield (TTM FCF / Market Cap)", "ROIC % (standardized)"]: return 2.0
    return 1.5 if metric == "Net Debt / EBITDA" else 1.0


def final_recommendation_banner(cat_scores: Dict[str, Optional[float]], reversal_total: Optional[float],
                                threshold: float) -> Tuple[str, Any]:
    scores = {k: (v if v is not None else 0.0) for k, v in cat_scores.items()}
    rev = reversal_total if reversal_total is not None else 0.0

    # Fundamental pass: All 5 categories >= threshold
    all_fund_pass = all(s >= threshold for s in scores.values())

    if all_fund_pass and rev >= threshold:
        return (f"✅ STRONG BUY ({int(threshold)}% Threshold)", FILL_GREEN)

    # Watch logic: Fundamental quality is high, technicals are low
    if all_fund_pass and rev < threshold:
        return (f"⚠ WATCH ({int(threshold)}% Fund. Pass)", FILL_YELLOW)

    if any(s < 30 for s in scores.values()):
        return ("❌ AVOID — Significant Fundamental Risks", FILL_RED)

    return ("⚠ HOLD / NEUTRAL", FILL_YELLOW)


def _normalize_reversal_pack(revpack: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(revpack, dict): revpack = {}
    fs = revpack.get("fund_score_pct") or revpack.get("fundamental_score")
    ts = revpack.get("tech_score_pct") or revpack.get("technical_score")
    total = revpack.get("total_score_pct")
    if total is None and fs is not None and ts is not None:
        total = 0.6 * float(fs) + 0.4 * float(ts)
    return {
        "fund_symbols": revpack.get("fund_symbols", {}) or revpack.get("fundamental_symbols", {}),
        "tech_symbols": revpack.get("tech_symbols", {}) or revpack.get("technical_symbols", {}),
        "fund_details": revpack.get("fund_details", {}) or revpack.get("fundamental", {}),
        "tech_details": revpack.get("tech_details", {}) or revpack.get("technical", {}),
        "fund_score_pct": fs, "tech_score_pct": ts, "total_score_pct": total,
    }


def _write_reversal_block(ws, start_row: int, title: str, symbols: Dict[str, str], details: Dict[str, Tuple[int, str]],
                          score: Optional[float]) -> int:
    ws[f"A{start_row}"] = title;
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6);
    start_row += 1
    for i, h in enumerate(["Condition", "Score", "Details"]):
        c = ws.cell(start_row, i + 1, h);
        c.fill = FILL_HDR;
        c.font = FONT_HDR
    ws.cell(start_row, 4).value = f"Score: {float(score):.1f}%" if score is not None else "Score: NA";
    start_row += 1
    for cond, sym in symbols.items():
        ws.cell(start_row, 1, cond);
        ws.cell(start_row, 2, sym)
        det = details.get(cond);
        ws.cell(start_row, 3, det[1] if isinstance(det, (list, tuple)) else str(det));
        start_row += 1
    return start_row + 1


def autosize_columns(ws):
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    val_len = len(str(cell.value))
                    if val_len > max_length: max_length = val_len
            except:
                pass
        ws.column_dimensions[column].width = max(min(max_length + 3, 70), 10)


def _add_cheat_sheet(wb: Workbook, threshold: float):
    ws = wb.create_sheet("Cheat Sheet", 1)
    ws["A1"] = "REPORT GUIDE & METRIC EXPLANATIONS";
    ws["A1"].font = Font(size=14, bold=True)

    # Section 1: Color Bands
    ws["A3"] = "1. SCORING BANDS (Colors)";
    ws["A3"].font = Font(bold=True)
    ws["A4"] = "GREEN";
    ws["B4"] = "High Quality / Pass";
    ws["C4"] = "Meets or exceeds strict criteria";
    ws["A4"].fill = FILL_GREEN
    ws["A5"] = "YELLOW";
    ws["B5"] = "Average / Warning";
    ws["C5"] = "Acceptable but not leading";
    ws["A5"].fill = FILL_YELLOW
    ws["A6"] = "RED";
    ws["B6"] = "Risky / Fail";
    ws["C6"] = "Potential fundamental or technical weakness";
    ws["A6"].fill = FILL_RED

    # Section 2: Logic
    ws["A8"] = "2. RECOMMENDATION RULES";
    ws["A8"].font = Font(bold=True)
    ws["A9"] = "✅ STRONG BUY";
    ws["B9"] = f"All fundamental categories AND Reversal score ≥ {int(threshold)}%";
    ws["A9"].fill = FILL_GREEN
    ws["A10"] = "⚠ WATCH";
    ws["B10"] = f"All fundamental categories ≥ {int(threshold)}% but Reversal score is low";
    ws["A10"].fill = FILL_YELLOW
    ws["A11"] = "❌ AVOID";
    ws["B11"] = "Any core fundamental category score < 30%";
    ws["A11"].fill = FILL_RED

    # Section 3: Metric Definitions
    ws["A13"] = "3. METRIC DEFINITIONS";
    ws["A13"].font = Font(bold=True)
    metrics_list = [
        ("P/E (TTM)", "Price-to-Earnings: Standard valuation relative to earnings."),
        ("EV/EBIT", "Enterprise Value / EBIT: Valuation including debt/cash levels."),
        ("FCF Yield", "Free Cash Flow / Market Cap: Real cash return of the business."),
        ("ROIC %", "Return on Invested Capital: Efficiency of capital usage."),
        ("Net Debt/EBITDA", "Leverage: Years of profit needed to pay off all debt."),
        ("Revenue CAGR", "5-Year growth rate of sales on a per-share basis."),
        ("Reversal Score", "Combo of fundamental improvement (Margins/CF) and Technical Trend.")
    ]
    for i, (m, d) in enumerate(metrics_list):
        ws.cell(14 + i, 1, m).font = Font(bold=True)
        ws.cell(14 + i, 2, d)

    autosize_columns(ws)


def create_report_workbook(tickers: List[str], thresholds: Dict[str, Dict[str, Dict[str, Any]]],
                           metrics_by_ticker: Dict[str, Dict[str, Any]], reversal_by_ticker: Dict[str, Dict[str, Any]],
                           out_path: str, target_threshold: float = 60.0):
    wb = Workbook();
    wb.remove(wb.active)
    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.append(
        ["Ticker", "Sector", "Avg Fund Score", "Reversal Score", "Recommendation", "Valuation", "Quality", "Safety",
         "Growth", "Risk", "Coverage %"])
    for c in ws_sum[1]: c.fill = FILL_HDR; c.font = FONT_HDR; c.alignment = ALIGN_CENTER

    _add_cheat_sheet(wb, target_threshold)

    category_maps = {"Valuation": "Valuation", "Profitability": "Quality", "Balance Sheet": "Safety",
                     "Growth": "Growth", "Risk": "Risk"}
    for t in tickers:
        m = metrics_by_ticker.get(t, {});
        revpack = _normalize_reversal_pack(reversal_by_ticker.get(t, {}))
        bucket = m.get("Sector Bucket", "Default (All)")
        ws = wb.create_sheet(t);
        ws["A1"] = f"{t} — {bucket}";
        ws.merge_cells("A1:C1");
        ws["A1"].font = Font(size=14, bold=True)
        cat_scores, cat_coverages, row = {}, {}, 3
        for cat_sheet, cat_display in category_maps.items():
            ws.cell(row, 1, cat_display).font = Font(bold=True, size=12);
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6);
            row += 1
            for i, h in enumerate(["Metric", "Value", "Rating", "Mode", "Limits", "Notes"]):
                c = ws.cell(row, i + 1, h);
                c.fill = FILL_HDR;
                c.font = FONT_HDR
            row += 1
            ratings, weights = {}, {}
            for metric in thresholds.get(cat_sheet, {}):
                if metric not in WHITELIST: continue
                val = m.get(metric);
                th = get_threshold_set(thresholds, cat_sheet, metric, bucket)
                rating, fill = "NA", FILL_GRAY
                if th and val is not None:
                    rating, fill = score_with_threshold_txt(val, th.get("green_txt"), th.get("yellow_txt"),
                                                            th.get("red_txt"))
                ratings[metric], weights[metric] = rating, _metric_weight(metric)
                ws.cell(row, 1, metric);
                c = ws.cell(row, 2, val);
                _apply_metric_value_format(c, metric, val)
                ws.cell(row, 3, rating).fill = fill;
                ws.cell(row, 4, bucket);
                ws.cell(row, 5, _limits_text(th));
                ws.cell(row, 6, th.get("notes", ""));
                row += 1
            raw, cov = compute_category_score_and_coverage(ratings, weights)
            adj = adjusted_from_raw_and_coverage(raw, cov)
            cat_scores[cat_display], cat_coverages[cat_display] = adj, cov
            ws.cell(row, 1, f"{cat_display} Adjusted: {adj:.1f}%" if adj is not None else "NA").font = Font(bold=True);
            row += 2
        rev_total = revpack.get("total_score_pct")
        rec_txt, rec_fill = final_recommendation_banner(cat_scores, rev_total, target_threshold)
        valid_cats = [s for s in cat_scores.values() if s is not None]
        avg_f = sum(valid_cats) / len(valid_cats) if valid_cats else 0.0
        ws["D1"] = "Avg Fund Score";
        ws["E1"] = avg_f;
        ws["E1"].fill = band_fill(avg_f)
        ws["D2"] = "Reversal Score";
        ws["E2"] = rev_total;
        ws["E2"].fill = reversal_fill(rev_total)
        ws["A2"] = "Recommendation";
        ws["B2"] = rec_txt;
        ws["B2"].fill = rec_fill
        avg_cov = sum(cat_coverages.values()) / len(cat_coverages) if cat_coverages else 0
        ws_sum.append([t, bucket, avg_f, rev_total, rec_txt, cat_scores.get("Valuation"), cat_scores.get("Quality"),
                       cat_scores.get("Safety"), cat_scores.get("Growth"), cat_scores.get("Risk"), avg_cov])
        row = _write_reversal_block(ws, row, "Fundamental Turnaround", revpack.get("fund_symbols", {}),
                                    revpack.get("fund_details", {}), revpack.get("fund_score_pct"))
        row = _write_reversal_block(ws, row, "Technical Confirmation", revpack.get("tech_symbols", {}),
                                    revpack.get("tech_details", {}), revpack.get("tech_score_pct"))
        autosize_columns(ws)
    autosize_columns(ws_sum);
    wb.save(out_path)