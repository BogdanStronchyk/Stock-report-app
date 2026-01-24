from typing import Any, Dict, List
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from config import (
    FILL_HDR, FONT_HDR, ALIGN_CENTER, ALIGN_WRAP,
    FILL_GREEN, FILL_YELLOW, FILL_RED
)
from checklist_loader import get_threshold_set
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

def create_report_workbook(
    tickers: List[str],
    thresholds: Dict[str, Dict[str, Dict[str, Any]]],
    metrics_by_ticker: Dict[str, Dict[str, Any]],
    reversal_by_ticker: Dict[str, Dict[str, str]],
    out_path: str
):
    wb = Workbook()
    wb.remove(wb.active)

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
                val = m.get(metric)
                th = get_threshold_set(thresholds, cat_sheet, metric, bucket)

                if th:
                    score, fill = score_with_threshold_txt(val, th["green_txt"], th["yellow_txt"], th["red_txt"])
                    limits = f"{th['green_txt']} | {th['yellow_txt']} | {th['red_txt']}"
                    notes = th.get("notes") or ""
                    mode = bucket if bucket in thresholds[cat_sheet][metric] else "Default (All)"
                else:
                    score, fill = ("NA", None)
                    limits, notes, mode = "", "", "Default (All)"

                ws.cell(row, 1).value = metric
                ws.cell(row, 2).value = val
                ws.cell(row, 3).value = score
                ws.cell(row, 4).value = mode
                ws.cell(row, 5).value = limits
                ws.cell(row, 6).value = notes

                if fill:
                    ws.cell(row, 3).fill = fill

                for c in range(1, 7):
                    ws.cell(row, c).alignment = ALIGN_WRAP

                row += 1

            row += 1

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
