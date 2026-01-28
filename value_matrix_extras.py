import math
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

_NUM = r"[+-]?\d*\.?\d+"

def _to_float(x: Any) -> Optional[float]:
    if x is None:
        return None
    try:
        if isinstance(x, (int, float, np.number)):
            v = float(x)
            if math.isnan(v) or math.isinf(v):
                return None
            return v
    except Exception:
        pass
    if isinstance(x, str):
        s = x.strip()
        if not s or s.lower() in {"na", "n/a", "none", "null", "-"}:
            return None
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1].strip()
        s = s.replace(",", "").replace("_", "").replace(" ", "")
        is_pct = s.endswith("%")
        if is_pct:
            s = s[:-1]
        # strip leading currency symbols
        while s and not (s[0].isdigit() or s[0] in "+-."):
            s = s[1:]
        if not s:
            return None
        mul = 1.0
        suf = s[-1].upper()
        if suf in ("K", "M", "B", "T"):
            mul = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}[suf]
            s = s[:-1]
        try:
            v = float(s) * mul
            if neg:
                v = -v
            if is_pct:
                v = v / 100.0
            if math.isnan(v) or math.isinf(v):
                return None
            return v
        except Exception:
            return None
    return None


def _annual_series(df: pd.DataFrame, row_names: List[str], n: int = 6) -> List[Optional[float]]:
    """Oldest->newest series of up to n periods for the first matching row."""
    if df is None or getattr(df, "empty", True):
        return []
    for rn in row_names:
        if rn in df.index:
            s = df.loc[rn]
            vals = pd.to_numeric(s.iloc[:n][::-1], errors="coerce")
            out: List[Optional[float]] = []
            for v in vals.values:
                if pd.isna(v):
                    out.append(None)
                else:
                    out.append(float(v))
            return out
    return []


def _quarterly_ttm(df: pd.DataFrame, row_names: List[str], n: int = 4) -> Optional[float]:
    if df is None or getattr(df, "empty", True):
        return None
    cols = list(df.columns)[:n]
    for rn in row_names:
        if rn in df.index:
            s = pd.to_numeric(df.loc[rn, cols], errors="coerce")
            if s.isna().all():
                return None
            return float(s.sum())
    return None


def _winsorize(xs: List[float], p_lo: float = 0.10, p_hi: float = 0.90) -> List[float]:
    ys = [float(v) for v in xs if v is not None and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))]
    if len(ys) < 3:
        return ys
    lo = float(np.quantile(ys, p_lo))
    hi = float(np.quantile(ys, p_hi))
    return [min(max(y, lo), hi) for y in ys]


def _median_pos(xs: List[float]) -> Optional[float]:
    ys = [float(v) for v in xs if v is not None and float(v) > 0 and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))]
    if not ys:
        return None
    return float(np.median(ys))


def _cagr(values: List[Optional[float]]) -> Optional[float]:
    xs = [v for v in values if v is not None and v > 0 and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))]
    if len(xs) < 2:
        return None
    first, last = xs[0], xs[-1]
    n = len(xs) - 1
    if n <= 0 or first <= 0 or last <= 0:
        return None
    return ((last / first) ** (1 / n) - 1) * 100.0


def _normalize_div_yield(div_yield_raw: Any) -> Optional[float]:
    """Normalize dividend yield into percent points (e.g., 0.5 means 0.5%).

    Yahoo usually provides a fraction (0.005). Some sources return already in percent.
    We add sanity to avoid common 100x mis-scaling (e.g., 0.4 being treated as 40%).
    """
    v = _to_float(div_yield_raw)
    if v is None:
        return None

    # Base normalization
    pct = float(v) * 100.0 if abs(v) <= 1.5 else float(v)

    # Sanity: yields >20% are rare for liquid large-caps; if we got here via the fraction branch,
    # interpret input as already percent and undo the 100x.
    if pct > 20.0 and abs(v) <= 1.5:
        pct = float(v)

    # Final guardrail
    if pct < 0.0 or pct > 60.0:
        return None
    return pct


def compute_value_matrix_extras(
    *,
    info: Dict[str, Any],
    fmp_bundle: Optional[Dict[str, Any]],
    price: Optional[float],
    market_cap: Optional[float],
    enterprise_value: Optional[float],
    annual_income: pd.DataFrame,
    annual_cashflow: pd.DataFrame,
    annual_balance_sheet: pd.DataFrame,
    quarterly_cashflow: pd.DataFrame,
) -> Tuple[Dict[str, Any], Dict[str, str]]:
    """Compute additional value-investing decision-matrix metrics.

    Returns (extra_metrics, extra_notes). Metrics keys are designed to match checklist names.
    """
    out: Dict[str, Any] = {}
    notes: Dict[str, str] = {}

    # ---------- Dividend yield ----------
    div_y = _normalize_div_yield(info.get("dividendYield"))
    if div_y is None and fmp_bundle:
        # FMP profile often returns 'lastDiv' and 'mktCap', not a yield; quote may have 'price'.
        # We keep this conservative: only use it if 'dividendYield' exists.
        prof = (fmp_bundle.get("profile") or [{}])
        if isinstance(prof, list) and prof and isinstance(prof[0], dict):
            div_y = _normalize_div_yield(prof[0].get("dividendYield"))
    out["Dividend Yield %"] = div_y

    # ---------- Annual FCF series (for normalized ratios) ----------
    ocf = _annual_series(annual_cashflow, ["Operating Cash Flow", "Total Cash From Operating Activities", "OperatingCashFlow"], 8)
    cap = _annual_series(annual_cashflow, ["Capital Expenditure", "CapitalExpenditure"], 8)

    fcf_series: List[Optional[float]] = []
    if ocf and cap:
        for o, c in zip(ocf, cap):
            fcf_series.append((o + c) if (o is not None and c is not None) else None)

    # Normalized FCF: 5Y median, winsorized, positive only
    fcf_5y = [v for v in fcf_series[-5:] if v is not None] if fcf_series else []
    fcf_norm = _median_pos(_winsorize(fcf_5y)) if len(fcf_5y) >= 3 else None

    if fcf_norm is None:
        notes["EV/FCF (Normalized 5Y Median)"] = "Normalized FCF unavailable (need >=3 annual FCF points, positive)."
        notes["FCF Yield (Normalized 5Y Median)"] = "Normalized FCF unavailable (need >=3 annual FCF points, positive)."

    if enterprise_value not in (None, 0) and fcf_norm not in (None, 0):
        out["EV/FCF (Normalized 5Y Median)"] = float(enterprise_value) / float(fcf_norm)
    else:
        out["EV/FCF (Normalized 5Y Median)"] = None

    if market_cap not in (None, 0) and fcf_norm not in (None, 0):
        out["FCF Yield (Normalized 5Y Median)"] = float(fcf_norm) / float(market_cap) * 100.0
    else:
        out["FCF Yield (Normalized 5Y Median)"] = None

    # ---------- Share count trend (3Y) ----------
    shares_series = _annual_series(
        annual_balance_sheet,
        [
            "Ordinary Shares Number",
            "OrdinarySharesNumber",
            "Share Issued",
            "ShareIssued",
            "Common Stock Shares Outstanding",
            "CommonStockSharesOutstanding",
            "IssuedCommonStock",
        ],
        6,
    )
    share_cagr_3y = None
    if len([v for v in shares_series if v is not None]) >= 4:
        share_cagr_3y = _cagr(shares_series[-4:])
    out["Share Count CAGR (3Y)"] = share_cagr_3y
    if share_cagr_3y is None:
        notes["Share Count CAGR (3Y)"] = "Share count series not available in balance sheet (needs >=4 annual points)."

    # ---------- Net buyback yield (3Y avg) ----------
    # Uses annual cash flow net repurchase outflow (repurchases + issuance).
    rep = _annual_series(
        annual_cashflow,
        ["Repurchase Of Capital Stock", "RepurchaseOfCapitalStock", "Common Stock Repurchased", "CommonStockRepurchased"],
        6,
    )
    iss = _annual_series(
        annual_cashflow,
        ["Issuance Of Capital Stock", "IssuanceOfCapitalStock", "Common Stock Issued", "CommonStockIssued"],
        6,
    )
    net_outflows: List[Optional[float]] = []
    if rep and iss:
        for r, i in zip(rep, iss):
            if r is None and i is None:
                net_outflows.append(None)
            else:
                rr = 0.0 if r is None else float(r)
                ii = 0.0 if i is None else float(i)
                # typical: repurchase negative (outflow), issuance positive (inflow)
                net_outflows.append(-(rr + ii))
    # avg of last 3 years where available
    buyback_yield_3y = None
    usable = [v for v in net_outflows[-3:] if v is not None] if net_outflows else []
    if market_cap not in (None, 0) and usable:
        buyback_yield_3y = float(np.mean(usable)) / float(market_cap) * 100.0
    out["Net Buyback Yield (3Y avg)"] = buyback_yield_3y
    if buyback_yield_3y is None:
        notes["Net Buyback Yield (3Y avg)"] = "Buyback yield unavailable (need repurchase/issuance cashflow + market cap)."

    # Shareholder yield (existing checklist metric)
    sh_yield = None
    if div_y is not None and buyback_yield_3y is not None:
        sh_yield = float(div_y) + float(buyback_yield_3y)
        notes["Shareholder Yield (Dividend + Buyback yield)"] = "Dividend yield (current) + net buyback yield (3Y avg)."
    out["Shareholder Yield (Dividend + Buyback yield)"] = sh_yield

    # ---------- Stock-based compensation burden ----------
    sbc_ttm = _quarterly_ttm(
        quarterly_cashflow,
        ["Stock Based Compensation", "StockBasedCompensation"],
        4,
    )
    sbc_pct_mcap = None
    sbc_pct_fcf = None
    if sbc_ttm is not None and market_cap not in (None, 0):
        sbc_pct_mcap = float(sbc_ttm) / float(market_cap) * 100.0
    if sbc_ttm is not None:
        # try to infer TTM FCF from quarterly cashflow as OCF + CapEx
        ocf_ttm = _quarterly_ttm(quarterly_cashflow, ["Operating Cash Flow", "Total Cash From Operating Activities"], 4)
        cap_ttm = _quarterly_ttm(quarterly_cashflow, ["Capital Expenditure", "CapitalExpenditure"], 4)
        fcf_ttm = (ocf_ttm + cap_ttm) if (ocf_ttm is not None and cap_ttm is not None) else None
        if fcf_ttm not in (None, 0):
            sbc_pct_fcf = float(sbc_ttm) / float(fcf_ttm) * 100.0

    out["SBC % of Market Cap (TTM)"] = sbc_pct_mcap
    out["SBC % of FCF (TTM)"] = sbc_pct_fcf
    if sbc_pct_mcap is None:
        notes["SBC % of Market Cap (TTM)"] = "SBC unavailable (missing quarterly cashflow row or market cap)."
    if sbc_pct_fcf is None:
        notes["SBC % of FCF (TTM)"] = "SBC/FCF unavailable (needs SBC + positive TTM FCF)."

    # ---------- ROIC trend (3Y, pp) ----------
    # Recompute ROIC proxy by year: NOPAT / (Assets - Current Liabilities)
    ebit_series = _annual_series(annual_income, ["EBIT", "Ebit", "Operating Income", "OperatingIncome"], 6)
    assets_series = _annual_series(annual_balance_sheet, ["Total Assets", "TotalAssets"], 6)
    cl_series = _annual_series(
        annual_balance_sheet,
        ["Current Liabilities", "CurrentLiabilities", "Total Current Liabilities"],
        6,
    )

    roic_series: List[Optional[float]] = []
    if ebit_series and assets_series and cl_series:
        # align by shortest length
        n = min(len(ebit_series), len(assets_series), len(cl_series))
        for i in range(n):
            e = ebit_series[-n + i]
            a = assets_series[-n + i]
            cl = cl_series[-n + i]
            if e is None or a is None or cl is None:
                roic_series.append(None)
                continue
            invested = float(a) - float(cl)
            if invested <= 0:
                roic_series.append(None)
                continue
            nopat = float(e) * (1.0 - 0.21)
            roic_series.append((nopat / invested) * 100.0)

    roic_delta = None
    usable_roic = [v for v in roic_series if v is not None] if roic_series else []
    if roic_series and len([v for v in roic_series[-4:] if v is not None]) >= 2 and roic_series[-1] is not None and roic_series[-4] is not None:
        roic_delta = float(roic_series[-1]) - float(roic_series[-4])
    out["ROIC Δ (3Y, pp)"] = roic_delta
    if roic_delta is None:
        notes["ROIC Δ (3Y, pp)"] = "ROIC trend unavailable (needs >=4 annual points for EBIT, Assets, Current Liabilities)."

    # ---------- Margin trend (3–5Y) ----------
    rev_series = _annual_series(annual_income, ["Total Revenue", "TotalRevenue"], 6)
    opinc_series = _annual_series(annual_income, ["Operating Income", "OperatingIncome", "EBIT", "Ebit"], 6)

    margin_series: List[Optional[float]] = []
    if rev_series and opinc_series:
        n = min(len(rev_series), len(opinc_series))
        for i in range(n):
            r = rev_series[-n + i]
            op = opinc_series[-n + i]
            if r in (None, 0) or op is None:
                margin_series.append(None)
            else:
                margin_series.append(float(op) / float(r) * 100.0)

    margin_trend = None
    # use 5y if we have 5, else 3y if we have 3+
    series_clean = [v for v in margin_series if v is not None] if margin_series else []
    if margin_series and margin_series[-1] is not None:
        if len([v for v in margin_series[-6:] if v is not None]) >= 5 and margin_series[-5] is not None:
            margin_trend = float(margin_series[-1]) - float(margin_series[-5])
        elif len([v for v in margin_series[-4:] if v is not None]) >= 3 and margin_series[-3] is not None:
            margin_trend = float(margin_series[-1]) - float(margin_series[-3])

    out["Margin Trend (3–5Y)"] = margin_trend
    if margin_trend is None:
        notes["Margin Trend (3–5Y)"] = "Margin trend unavailable (needs multi-year revenue + operating income)."

    return out, notes
