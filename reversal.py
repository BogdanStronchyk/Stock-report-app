import math
from typing import Any, Dict, Optional, Tuple
import pandas as pd
import yfinance as yf


# ==========================
# Helpers
# ==========================
def _to_num(x):
    try:
        v = float(x)
        if math.isnan(v) or math.isinf(v): return None
        return v
    except:
        return None


def _safe_series(df: pd.DataFrame, row_names, n: int = 8) -> Optional[pd.Series]:
    if df is None or df.empty: return None
    for rn in row_names:
        if rn in df.index: return pd.to_numeric(df.loc[rn].iloc[:n], errors="coerce")
    return None


def _get_close_series(hist: pd.DataFrame) -> pd.Series:
    """Safely extract Close or Adj Close from history."""
    if hist is None or hist.empty: return pd.Series(dtype=float)
    col = "Adj Close" if "Adj Close" in hist.columns else "Close"
    if col in hist.columns:
        return pd.to_numeric(hist[col], errors="coerce").dropna()
    return pd.Series(dtype=float)


def _ttm_and_prev_ttm(series: pd.Series) -> Tuple[Optional[float], Optional[float]]:
    if series is None or series.dropna().shape[0] < 8: return (None, None)
    return (_to_num(series.iloc[:4].sum()), _to_num(series.iloc[4:8].sum()))


def _score_symbol(points: int) -> str:
    return "🟢" if points == 2 else ("🟡" if points == 1 else "🔴")


# ==========================
# Fundamental Turnaround
# ==========================
def _fund_margin_stabilization(q_income: pd.DataFrame) -> Tuple[int, str]:
    rev = _safe_series(q_income, ["Total Revenue", "TotalRevenue"], 8)
    opi = _safe_series(q_income, ["Operating Income", "OperatingIncome"], 8)
    if rev is None or opi is None: return (0, "Missing data")

    rev_ttm, rev_prev = _ttm_and_prev_ttm(rev)
    op_ttm, op_prev = _ttm_and_prev_ttm(opi)

    if None in (rev_ttm, rev_prev, op_ttm, op_prev) or rev_ttm == 0 or rev_prev == 0:
        return (0, "Insufficient TTM data")

    op_delta = (op_ttm / rev_ttm - op_prev / rev_prev) * 100
    if op_delta >= 1.0: return (2, f"Op Margin +{op_delta:.2f}pp")
    if op_delta >= 0.0: return (1, f"Op Margin stable (+{op_delta:.2f}pp)")
    return (0, "Margin deterioration")


def _fund_cashflow_reversal(q_cf: pd.DataFrame, q_income: pd.DataFrame) -> Tuple[int, str]:
    ocf = _safe_series(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 8)
    cap = _safe_series(q_cf, ["Capital Expenditure", "CapitalExpenditure"], 8)
    if ocf is None or cap is None: return (0, "Missing CF data")

    ocf_ttm, ocf_prev = _ttm_and_prev_ttm(ocf)
    cap_ttm, cap_prev = _ttm_and_prev_ttm(cap)

    if None in (ocf_ttm, ocf_prev): return (0, "Insufficient CF data")

    fcf_ttm = ocf_ttm + (cap_ttm or 0)
    fcf_prev = ocf_prev + (cap_prev or 0)

    if fcf_ttm > 0 and fcf_ttm > fcf_prev: return (2, "FCF Positive & Growing")
    if fcf_ttm > fcf_prev: return (1, "FCF Improving (still neg or weak)")
    return (0, "FCF Declining")


def _fund_balance_sheet_healing(annual_bs: pd.DataFrame) -> Tuple[int, str]:
    if annual_bs is None or annual_bs.empty: return (0, "Missing BS")
    debt = _safe_series(annual_bs, ["Total Debt", "TotalDebt"], 2)
    cash = _safe_series(annual_bs, ["Cash And Cash Equivalents"], 2)

    if debt is not None and cash is not None and len(debt) >= 2:
        nd_now = debt.iloc[0] - cash.iloc[0]
        nd_prev = debt.iloc[1] - cash.iloc[1]
        if nd_now < nd_prev: return (2, "Net Debt Decreasing")
        if nd_now < 0: return (2, "Net Cash Position")
    return (0, "Debt flat/increasing")


def _fund_roic_inflection(metrics: Dict[str, Any]) -> Tuple[int, str]:
    roic = metrics.get("ROIC % (standardized)")
    if roic is None: return (0, "NA")
    if roic > 15: return (2, f"ROIC High ({roic:.1f}%)")
    if roic > 8: return (1, f"ROIC Stable ({roic:.1f}%)")
    return (0, f"ROIC Weak ({roic:.1f}%)")


def _fund_value_divergence(metrics: Dict[str, Any]) -> Tuple[int, str]:
    ev_ebit = metrics.get("EV/EBIT")
    if ev_ebit is None: return (0, "NA")
    if ev_ebit < 10: return (2, f"Cheap (EV/EBIT {ev_ebit:.1f}x)")
    if ev_ebit < 18: return (1, f"Fair (EV/EBIT {ev_ebit:.1f}x)")
    return (0, f"Expensive ({ev_ebit:.1f}x)")


# ==========================
# Technical Confirmation
# ==========================
def _tech_price_above_200(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 200: return (0, "No Data")
    ma200 = close.rolling(200).mean().iloc[-1]
    curr = close.iloc[-1]
    if curr > ma200 * 1.01: return (2, "Price > MA200")
    if curr > ma200: return (1, "Testing MA200")
    return (0, "Below MA200")


def _tech_ma_structure(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 200: return (0, "No Data")
    ma50 = close.rolling(50).mean().iloc[-1]
    ma200 = close.rolling(200).mean().iloc[-1]
    if ma50 > ma200: return (2, "Golden Cross (50>200)")
    return (0, "Bearish Structure")


def _tech_rsi(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 20: return (0, "No Data")
    delta = close.diff()
    gain = (delta.where(delta > 0, 0)).rolling(14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(14).mean()
    rs = gain / loss
    rsi = 100 - (100 / (1 + rs)).iloc[-1]

    if rsi > 50: return (2, f"Bullish Momentum (RSI {rsi:.0f})")
    if rsi > 40: return (1, f"Neutral (RSI {rsi:.0f})")
    return (0, f"Bearish (RSI {rsi:.0f})")


def _tech_drawdown(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 50: return (0, "No Data")
    peak = close.rolling(252, min_periods=1).max().iloc[-1]
    curr = close.iloc[-1]
    dd = (curr / peak - 1) * 100
    if dd > -15: return (2, f"Near Highs ({dd:.1f}%)")
    if dd > -30: return (1, f"Recovering ({dd:.1f}%)")
    return (0, f"Deep Drawdown ({dd:.1f}%)")


# ==========================
# Main Scoring
# ==========================
FUND_WEIGHTS = {"Margins": 0.25, "Cashflow": 0.25, "Balance Sheet": 0.15, "ROIC": 0.15, "Valuation": 0.20}
TECH_WEIGHTS = {"Trend (MA200)": 0.30, "Structure (50/200)": 0.20, "Momentum (RSI)": 0.25, "Drawdown": 0.25}


def _calculate_weighted(items: Dict[str, Tuple[int, str]], weights: Dict[str, float]) -> float:
    score = 0.0
    total_w = 0.0
    for k, w in weights.items():
        if k in items:
            score += (items[k][0] / 2.0) * w
            total_w += w
    return (score / total_w * 100.0) if total_w > 0 else 0.0


def trend_reversal_scores_from_data(*, q_income=None, q_cf=None, annual_bs=None, h_1y=None, h_2y=None, metrics=None,
                                    **kwargs) -> Dict[str, Any]:
    # Fundamental
    fund = {}
    fund["Margins"] = _fund_margin_stabilization(q_income)
    fund["Cashflow"] = _fund_cashflow_reversal(q_cf, q_income)
    fund["Balance Sheet"] = _fund_balance_sheet_healing(annual_bs)
    fund["ROIC"] = _fund_roic_inflection(metrics or {})
    fund["Valuation"] = _fund_value_divergence(metrics or {})

    # Technical
    tech = {}
    tech["Trend (MA200)"] = _tech_price_above_200(h_2y)
    tech["Structure (50/200)"] = _tech_ma_structure(h_2y)
    tech["Momentum (RSI)"] = _tech_rsi(h_1y)
    tech["Drawdown"] = _tech_drawdown(h_1y)

    f_score = _calculate_weighted(fund, FUND_WEIGHTS)
    t_score = _calculate_weighted(tech, TECH_WEIGHTS)
    total = (0.6 * f_score) + (0.4 * t_score)

    return {
        "fundamental_score": f_score,
        "technical_score": t_score,
        "total_score_pct": total,
        "fund_symbols": {k: _score_symbol(v[0]) for k, v in fund.items()},
        "tech_symbols": {k: _score_symbol(v[0]) for k, v in tech.items()},
        "fund_details": {k: v for k, v in fund.items()},
        "tech_details": {k: v for k, v in tech.items()},
    }


def trend_reversal_scores(tkr: yf.Ticker, metrics: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    return trend_reversal_scores_from_data(
        q_income=tkr.quarterly_income_stmt,
        q_cf=tkr.quarterly_cashflow,
        annual_bs=tkr.balance_sheet,
        h_1y=tkr.history(period="1y"),
        h_2y=tkr.history(period="2y"),
        metrics=metrics
    )