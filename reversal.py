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
        if rn in df.index:
            row = df.loc[rn]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            return pd.to_numeric(row.iloc[:n], errors="coerce")
    return None


def _get_close_series(hist: pd.DataFrame) -> pd.Series:
    if hist is None or hist.empty: return pd.Series(dtype=float)
    df = hist.copy()
    if isinstance(df.columns, pd.MultiIndex): df.columns = df.columns.get_level_values(0)
    col = "Adj Close" if "Adj Close" in df.columns else "Close"
    if col in df.columns: return pd.to_numeric(df[col], errors="coerce").dropna()
    return pd.Series(dtype=float)


def _calc_trend_score(current: Optional[float], previous: Optional[float], label: str) -> Tuple[int, str]:
    if current is None or previous is None: return (0, "Insufficient Data")
    if current > previous: return (2, f"{label} Rising")
    if current > previous * 0.95: return (1, f"{label} Stable")
    return (0, f"{label} Declining")


def _score_symbol(points: int) -> str:
    return "🟢" if points == 2 else ("🟡" if points == 1 else "🔴")


# ==========================
# Fundamental Turnaround (With Fallback)
# ==========================
def _fund_margin_trend(q_income: pd.DataFrame, a_income: pd.DataFrame) -> Tuple[int, str]:
    # 1. Try Quarterly TTM (Need 8 qtrs)
    q_rev = _safe_series(q_income, ["Total Revenue", "TotalRevenue"], 8)
    q_opi = _safe_series(q_income, ["Operating Income", "OperatingIncome"], 8)

    if q_rev is not None and len(q_rev) >= 8 and q_opi is not None and len(q_opi) >= 8:
        ttm_rev = q_rev.iloc[:4].sum();
        prev_rev = q_rev.iloc[4:8].sum()
        ttm_opi = q_opi.iloc[:4].sum();
        prev_opi = q_opi.iloc[4:8].sum()
        if ttm_rev > 0 and prev_rev > 0:
            m_now = ttm_opi / ttm_rev * 100
            m_prev = prev_opi / prev_rev * 100
            delta = m_now - m_prev
            if delta >= 0.5: return (2, f"Op Margin +{delta:.1f}pp (TTM)")
            if delta > -0.5: return (1, "Op Margin Stable (TTM)")
            return (0, "Op Margin Falling (TTM)")

    # 2. Fallback to Annual (Need 2 years)
    a_rev = _safe_series(a_income, ["Total Revenue", "TotalRevenue"], 2)
    a_opi = _safe_series(a_income, ["Operating Income", "OperatingIncome"], 2)
    if a_rev is not None and len(a_rev) >= 2 and a_opi is not None and len(a_opi) >= 2:
        rev_now = a_rev.iloc[0];
        rev_prev = a_rev.iloc[1]
        opi_now = a_opi.iloc[0];
        opi_prev = a_opi.iloc[1]
        if rev_now > 0 and rev_prev > 0:
            m_now = opi_now / rev_now * 100
            m_prev = opi_prev / rev_prev * 100
            if m_now > m_prev: return (2, "Op Margin Rising (Annual)")
            return (0, "Op Margin Falling (Annual)")

    return (0, "No Margin Trend Data")


def _fund_cashflow_trend(q_cf: pd.DataFrame, a_cf: pd.DataFrame) -> Tuple[int, str]:
    # 1. Try Quarterly TTM
    q_ocf = _safe_series(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 8)
    if q_ocf is not None and len(q_ocf) >= 8:
        ocf_now = q_ocf.iloc[:4].sum()
        ocf_prev = q_ocf.iloc[4:8].sum()
        return _calc_trend_score(ocf_now, ocf_prev, "OCF (TTM)")

    # 2. Fallback Annual
    a_ocf = _safe_series(a_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 2)
    if a_ocf is not None and len(a_ocf) >= 2:
        return _calc_trend_score(a_ocf.iloc[0], a_ocf.iloc[1], "OCF (Annual)")

    return (0, "No CF Trend Data")


def _fund_balance_sheet_healing(annual_bs: pd.DataFrame) -> Tuple[int, str]:
    if annual_bs is None or annual_bs.empty: return (0, "Missing BS")
    debt = _safe_series(annual_bs, ["Total Debt", "TotalDebt"], 2)
    cash = _safe_series(annual_bs, ["Cash And Cash Equivalents"], 2)
    if debt is not None and cash is not None and len(debt) >= 2:
        nd_now = debt.iloc[0] - cash.iloc[0]
        nd_prev = debt.iloc[1] - cash.iloc[1]
        if nd_now < nd_prev: return (2, "Net Debt Decreasing")
        if nd_now < 0: return (2, "Net Cash Position")
        if nd_now < nd_prev * 1.05: return (1, "Net Debt Stable")
    return (0, "Debt Increasing")


def _fund_roic_check(metrics: Dict[str, Any]) -> Tuple[int, str]:
    roic = metrics.get("ROIC % (standardized)")
    if roic is None: return (0, "NA")
    if roic > 12: return (2, f"ROIC Strong ({roic:.1f}%)")
    if roic > 6: return (1, f"ROIC Stable ({roic:.1f}%)")
    return (0, f"ROIC Weak ({roic:.1f}%)")


def _fund_value_check(metrics: Dict[str, Any]) -> Tuple[int, str]:
    val = metrics.get("EV/EBIT")
    if val is None: return (0, "NA")
    if val < 12: return (2, f"Value (EV/EBIT {val:.1f}x)")
    if val < 20: return (1, f"Fair (EV/EBIT {val:.1f}x)")
    return (0, f"Rich (EV/EBIT {val:.1f}x)")


# ==========================
# Technical Confirmation
# ==========================
def _tech_ma_trend(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 200: return (0, "No Data (<200d)")
    ma200 = close.rolling(200).mean().iloc[-1]
    price = close.iloc[-1]
    if price > ma200: return (2, "Price > 200d MA")
    if price > ma200 * 0.95: return (1, "Testing 200d MA")
    return (0, "Below 200d MA")


def _tech_rsi_mom(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 20: return (0, "No Data")
    delta = close.diff()
    gain = (delta.where(delta > 0, 0)).rolling(14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(14).mean()
    rs = gain / loss
    rsi = 100 - (100 / (1 + rs)).iloc[-1]
    if rsi > 50: return (2, f"Bullish (RSI {rsi:.0f})")
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
    return (0, f"In Drawdown ({dd:.1f}%)")


def _tech_structure(hist: pd.DataFrame) -> Tuple[int, str]:
    close = _get_close_series(hist)
    if len(close) < 55: return (0, "No Data")
    ma50 = close.rolling(50).mean().iloc[-1]
    ma20 = close.rolling(20).mean().iloc[-1]
    if ma20 > ma50: return (2, "Short-term Bullish (20>50)")
    return (0, "Short-term Bearish")


# ==========================
# Main Scoring
# ==========================
FUND_WEIGHTS = {"Margins": 0.3, "Cashflow": 0.3, "Balance Sheet": 0.2, "ROIC": 0.1, "Valuation": 0.1}
TECH_WEIGHTS = {"Trend (200d)": 0.4, "Momentum (RSI)": 0.2, "Drawdown": 0.2, "Structure": 0.2}


def _calculate_weighted(items: Dict[str, Tuple[int, str]], weights: Dict[str, float]) -> float:
    score = 0.0
    total_w = 0.0
    for k, w in weights.items():
        if k in items:
            score += (items[k][0] / 2.0) * w
            total_w += w
    return (score / total_w * 100.0) if total_w > 0 else 0.0


def trend_reversal_scores_from_data(*, q_income=None, q_cf=None, annual_income=None, annual_cf=None, annual_bs=None,
                                    h_1y=None, h_2y=None, metrics=None, **kwargs) -> Dict[str, Any]:
    fund = {}
    fund["Margins"] = _fund_margin_trend(q_income, annual_income)
    fund["Cashflow"] = _fund_cashflow_trend(q_cf, annual_cf)
    fund["Balance Sheet"] = _fund_balance_sheet_healing(annual_bs)
    fund["ROIC"] = _fund_roic_check(metrics or {})
    fund["Valuation"] = _fund_value_check(metrics or {})

    tech = {}
    tech["Trend (200d)"] = _tech_ma_trend(h_2y)
    tech["Momentum (RSI)"] = _tech_rsi_mom(h_1y)
    tech["Drawdown"] = _tech_drawdown(h_1y)
    tech["Structure"] = _tech_structure(h_1y)

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