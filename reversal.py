
import math
from typing import Any, Dict, Optional, Tuple

import numpy as np
import pandas as pd
import yfinance as yf


# ==========================
# Helpers
# ==========================
def _to_num(x):
    try:
        v = float(x)
        if math.isnan(v) or math.isinf(v):
            return None
        return v
    except Exception:
        return None


def _safe_series(df: pd.DataFrame, row_names, n: int = 8) -> Optional[pd.Series]:
    if df is None or df.empty:
        return None
    for rn in row_names:
        if rn in df.index:
            return pd.to_numeric(df.loc[rn].iloc[:n], errors="coerce")
    return None


def _ttm_and_prev_ttm(series: pd.Series) -> Tuple[Optional[float], Optional[float]]:
    if series is None or series.dropna().shape[0] < 8:
        return (None, None)
    return (_to_num(series.iloc[:4].sum()), _to_num(series.iloc[4:8].sum()))


def _pct_change(curr: Optional[float], prev: Optional[float]) -> Optional[float]:
    if curr is None or prev is None or prev == 0:
        return None
    return (curr / prev - 1.0) * 100.0


def _score_symbol(points: int) -> str:
    return "🟢" if points == 2 else ("🟡" if points == 1 else "🔴")


def _rsi(close: pd.Series, period: int = 14) -> Optional[float]:
    close = pd.to_numeric(close, errors="coerce").dropna()
    if close.shape[0] < period + 5:
        return None
    delta = close.diff()
    gain = delta.clip(lower=0).rolling(period).mean()
    loss = (-delta.clip(upper=0)).rolling(period).mean()
    rs = gain / loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))
    return _to_num(rsi.iloc[-1])


def _atr_percent(hist: pd.DataFrame, period: int = 14) -> Optional[float]:
    if hist is None or hist.empty:
        return None
    for c in ["High", "Low", "Close"]:
        if c not in hist.columns:
            return None
    high = pd.to_numeric(hist["High"], errors="coerce")
    low = pd.to_numeric(hist["Low"], errors="coerce")
    close = pd.to_numeric(hist["Close"], errors="coerce")
    prev_close = close.shift(1)
    tr = pd.concat([(high - low).abs(), (high - prev_close).abs(), (low - prev_close).abs()], axis=1).max(axis=1)
    atr = tr.rolling(period).mean()
    last_atr = _to_num(atr.iloc[-1])
    last_close = _to_num(close.iloc[-1])
    if last_atr is None or last_close in (None, 0):
        return None
    return (last_atr / last_close) * 100.0


# ==========================
# Fundamental Turnaround (7)
# ==========================
def _fund_margin_stabilization(q_income: pd.DataFrame) -> Tuple[int, str]:
    rev = _safe_series(q_income, ["Total Revenue", "TotalRevenue"], 8)
    opi = _safe_series(q_income, ["Operating Income", "OperatingIncome"], 8)
    gp = _safe_series(q_income, ["Gross Profit", "GrossProfit"], 8)

    rev_ttm, rev_prev = _ttm_and_prev_ttm(rev) if rev is not None else (None, None)
    op_ttm, op_prev = _ttm_and_prev_ttm(opi) if opi is not None else (None, None)
    gp_ttm, gp_prev = _ttm_and_prev_ttm(gp) if gp is not None else (None, None)

    if None in (rev_ttm, rev_prev, op_ttm, op_prev) or rev_ttm in (None, 0) or rev_prev in (None, 0):
        return (0, "Missing quarterly revenue/op income for TTM comparison.")

    opm_ttm = op_ttm / rev_ttm * 100.0
    opm_prev = op_prev / rev_prev * 100.0
    op_delta = opm_ttm - opm_prev

    gm_delta = 0.0
    if None not in (gp_ttm, gp_prev):
        gm_ttm = gp_ttm / rev_ttm * 100.0
        gm_prev = gp_prev / rev_prev * 100.0
        gm_delta = gm_ttm - gm_prev

    if op_delta >= 1.0 and gm_delta >= -0.5:
        return (2, f"Op margin +{op_delta:.2f}pp, GM Δ {gm_delta:.2f}pp.")
    if op_delta >= 0.0:
        return (1, f"Op margin +{op_delta:.2f}pp.")
    return (0, f"Op margin {op_delta:.2f}pp deterioration.")


def _fund_cashflow_reversal(q_cf: pd.DataFrame, q_income: pd.DataFrame) -> Tuple[int, str]:
    ocf = _safe_series(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 8)
    cap = _safe_series(q_cf, ["Capital Expenditure", "CapitalExpenditure"], 8)
    ni = _safe_series(q_income, ["Net Income", "NetIncome"], 8)

    ocf_ttm, ocf_prev = _ttm_and_prev_ttm(ocf) if ocf is not None else (None, None)
    cap_ttm, cap_prev = _ttm_and_prev_ttm(cap) if cap is not None else (None, None)
    ni_ttm, _ = _ttm_and_prev_ttm(ni) if ni is not None else (None, None)

    if None in (ocf_ttm, ocf_prev, cap_ttm, cap_prev):
        return (0, "Missing cashflow data for TTM FCF.")

    fcf_ttm = ocf_ttm + cap_ttm  # capex is negative
    fcf_prev = ocf_prev + cap_prev
    improved = fcf_ttm > fcf_prev

    cfo_ni = None
    if ocf_ttm is not None and ni_ttm not in (None, 0):
        cfo_ni = ocf_ttm / ni_ttm

    if fcf_ttm > 0 and (improved or (fcf_prev is not None and fcf_prev <= 0)):
        if cfo_ni is None or cfo_ni >= 1.0:
            return (2, f"FCF positive & improving. CFO/NI={cfo_ni:.2f}" if cfo_ni is not None else "FCF positive & improving.")
        return (1, f"FCF positive but CFO/NI weak ({cfo_ni:.2f}).")

    if improved:
        return (1, "FCF improving but still weak/negative.")
    return (0, "FCF not improving.")


def _fund_balance_sheet_healing(annual_bs: pd.DataFrame, q_income: pd.DataFrame) -> Tuple[int, str]:
    if annual_bs is None or annual_bs.empty:
        return (0, "Missing annual balance sheet.")

    cash = _safe_series(annual_bs, ["Cash And Cash Equivalents", "CashAndCashEquivalents"], 2)
    debt = _safe_series(annual_bs, ["Total Debt", "TotalDebt", "Long Term Debt", "LongTermDebt"], 2)

    net_debt_trend = None
    if cash is not None and debt is not None and cash.dropna().shape[0] >= 2 and debt.dropna().shape[0] >= 2:
        nd_now = _to_num(debt.iloc[0]) - _to_num(cash.iloc[0])
        nd_prev = _to_num(debt.iloc[1]) - _to_num(cash.iloc[1])
        if None not in (nd_now, nd_prev):
            net_debt_trend = nd_now - nd_prev

    op = _safe_series(q_income, ["Operating Income", "OperatingIncome"], 8)
    ie = _safe_series(q_income, ["Interest Expense", "InterestExpense"], 8)
    op_ttm, op_prev = _ttm_and_prev_ttm(op) if op is not None else (None, None)
    ie_ttm, ie_prev = _ttm_and_prev_ttm(ie) if ie is not None else (None, None)

    ic = None
    ic_prev = None
    if op_ttm is not None and ie_ttm not in (None, 0):
        ic = op_ttm / abs(ie_ttm)
    if op_prev is not None and ie_prev not in (None, 0):
        ic_prev = op_prev / abs(ie_prev)

    debt_ok = net_debt_trend is not None and net_debt_trend < 0
    ic_ok = ic is not None and ic >= 2.0
    ic_improve = ic is not None and ic_prev is not None and (ic - ic_prev) >= 0.5

    if debt_ok and (ic_ok or ic_improve):
        return (2, f"Net debt decreasing; IC {ic:.2f}.")
    if debt_ok or ic_ok or ic_improve:
        return (1, "Partial leverage/IC improvement.")
    return (0, "No clear balance-sheet improvement.")


def _fund_roic_inflection(metrics: Optional[Dict[str, Any]]) -> Tuple[int, str]:
    if not metrics:
        return (0, "Missing ROIC proxy.")
    roic = metrics.get("ROIC % (standardized)")
    if roic is None:
        return (0, "ROIC proxy unavailable.")
    if roic >= 10:
        return (2, f"ROIC healthy ({roic:.1f}%).")
    if roic >= 6:
        return (1, f"ROIC ok ({roic:.1f}%).")
    return (0, f"ROIC weak ({roic:.1f}%).")


def _fund_capex_discipline(q_cf: pd.DataFrame, q_income: pd.DataFrame) -> Tuple[int, str]:
    cap = _safe_series(q_cf, ["Capital Expenditure", "CapitalExpenditure"], 8)
    ocf = _safe_series(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 8)
    rev = _safe_series(q_income, ["Total Revenue", "TotalRevenue"], 8)

    cap_ttm, cap_prev = _ttm_and_prev_ttm(cap) if cap is not None else (None, None)
    ocf_ttm, ocf_prev = _ttm_and_prev_ttm(ocf) if ocf is not None else (None, None)
    rev_ttm, rev_prev = _ttm_and_prev_ttm(rev) if rev is not None else (None, None)

    if None in (cap_ttm, cap_prev, ocf_ttm, ocf_prev, rev_ttm, rev_prev) or rev_ttm in (None, 0) or rev_prev in (None, 0):
        return (0, "Missing capex/revenue for intensity trend.")

    fcf_ttm = ocf_ttm + cap_ttm
    fcf_prev = ocf_prev + cap_prev
    fcf_improve = fcf_ttm > fcf_prev

    cap_int = abs(cap_ttm) / rev_ttm * 100.0
    cap_int_prev = abs(cap_prev) / rev_prev * 100.0
    cap_down = (cap_int - cap_int_prev) <= -0.2
    cap_flat = abs(cap_int - cap_int_prev) < 0.2

    rev_chg = _pct_change(rev_ttm, rev_prev)

    if (cap_down or cap_flat) and (rev_chg is None or rev_chg >= -3.0) and fcf_improve:
        return (2 if cap_down else 1, f"Capex int {cap_int:.2f}% (prev {cap_int_prev:.2f}%), FCF improving.")
    if fcf_improve:
        return (1, "FCF improving; capex intensity not clearly improving.")
    return (0, "No capex discipline improvement.")


def _fund_revenue_quality(q_income: pd.DataFrame, info: Dict[str, Any]) -> Tuple[int, str]:
    rev = _safe_series(q_income, ["Total Revenue", "TotalRevenue"], 8)
    gp = _safe_series(q_income, ["Gross Profit", "GrossProfit"], 8)
    rev_ttm, rev_prev = _ttm_and_prev_ttm(rev) if rev is not None else (None, None)

    if None in (rev_ttm, rev_prev) or rev_prev in (None, 0):
        return (0, "Missing revenue history.")

    rev_chg = _pct_change(rev_ttm, rev_prev)

    gm_delta = 0.0
    if gp is not None:
        gp_ttm, gp_prev = _ttm_and_prev_ttm(gp)
        if None not in (gp_ttm, gp_prev) and rev_ttm and rev_prev:
            gm_delta = (gp_ttm / rev_ttm * 100.0) - (gp_prev / rev_prev * 100.0)

    shares = info.get("sharesOutstanding")
    try:
        shares = float(shares) if shares else None
    except Exception:
        shares = None

    revps_chg = None
    if shares and shares > 0:
        revps_chg = _pct_change(rev_ttm / shares, rev_prev / shares)

    if rev_chg is not None and rev_chg >= 0 and gm_delta >= -0.5 and (revps_chg is None or revps_chg >= -1.0):
        return (2, f"Rev {rev_chg:.1f}%, GM Δ {gm_delta:.2f}pp.")
    if rev_chg is not None and rev_chg >= -3.0 and gm_delta >= -1.0:
        return (1, f"Rev stabilizing ({rev_chg:.1f}%), GM Δ {gm_delta:.2f}pp.")
    return (0, f"Rev weak ({rev_chg:.1f}%), GM Δ {gm_delta:.2f}pp.")


def _fund_value_divergence(metrics: Optional[Dict[str, Any]]) -> Tuple[int, str]:
    if not metrics:
        return (0, "Missing valuation context.")
    ps = metrics.get("P/S")
    ev_fcf = metrics.get("EV/FCF")
    regime = metrics.get("NUPL Regime", "NA")

    flags = 0
    if ps is not None and ps <= 1.5:
        flags += 1
    if ev_fcf is not None and ev_fcf <= 18:
        flags += 1
    if regime in ("Stress", "Neutral", "Capitulation", "Deep Capitulation"):
        flags += 1

    if flags >= 3:
        return (2, f"Cheap + supportive (P/S={ps}, EV/FCF={ev_fcf}, NUPL={regime}).")
    if flags == 2:
        return (1, f"Partial cheapness (P/S={ps}, EV/FCF={ev_fcf}, NUPL={regime}).")
    return (0, f"No divergence (P/S={ps}, EV/FCF={ev_fcf}, NUPL={regime}).")


# ==========================
# Technical Confirmation (7)
# ==========================
def _tech_price_above_200(hist: pd.DataFrame) -> Tuple[int, str]:
    if hist is None or hist.empty:
        return (0, "Missing history.")
    close = pd.to_numeric(hist.get("Adj Close", hist.get("Close")), errors="coerce").dropna()
    if close.shape[0] < 210:
        return (0, "Insufficient history for MA200.")
    ma200 = close.rolling(200).mean().iloc[-1]
    p = close.iloc[-1]
    if p > ma200 * 1.02:
        return (2, "Price > MA200 by >2%.")
    if p > ma200:
        return (1, "Price above MA200 (weak).")
    return (0, "Price below MA200.")


def _tech_ma_structure(hist: pd.DataFrame) -> Tuple[int, str]:
    if hist is None or hist.empty:
        return (0, "Missing history.")
    close = pd.to_numeric(hist.get("Adj Close", hist.get("Close")), errors="coerce").dropna()
    if close.shape[0] < 210:
        return (0, "Insufficient history.")
    ma50 = close.rolling(50).mean()
    ma200 = close.rolling(200).mean()
    slope50 = _to_num(ma50.iloc[-1] - ma50.iloc[-21])
    if ma50.iloc[-1] > ma200.iloc[-1] and slope50 is not None and slope50 > 0:
        return (2, "MA50 > MA200 and rising.")
    if slope50 is not None and slope50 > 0:
        return (1, "MA50 rising.")
    return (0, "MA structure bearish.")


def _tech_breakout_higher_low(hist: pd.DataFrame) -> Tuple[int, str]:
    if hist is None or hist.empty:
        return (0, "Missing history.")
    close = pd.to_numeric(hist.get("Adj Close", hist.get("Close")), errors="coerce").dropna()
    if close.shape[0] < 80:
        return (0, "Insufficient history.")
    recent = close.iloc[-21:-1]
    prior = close.iloc[-41:-21]
    breakout = close.iloc[-1] > recent.max()
    hl = recent.min() > prior.min()
    if breakout and hl:
        return (2, "Higher low + 20D breakout.")
    if breakout or hl:
        return (1, "Partial: breakout or higher low.")
    return (0, "No breakout/higher-low.")


def _tech_rsi(hist: pd.DataFrame) -> Tuple[int, str]:
    if hist is None or hist.empty:
        return (0, "Missing history.")
    close = pd.to_numeric(hist.get("Adj Close", hist.get("Close")), errors="coerce").dropna()
    r = _rsi(close, 14)
    if r is None:
        return (0, "RSI unavailable.")
    if r >= 55:
        return (2, f"RSI {r:.1f} bullish.")
    if r >= 50:
        return (1, f"RSI {r:.1f} neutral+.")
    return (0, f"RSI {r:.1f} weak.")


def _tech_volume_accumulation(hist: pd.DataFrame) -> Tuple[int, str]:
    if hist is None or hist.empty or "Volume" not in hist.columns:
        return (0, "Missing volume.")
    close = pd.to_numeric(hist.get("Adj Close", hist.get("Close")), errors="coerce").dropna()
    vol = pd.to_numeric(hist["Volume"], errors="coerce").fillna(0)
    if close.shape[0] < 60:
        return (0, "Insufficient history.")
    ret = close.pct_change().fillna(0)
    up = vol.where(ret > 0).rolling(20).sum().iloc[-1]
    dn = vol.where(ret < 0).rolling(20).sum().iloc[-1]
    if dn <= 0:
        return (1, "No down-volume; unclear.")
    ratio = up / dn
    if ratio >= 1.5:
        return (2, f"Up-volume dominance {ratio:.2f}.")
    if ratio >= 1.1:
        return (1, f"Moderate accumulation {ratio:.2f}.")
    return (0, f"Distribution {ratio:.2f}.")


def _tech_volatility_regime(hist_1y: pd.DataFrame) -> Tuple[int, str]:
    if hist_1y is None or hist_1y.empty:
        return (0, "Missing history.")
    atrp = _atr_percent(hist_1y, 14)
    if atrp is None:
        return (0, "ATR% unavailable.")
    close = pd.to_numeric(hist_1y.get("Adj Close", hist_1y.get("Close")), errors="coerce").dropna()
    if close.shape[0] < 220:
        return (0, "Insufficient history.")
    samples = []
    for i in range(200, close.shape[0]):
        v = _atr_percent(hist_1y.iloc[i-60:i], 14)
        if v is not None:
            samples.append(v)
    if len(samples) < 30:
        return (1, f"ATR%={atrp:.2f}% (median NA).")
    med = float(np.median(samples))
    if atrp < med * 0.85:
        return (2, f"Volatility compressed (ATR% {atrp:.2f} < med {med:.2f}).")
    if atrp < med * 1.05:
        return (1, f"Volatility normalizing (ATR% {atrp:.2f}).")
    return (0, f"Volatility elevated (ATR% {atrp:.2f}).")


def _tech_drawdown_recovery(hist_1y: pd.DataFrame) -> Tuple[int, str]:
    if hist_1y is None or hist_1y.empty:
        return (0, "Missing history.")
    close = pd.to_numeric(hist_1y.get("Adj Close", hist_1y.get("Close")), errors="coerce").dropna()
    if close.shape[0] < 200:
        return (0, "Insufficient history.")
    hi = close.rolling(252).max().iloc[-1]
    p = close.iloc[-1]
    if hi in (None, 0):
        return (0, "52w high NA.")
    dd = (p / hi - 1.0) * 100.0
    if dd >= -20:
        return (2, f"Drawdown {dd:.1f}% recovered.")
    if dd >= -35:
        return (1, f"Drawdown {dd:.1f}% improving.")
    return (0, f"Drawdown {dd:.1f}% deep.")


# ==========================
# Weights + scoring
# ==========================
FUND_WEIGHTS = {
    "Margin stabilization (TTM vs prior)": 0.18,
    "Cashflow reversal (TTM FCF + quality)": 0.22,
    "Balance-sheet healing (net debt + IC)": 0.16,
    "ROIC health/inflection (proxy)": 0.10,
    "Capex discipline (intensity + FCF)": 0.08,
    "Revenue quality (Rev/share + GM)": 0.10,
    "Fundamentals vs valuation divergence": 0.16,
}

TECH_WEIGHTS = {
    "Price above MA200": 0.18,
    "MA structure improving": 0.16,
    "Higher low + breakout": 0.18,
    "RSI regime shift": 0.12,
    "Volume accumulation": 0.14,
    "Volatility regime": 0.10,
    "Drawdown recovery (52w)": 0.12,
}

COMBINED_SCORE_WEIGHTS = {
    "fundamental": 0.60,
    "technical": 0.40,
}


def _weighted_score(items: Dict[str, Tuple[int, str]], weights: Dict[str, float]) -> float:
    total = 0.0
    wsum = 0.0
    for k, w in weights.items():
        if k not in items:
            continue
        pts = items[k][0]
        total += (pts / 2.0) * w
        wsum += w
    return 0.0 if wsum == 0 else (total / wsum) * 100.0


def trend_reversal_scores(tkr: yf.Ticker, metrics: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    info = tkr.get_info()

    q_income = tkr.quarterly_income_stmt
    q_cf = tkr.quarterly_cashflow
    annual_bs = tkr.balance_sheet

    h_1y = tkr.history(period="1y", interval="1d", auto_adjust=False)
    h_2y = tkr.history(period="2y", interval="1d", auto_adjust=False)

    fundamental: Dict[str, Tuple[int, str]] = {}
    fundamental["Margin stabilization (TTM vs prior)"] = _fund_margin_stabilization(q_income)
    fundamental["Cashflow reversal (TTM FCF + quality)"] = _fund_cashflow_reversal(q_cf, q_income)
    fundamental["Balance-sheet healing (net debt + IC)"] = _fund_balance_sheet_healing(annual_bs, q_income)
    fundamental["ROIC health/inflection (proxy)"] = _fund_roic_inflection(metrics)
    fundamental["Capex discipline (intensity + FCF)"] = _fund_capex_discipline(q_cf, q_income)
    fundamental["Revenue quality (Rev/share + GM)"] = _fund_revenue_quality(q_income, info)
    fundamental["Fundamentals vs valuation divergence"] = _fund_value_divergence(metrics)

    technical: Dict[str, Tuple[int, str]] = {}
    technical["Price above MA200"] = _tech_price_above_200(h_2y)
    technical["MA structure improving"] = _tech_ma_structure(h_2y)
    technical["Higher low + breakout"] = _tech_breakout_higher_low(h_2y)
    technical["RSI regime shift"] = _tech_rsi(h_1y)
    technical["Volume accumulation"] = _tech_volume_accumulation(h_1y)
    technical["Volatility regime"] = _tech_volatility_regime(h_1y)
    technical["Drawdown recovery (52w)"] = _tech_drawdown_recovery(h_1y)

    f_score = _weighted_score(fundamental, FUND_WEIGHTS)
    t_score = _weighted_score(technical, TECH_WEIGHTS)
    combined_score = (
        f_score * COMBINED_SCORE_WEIGHTS.get("fundamental", 0.5)
        + t_score * COMBINED_SCORE_WEIGHTS.get("technical", 0.5)
    )

    f_sym = {k: _score_symbol(v[0]) for k, v in fundamental.items()}
    t_sym = {k: _score_symbol(v[0]) for k, v in technical.items()}

    counts = {
        "fund_green": sum(1 for v in f_sym.values() if v == "🟢"),
        "fund_gy": sum(1 for v in f_sym.values() if v in ("🟢", "🟡")),
        "tech_green": sum(1 for v in t_sym.values() if v == "🟢"),
        "tech_gy": sum(1 for v in t_sym.values() if v in ("🟢", "🟡")),
    }

    return {
        "fundamental": fundamental,
        "technical": technical,
        "fundamental_symbols": f_sym,
        "technical_symbols": t_sym,
        "fundamental_score": f_score,
        "technical_score": t_score,
        "combined_score": combined_score,
        "counts": counts,
    }
