
import math
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import yfinance as yf

from sector_map import map_sector


# -------------------------------------------------------------------
# Small helpers
# -------------------------------------------------------------------
def safe_get(d: Dict[str, Any], key: str) -> Optional[float]:
    v = d.get(key, None)
    try:
        if v is None:
            return None
        if isinstance(v, (int, float, np.floating)) and not pd.isna(v):
            return float(v)
        return None
    except Exception:
        return None


def get_row(df: pd.DataFrame, names: List[str]) -> Optional[pd.Series]:
    """Used by reversal.py. Keeps behavior stable."""
    if df is None or df.empty:
        return None
    for n in names:
        if n in df.index:
            return df.loc[n]
    return None


def _adj_close(hist: pd.DataFrame) -> pd.Series:
    """Prefer Adjusted Close (dividends+splits) for long-horizon work."""
    if hist is None or hist.empty:
        return pd.Series(dtype=float)
    if "Adj Close" in hist.columns:
        return hist["Adj Close"].astype(float)
    return hist["Close"].astype(float)


def twap(prices: pd.Series) -> Optional[float]:
    if prices is None or prices.empty:
        return None
    s = pd.to_numeric(prices, errors="coerce").dropna()
    if s.empty:
        return None
    return float(s.mean())


def annual_series(df: pd.DataFrame, row_names: List[str], n: int = 5) -> List[Optional[float]]:
    s = get_row(df, row_names)
    if s is None:
        return []
    vals = pd.to_numeric(s.iloc[:n][::-1], errors="coerce")  # oldest -> newest
    return [None if pd.isna(v) else float(v) for v in vals.values]


def cagr(values: List[Optional[float]]) -> Optional[float]:
    vals = [v for v in values if v is not None and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))]
    if len(vals) < 2:
        return None
    first, last = vals[0], vals[-1]
    n = len(vals) - 1
    if first <= 0 or last <= 0:
        return None
    return ((last / first) ** (1 / n) - 1) * 100.0


def last_n_quarters_sum(df: pd.DataFrame, row_names: List[str], n: int = 4) -> Optional[float]:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    cols = cols[:n] if len(cols) >= n else cols
    for rn in row_names:
        if rn in df.index:
            s = df.loc[rn, cols]
            try:
                return float(pd.to_numeric(s, errors="coerce").sum())
            except Exception:
                return None
    return None


def realized_vol_1y(hist_1y: pd.DataFrame) -> Optional[float]:
    """Annualized realized volatility from daily returns (%) using adjusted close."""
    if hist_1y is None or hist_1y.empty:
        return None
    close = _adj_close(hist_1y).dropna()
    if len(close) < 50:
        return None
    ret = close.pct_change().dropna()
    return float(ret.std() * math.sqrt(252) * 100.0)


def worst_weekly_return_3y(hist_3y: pd.DataFrame) -> Optional[float]:
    """Worst weekly return over last ~3y using adjusted close (%)."""
    if hist_3y is None or hist_3y.empty:
        return None
    close = _adj_close(hist_3y).dropna()
    if len(close) < 200:
        return None
    weekly = close.resample("W-FRI").last().pct_change().dropna()
    if weekly.empty:
        return None
    return float(weekly.min() * 100.0)


def approx_roic_percent(
    annual_income: pd.DataFrame,
    annual_bs: pd.DataFrame,
    tax_rate_assumption: float = 0.21
) -> Optional[float]:
    """Simple ROIC proxy: NOPAT / (Total Assets - Current Liabilities)."""
    if annual_income is None or annual_income.empty or annual_bs is None or annual_bs.empty:
        return None

    ebit = None
    for key in ["EBIT", "Ebit", "Operating Income", "OperatingIncome"]:
        if key in annual_income.index:
            ebit = annual_income.loc[key].iloc[0]
            break
    if ebit is None or pd.isna(ebit):
        return None

    total_assets = None
    curr_liab = None
    for key in ["Total Assets", "TotalAssets"]:
        if key in annual_bs.index:
            total_assets = annual_bs.loc[key].iloc[0]
            break
    for key in ["Current Liabilities", "CurrentLiabilities", "Total Current Liabilities"]:
        if key in annual_bs.index:
            curr_liab = annual_bs.loc[key].iloc[0]
            break
    if total_assets is None or curr_liab is None or pd.isna(total_assets) or pd.isna(curr_liab):
        return None

    invested = float(total_assets) - float(curr_liab)
    if invested <= 0:
        return None

    nopat = float(ebit) * (1 - tax_rate_assumption)
    return (nopat / invested) * 100.0


# -------------------------------------------------------------------
# Compound "Stock NUPL" v2 (improved proxy)
# -------------------------------------------------------------------
def anchored_vwap(price: pd.Series, volume: pd.Series) -> pd.Series:
    """Anchored VWAP: cumulative(price*vol)/cumulative(vol)."""
    p = pd.to_numeric(price, errors="coerce").ffill()
    v = pd.to_numeric(volume, errors="coerce").fillna(0.0)
    pv = (p * v).cumsum()
    vv = v.cumsum().replace(0, np.nan)
    return pv / vv


def turnover_realized_price(price: pd.Series, volume: pd.Series, shares_outstanding: Optional[float]) -> pd.Series:
    """
    Realized-price proxy via share turnover:
      f_t = min(0.25, volume/shares_outstanding)
      R_t = (1-f_t)R_{t-1} + f_t P_t

    If shares_outstanding is missing, we use a slow 1% daily refresh.
    """
    p = pd.to_numeric(price, errors="coerce").ffill()
    v = pd.to_numeric(volume, errors="coerce").fillna(0.0)

    if shares_outstanding is None or shares_outstanding <= 0:
        f = pd.Series(0.01, index=p.index)
    else:
        f = (v / float(shares_outstanding)).clip(lower=0.0, upper=0.25)

    r = pd.Series(index=p.index, dtype=float)
    if p.empty:
        return r

    r.iloc[0] = float(p.iloc[0])
    for i in range(1, len(p)):
        ft = float(f.iloc[i])
        r.iloc[i] = (1.0 - ft) * float(r.iloc[i - 1]) + ft * float(p.iloc[i])
    return r


def zscore_last(series: pd.Series, window: int = 2520) -> Optional[float]:
    """Last z-score vs trailing window (2520 ~ 10y trading days)."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < max(100, window // 4):
        return None
    tail = s.iloc[-window:] if len(s) >= window else s
    mu = float(tail.mean())
    sd = float(tail.std(ddof=0))
    if sd == 0 or math.isnan(sd):
        return None
    return float((tail.iloc[-1] - mu) / sd)


def stock_nupl_regime_from_z(z: Optional[float]) -> str:
    """Z-based regime mapping for stock-NUPL proxy."""
    if z is None or (isinstance(z, float) and (math.isnan(z) or math.isinf(z))):
        return "NA"
    if z <= -2.5:
        return "Deep Capitulation"
    if z <= -1.5:
        return "Capitulation"
    if z <= -0.5:
        return "Stress"
    if z < 0.5:
        return "Neutral"
    if z < 1.5:
        return "Optimism"
    return "Euphoria"


def normalize_cycle_proxy(current_price: float, ath: float) -> Optional[float]:
    if ath is None or ath <= 0:
        return None
    return (current_price - ath) / ath


# -------------------------------------------------------------------
# Main metrics computation
# -------------------------------------------------------------------
def compute_metrics_v2(ticker: str) -> Dict[str, Any]:
    """Compute checklist metrics + improved compound Stock NUPL v2."""
    tkr = yf.Ticker(ticker)
    info = tkr.get_info()

    sector = info.get("sector")
    industry = info.get("industry")
    sector_bucket = map_sector(sector, industry)

    # Price histories (keep Adj Close column)
    h10 = tkr.history(period="10y", interval="1d", auto_adjust=False)
    h5 = tkr.history(period="5y", interval="1d", auto_adjust=False)
    h1y = tkr.history(period="1y", interval="1d", auto_adjust=False)
    h3y = tkr.history(period="3y", interval="1d", auto_adjust=False)

    p10 = _adj_close(h10)
    p5 = _adj_close(h5)
    p1y = _adj_close(h1y)
    p3y = _adj_close(h3y)

    price = float(p5.iloc[-1]) if not p5.empty else safe_get(info, "currentPrice")

    # --- NUPL legacy (now based on Adj Close) ---
    twap10 = twap(p10)
    twap5 = twap(p5)
    nupl10 = (price - twap10) / price if (price and twap10) else None
    nupl5 = (price - twap5) / price if (price and twap5) else None

    ath_adj = float(p10.max()) if not p10.empty else None
    cycle_proxy = normalize_cycle_proxy(price, ath_adj) if (price and ath_adj) else None

    # --- NUPL v2: VWAP + turnover-realized price ---
    vwap10_series = anchored_vwap(p10, h10["Volume"]) if (h10 is not None and not h10.empty and "Volume" in h10.columns) else pd.Series(dtype=float)
    vwap10 = float(vwap10_series.iloc[-1]) if not vwap10_series.empty else None
    nupl_vwap10 = (price - vwap10) / price if (price and vwap10) else None

    shares_out = safe_get(info, "sharesOutstanding")
    realized_series = turnover_realized_price(p10, h10["Volume"], shares_out) if (h10 is not None and not h10.empty and "Volume" in h10.columns) else pd.Series(dtype=float)
    realized_price = float(realized_series.iloc[-1]) if not realized_series.empty else None
    nupl_turnover = (price - realized_price) / price if (price and realized_price) else None

    # Vol-scaled drawdown sigma (optional)
    vol_1y = realized_vol_1y(h1y)
    dd_sigma = None
    if cycle_proxy is not None and vol_1y not in (None, 0):
        dd_sigma = float(cycle_proxy) / (float(vol_1y) / 100.0)

    # Composite raw (kept near -1..+1)
    composite_nupl = None
    weights = {"turn": 0.50, "vwap": 0.30, "cycle": 0.20}
    vals = {"turn": nupl_turnover, "vwap": nupl_vwap10, "cycle": cycle_proxy}
    wsum = sum(weights[k] for k in weights if vals[k] is not None)
    if wsum > 0:
        composite_nupl = sum((weights[k] / wsum) * vals[k] for k in weights if vals[k] is not None)

    # Z-score composite (regime-stable)
    z_turn = zscore_last((p10 - realized_series) / p10) if (not p10.empty and not realized_series.empty) else None
    z_vwap = zscore_last((p10 - vwap10_series) / p10) if (not p10.empty and not vwap10_series.empty) else None
    z_cycle = None
    if ath_adj not in (None, 0) and not p10.empty:
        z_cycle = zscore_last((p10 - ath_adj) / ath_adj)

    z_weights = {"turn": 0.50, "vwap": 0.30, "cycle": 0.20}
    z_vals = {"turn": z_turn, "vwap": z_vwap, "cycle": z_cycle}
    zsum = sum(z_weights[k] for k in z_weights if z_vals[k] is not None)
    composite_z = None
    if zsum > 0:
        composite_z = sum((z_weights[k] / zsum) * z_vals[k] for k in z_weights if z_vals[k] is not None)

    nupl_regime = stock_nupl_regime_from_z(composite_z)

    # --- Statements ---
    annual_income = tkr.income_stmt
    annual_bs = tkr.balance_sheet
    annual_cf = tkr.cashflow

    # Valuation basics
    trailing_pe = safe_get(info, "trailingPE")
    ev_ebitda = safe_get(info, "enterpriseToEbitda")
    beta = safe_get(info, "beta")
    short_pct = safe_get(info, "shortPercentOfFloat")
    short_ratio = safe_get(info, "shortRatio")
    ev = safe_get(info, "enterpriseValue")
    mcap = safe_get(info, "marketCap")

    # TTM revenue & FCF
    q_income = tkr.quarterly_income_stmt
    q_cf = tkr.quarterly_cashflow

    ttm_rev = last_n_quarters_sum(q_income, ["Total Revenue", "TotalRevenue"], n=4)

    ttm_ocf = last_n_quarters_sum(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], n=4)
    ttm_capex = last_n_quarters_sum(q_cf, ["Capital Expenditure", "CapitalExpenditure"], n=4)
    ttm_fcf = (ttm_ocf + ttm_capex) if (ttm_ocf is not None and ttm_capex is not None) else None

    # P/S
    ps = safe_get(info, "priceToSalesTrailing12Months")
    if ps is None and mcap is not None and ttm_rev not in (None, 0):
        ps = float(mcap) / float(ttm_rev)

    # EV/EBIT
    ev_ebit = None
    if ev is not None and annual_income is not None and not annual_income.empty:
        ebit = None
        for k in ["EBIT", "Ebit", "Operating Income", "OperatingIncome"]:
            if k in annual_income.index:
                ebit = annual_income.loc[k].iloc[0]
                break
        if ebit is not None and not pd.isna(ebit) and float(ebit) != 0:
            ev_ebit = float(ev) / float(ebit)

    earnings_yield = (1.0 / ev_ebit) * 100.0 if ev_ebit not in (None, 0) else None

    # EV/FCF, Price/FCF, FCF Yield
    ev_fcf = float(ev) / float(ttm_fcf) if (ev is not None and ttm_fcf not in (None, 0)) else None
    p_fcf = float(mcap) / float(ttm_fcf) if (mcap is not None and ttm_fcf not in (None, 0)) else None
    fcf_yield = float(ttm_fcf) / float(mcap) * 100.0 if (ttm_fcf is not None and mcap not in (None, 0)) else None

    # Profitability
    gross_m = safe_get(info, "grossMargins")
    op_m = safe_get(info, "operatingMargins")
    net_m = safe_get(info, "profitMargins")
    roe = safe_get(info, "returnOnEquity")

    gross_m_pct = gross_m * 100 if gross_m is not None else None
    op_m_pct = op_m * 100 if op_m is not None else None
    net_m_pct = net_m * 100 if net_m is not None else None
    roe_pct = roe * 100 if roe is not None else None

    roic_pct = approx_roic_percent(annual_income, annual_bs)
    fcf_margin = (ttm_fcf / ttm_rev) * 100.0 if (ttm_fcf is not None and ttm_rev not in (None, 0)) else None

    # CFO/NI
    ttm_ni = last_n_quarters_sum(q_income, ["Net Income", "NetIncome"], n=4)
    cfo_ni = float(ttm_ocf) / float(ttm_ni) * 100.0 if (ttm_ocf is not None and ttm_ni not in (None, 0)) else None

    # EV/Gross Profit
    ev_gross_profit = None
    if ev is not None and annual_income is not None and not annual_income.empty:
        gp = None
        for k in ["Gross Profit", "GrossProfit"]:
            if k in annual_income.index:
                gp = annual_income.loc[k].iloc[0]
                break
        if gp is not None and not pd.isna(gp) and float(gp) != 0:
            ev_gross_profit = float(ev) / float(gp)

    # Balance sheet: net debt
    net_debt = None
    if annual_bs is not None and not annual_bs.empty:
        cash = None
        debt = None
        for k in ["Cash And Cash Equivalents", "CashAndCashEquivalents"]:
            if k in annual_bs.index:
                cash = annual_bs.loc[k].iloc[0]
                break
        for k in ["Total Debt", "TotalDebt", "Long Term Debt", "LongTermDebt"]:
            if k in annual_bs.index:
                debt = annual_bs.loc[k].iloc[0]
                break
        if cash is not None and debt is not None and not pd.isna(cash) and not pd.isna(debt):
            net_debt = float(debt) - float(cash)

    ebitda = safe_get(info, "ebitda")
    nd_ebitda = float(net_debt) / float(ebitda) if (net_debt is not None and ebitda not in (None, 0)) else None
    nd_fcf = float(net_debt) / float(ttm_fcf) if (net_debt is not None and ttm_fcf not in (None, 0)) else None

    # Interest coverage
    interest_cov = None
    fcf_interest = None
    if annual_income is not None and not annual_income.empty:
        ebit = None
        interest_exp = None
        for k in ["EBIT", "Ebit", "Operating Income", "OperatingIncome"]:
            if k in annual_income.index:
                ebit = annual_income.loc[k].iloc[0]
                break
        for k in ["Interest Expense", "InterestExpense"]:
            if k in annual_income.index:
                interest_exp = annual_income.loc[k].iloc[0]
                break
        if ebit is not None and interest_exp is not None and not pd.isna(ebit) and not pd.isna(interest_exp):
            denom = abs(float(interest_exp))
            if denom > 0:
                interest_cov = float(ebit) / denom
                fcf_interest = float(ttm_fcf) / denom if ttm_fcf is not None else None

    # Liquidity ratios
    current_ratio = safe_get(info, "currentRatio")
    quick_ratio = safe_get(info, "quickRatio")

    # Cash/Assets
    cash_assets = None
    if annual_bs is not None and not annual_bs.empty:
        cash = None
        assets = None
        for k in ["Cash And Cash Equivalents", "CashAndCashEquivalents"]:
            if k in annual_bs.index:
                cash = annual_bs.loc[k].iloc[0]
                break
        for k in ["Total Assets", "TotalAssets"]:
            if k in annual_bs.index:
                assets = annual_bs.loc[k].iloc[0]
                break
        if cash is not None and assets is not None and float(assets) != 0:
            cash_assets = float(cash) / float(assets) * 100.0

    # Growth
    revs = annual_series(annual_income, ["Total Revenue", "TotalRevenue"], 5)
    rev_cagr = cagr(revs)

    shares = safe_get(info, "sharesOutstanding")
    revps_cagr = None
    if shares and shares > 0 and len(revs) >= 2:
        revps = [r / shares if r is not None else None for r in revs]
        revps_cagr = cagr(revps)

    fcfps_cagr = None
    if shares and shares > 0 and annual_cf is not None and not annual_cf.empty:
        ocf_s = annual_series(annual_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 5)
        capex_s = annual_series(annual_cf, ["Capital Expenditure", "CapitalExpenditure"], 5)
        fcf_s = [(o + c) if (o is not None and c is not None) else None for o, c in zip(ocf_s, capex_s)]
        fcfps = [f / shares if f is not None else None for f in fcf_s]
        fcfps_cagr = cagr(fcfps)

    # Risk
    max_dd = None
    if not p3y.empty:
        running_max = p3y.cummax()
        dd = (p3y / running_max - 1.0) * 100.0
        max_dd = float(dd.min())

    avg_dollar_vol = None
    try:
        h3m = tkr.history(period="3mo", interval="1d", auto_adjust=False)
        if h3m is not None and not h3m.empty:
            px = _adj_close(h3m)
            dv = (px * h3m["Volume"]).dropna()
            if len(dv) > 10:
                avg_dollar_vol = float(dv.mean())
    except Exception:
        pass

    worst_week_3y = worst_weekly_return_3y(h3y)

    return {
        "Ticker": ticker,
        "Price": price,
        "Yahoo Sector": sector,
        "Yahoo Industry": industry,
        "Sector Bucket": sector_bucket,

        # Valuation (must match checklist sheet names)
        "P/E (TTM, positive EPS)": trailing_pe,
        "P/S": ps,
        "EV/EBITDA": ev_ebitda,
        "EV/EBIT": ev_ebit,
        "EV/FCF": ev_fcf,
        "Price/FCF": p_fcf,
        "FCF Yield (TTM FCF / Market Cap)": fcf_yield,
        "Earnings Yield (EBIT / EV)": earnings_yield,
        "EV/Gross Profit": ev_gross_profit,

        # Profitability
        "Gross Margin %": gross_m_pct,
        "Operating Margin %": op_m_pct,
        "Net Margin %": net_m_pct,
        "ROE %": roe_pct,
        "ROIC % (standardized)": roic_pct,
        "FCF Margin % (TTM FCF / TTM Revenue)": fcf_margin,
        "CFO / Net Income": cfo_ni,

        # Balance sheet
        "Net Debt / EBITDA": nd_ebitda,
        "Net Debt / FCF (years)": nd_fcf,
        "Interest Coverage (EBIT / Interest)": interest_cov,
        "FCF / Interest Expense": fcf_interest,
        "Current Ratio": current_ratio,
        "Quick Ratio": quick_ratio,
        "Cash / Total Assets": cash_assets,

        # Growth
        "Revenue CAGR (5Y)": rev_cagr,
        "Revenue per Share CAGR (5Y)": revps_cagr,
        "FCF per Share CAGR (5Y)": fcfps_cagr,
        "Dividend Growth (if applicable)": None,

        # Risk
        "Max Drawdown (3–5Y)": max_dd,
        "Realized Volatility (1Y, annualized)": vol_1y,
        "Beta (5Y)": beta,
        "Short Interest (% float)": (short_pct * 100.0) if short_pct is not None else None,
        "Days to Cover": short_ratio,
        "Avg Daily $ Volume (3M)": avg_dollar_vol,
        "Market Cap": mcap,
        "Worst Weekly Return (3Y)": worst_week_3y,

        # NUPL legacy keys (kept for compatibility)
        "NUPL10Y": nupl10,
        "NUPL5Y": nupl5,
        "CycleProxy": cycle_proxy,

        # NUPL v2 components (new)
        "NUPL VWAP10 Proxy": nupl_vwap10,
        "NUPL Turnover Proxy": nupl_turnover,
        "Realized Price (Turnover Proxy)": realized_price,
        "VWAP10 (Anchored)": vwap10,
        "Drawdown Sigma (ATH/Vol)": dd_sigma,

        # Composite (report uses these keys)
        "Composite NUPL": composite_nupl,
        "Composite NUPL Z": composite_z,
        "NUPL Regime": nupl_regime,
    }
