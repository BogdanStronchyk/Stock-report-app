import math
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import yfinance as yf

from config import NUPL_REGIMES
from sector_map import map_sector


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
    if df is None or df.empty:
        return None
    for n in names:
        if n in df.index:
            return df.loc[n]
    return None


def twap(prices: pd.Series) -> Optional[float]:
    if prices is None or prices.empty:
        return None
    return float(prices.mean())


def normalize_cycle_proxy(current_price: float, ath: float) -> Optional[float]:
    if ath <= 0:
        return None
    return (current_price - ath) / ath


def nupl_regime(x: Optional[float]) -> str:
    if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
        return "NA"
    for lo, hi, name in NUPL_REGIMES:
        if x >= lo and x < hi:
            return name
    return "NA"


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
            return float(pd.to_numeric(s, errors="coerce").sum())
    return None


def realized_vol_1y(hist_1y: pd.DataFrame) -> Optional[float]:
    if hist_1y is None or hist_1y.empty:
        return None
    close = hist_1y["Close"].dropna()
    if len(close) < 50:
        return None
    ret = close.pct_change().dropna()
    return float(ret.std() * math.sqrt(252) * 100.0)


def worst_weekly_return_3y(hist_3y: pd.DataFrame) -> Optional[float]:
    if hist_3y is None or hist_3y.empty:
        return None
    close = hist_3y["Close"].dropna()
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


def compute_metrics_v2(ticker: str) -> Dict[str, Any]:
    tkr = yf.Ticker(ticker)
    info = tkr.get_info()

    sector = info.get("sector")
    industry = info.get("industry")
    sector_bucket = map_sector(sector, industry)

    h10 = tkr.history(period="10y", interval="1d", auto_adjust=False)
    h5  = tkr.history(period="5y", interval="1d", auto_adjust=False)
    h1y = tkr.history(period="1y", interval="1d", auto_adjust=False)
    h3y = tkr.history(period="3y", interval="1d", auto_adjust=False)

    price = float(h5["Close"].iloc[-1]) if (h5 is not None and not h5.empty) else safe_get(info, "currentPrice")

    twap10 = twap(h10["Close"]) if h10 is not None and not h10.empty else None
    twap5  = twap(h5["Close"]) if h5 is not None and not h5.empty else None
    nupl10 = (price - twap10) / price if (price and twap10) else None
    nupl5  = (price - twap5)  / price if (price and twap5)  else None
    ath = float(h10["High"].max()) if h10 is not None and not h10.empty else None
    cycle = normalize_cycle_proxy(price, ath) if (price and ath) else None

    composite_nupl = None
    if nupl10 is not None and nupl5 is not None and cycle is not None:
        composite_nupl = 0.40 * nupl10 + 0.35 * nupl5 + 0.25 * cycle

    annual_income = tkr.income_stmt
    annual_bs = tkr.balance_sheet
    annual_cf = tkr.cashflow

    trailing_pe = safe_get(info, "trailingPE")
    ev_ebitda = safe_get(info, "enterpriseToEbitda")
    beta = safe_get(info, "beta")
    short_pct = safe_get(info, "shortPercentOfFloat")
    short_ratio = safe_get(info, "shortRatio")
    ev = safe_get(info, "enterpriseValue")
    mcap = safe_get(info, "marketCap")

    q_income = tkr.quarterly_income_stmt
    q_cf = tkr.quarterly_cashflow
    ttm_rev = last_n_quarters_sum(q_income, ["Total Revenue", "TotalRevenue"], n=4)

    ttm_ocf = last_n_quarters_sum(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], n=4)
    ttm_capex = last_n_quarters_sum(q_cf, ["Capital Expenditure", "CapitalExpenditure"], n=4)
    ttm_fcf = (ttm_ocf + ttm_capex) if (ttm_ocf is not None and ttm_capex is not None) else None

    ps = safe_get(info, "priceToSalesTrailing12Months")
    if ps is None and mcap is not None and ttm_rev not in (None, 0):
        ps = float(mcap) / float(ttm_rev)

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

    ev_fcf = float(ev) / float(ttm_fcf) if (ev is not None and ttm_fcf not in (None, 0)) else None
    p_fcf = float(mcap) / float(ttm_fcf) if (mcap is not None and ttm_fcf not in (None, 0)) else None
    fcf_yield = float(ttm_fcf) / float(mcap) * 100.0 if (ttm_fcf is not None and mcap not in (None, 0)) else None

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

    ttm_ni = last_n_quarters_sum(q_income, ["Net Income", "NetIncome"], n=4)
    cfo_ni = float(ttm_ocf) / float(ttm_ni) * 100.0 if (ttm_ocf is not None and ttm_ni not in (None, 0)) else None

    ev_gross_profit = None
    if ev is not None and annual_income is not None and not annual_income.empty:
        gp = None
        for k in ["Gross Profit", "GrossProfit"]:
            if k in annual_income.index:
                gp = annual_income.loc[k].iloc[0]
                break
        if gp is not None and not pd.isna(gp) and float(gp) != 0:
            ev_gross_profit = float(ev) / float(gp)

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

    current_ratio = safe_get(info, "currentRatio")
    quick_ratio = safe_get(info, "quickRatio")

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

    max_dd = None
    if h3y is not None and not h3y.empty:
        close = h3y["Close"].dropna()
        if len(close) > 50:
            running_max = close.cummax()
            dd = (close / running_max - 1.0) * 100.0
            max_dd = float(dd.min())

    avg_dollar_vol = None
    try:
        h3m = tkr.history(period="3mo", interval="1d")
        if h3m is not None and not h3m.empty:
            dv = (h3m["Close"] * h3m["Volume"]).dropna()
            if len(dv) > 10:
                avg_dollar_vol = float(dv.mean())
    except Exception:
        pass

    vol_1y = realized_vol_1y(h1y)
    worst_week_3y = worst_weekly_return_3y(h3y)

    return {
        "Ticker": ticker,
        "Price": price,
        "Yahoo Sector": sector,
        "Yahoo Industry": industry,
        "Sector Bucket": sector_bucket,

        "P/E (TTM, positive EPS)": trailing_pe,
        "P/S": ps,
        "EV/EBITDA": ev_ebitda,
        "EV/EBIT": ev_ebit,
        "EV/FCF": ev_fcf,
        "Price/FCF": p_fcf,
        "FCF Yield (TTM FCF / Market Cap)": fcf_yield,
        "Earnings Yield (EBIT / EV)": earnings_yield,
        "EV/Gross Profit": ev_gross_profit,

        "Gross Margin %": gross_m_pct,
        "Operating Margin %": op_m_pct,
        "Net Margin %": net_m_pct,
        "ROE %": roe_pct,
        "ROIC % (standardized)": roic_pct,
        "FCF Margin % (TTM FCF / TTM Revenue)": fcf_margin,
        "CFO / Net Income": cfo_ni,

        "Net Debt / EBITDA": nd_ebitda,
        "Net Debt / FCF (years)": nd_fcf,
        "Interest Coverage (EBIT / Interest)": interest_cov,
        "FCF / Interest Expense": fcf_interest,
        "Current Ratio": current_ratio,
        "Quick Ratio": quick_ratio,
        "Cash / Total Assets": cash_assets,

        "Revenue CAGR (5Y)": rev_cagr,
        "Revenue per Share CAGR (5Y)": revps_cagr,
        "FCF per Share CAGR (5Y)": fcfps_cagr,
        "Dividend Growth (if applicable)": None,

        "Max Drawdown (3–5Y)": max_dd,
        "Realized Volatility (1Y, annualized)": vol_1y,
        "Beta (5Y)": beta,
        "Short Interest (% float)": (short_pct * 100.0) if short_pct is not None else None,
        "Days to Cover": short_ratio,
        "Avg Daily $ Volume (3M)": avg_dollar_vol,
        "Market Cap": mcap,
        "Worst Weekly Return (3Y)": worst_week_3y,

        "NUPL10Y": nupl10,
        "NUPL5Y": nupl5,
        "CycleProxy": cycle,
        "Composite NUPL": composite_nupl,
        "NUPL Regime": nupl_regime(composite_nupl),
    }
