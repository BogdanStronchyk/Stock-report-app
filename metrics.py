import math
import time
from typing import Any, Dict, List, Optional
import pandas as pd
import yfinance as yf
from cache_utils import DiskCache, yf_call, yf_cache_settings
from config import FORCE_FMP_FALLBACK
from value_matrix_extras import compute_value_matrix_extras
from sector_map import map_sector
from fmp_provider import FMPClient

_MULTIPLIERS = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}


def ensure_df(val: Any) -> pd.DataFrame:
    return val if val is not None else pd.DataFrame()


def _to_float(value: Any) -> Optional[float]:
    if value is None: return None
    try:
        v = float(value)
        if math.isnan(v) or math.isinf(v): return None
        return v
    except:
        return None


def safe_get(d: Dict[str, Any], key: str) -> Optional[float]:
    return _to_float(d.get(key))


def _fmp_get_num(bundle: Dict[str, Any], section: str, *keys: str) -> Optional[float]:
    rec = bundle.get(section)
    if isinstance(rec, list) and rec: rec = rec[0]
    if not isinstance(rec, dict): return None
    for k in keys:
        v = _to_float(rec.get(k))
        if v is not None: return v
    return None


def _download_history(symbol: str, period: str) -> pd.DataFrame:
    try:
        return yf_call(
            lambda: yf.download(symbol, period=period, interval="1d", auto_adjust=False, progress=False, threads=False))
    except:
        return pd.DataFrame()


def _history_retry(symbol: str, tkr: yf.Ticker, period: str) -> pd.DataFrame:
    for _ in range(2):
        try:
            h = yf_call(lambda: tkr.history(period=period, interval="1d", auto_adjust=False))
            if h is not None and not h.empty: return h
        except:
            time.sleep(0.5)
    return _download_history(symbol, period)


def annual_series(df: pd.DataFrame, row_names: List[str], n: int = 5) -> List[Optional[float]]:
    if df is None or df.empty: return []
    for r in row_names:
        if r in df.index:
            vals = pd.to_numeric(df.loc[r].iloc[:n][::-1], errors="coerce")
            return [float(v) if not pd.isna(v) else None for v in vals]
    return []


def cagr(values: List[Optional[float]]) -> Optional[float]:
    vals = [v for v in values if v is not None and v > 0]
    if len(vals) < 2: return None
    return ((vals[-1] / vals[0]) ** (1 / (len(vals) - 1)) - 1) * 100.0


def last_n_quarters_sum(df: pd.DataFrame, row_names: List[str]) -> Optional[float]:
    if df is None or df.empty: return None
    cols = list(df.columns)[:4]
    for rn in row_names:
        if rn in df.index:
            try:
                return float(pd.to_numeric(df.loc[rn, cols], errors="coerce").sum())
            except:
                pass
    return None


def approx_roic_percent(annual_income: pd.DataFrame, annual_bs: pd.DataFrame) -> Optional[float]:
    if annual_income.empty or annual_bs.empty: return None
    ebit = None
    for k in ["EBIT", "Ebit", "Operating Income", "OperatingIncome"]:
        if k in annual_income.index: ebit = _to_float(annual_income.loc[k].iloc[0]); break
    total_assets = None
    for k in ["Total Assets", "TotalAssets"]:
        if k in annual_bs.index: total_assets = _to_float(annual_bs.loc[k].iloc[0]); break
    curr_liab = None
    for k in ["Current Liabilities", "CurrentLiabilities", "Total Current Liabilities"]:
        if k in annual_bs.index: curr_liab = _to_float(annual_bs.loc[k].iloc[0]); break

    if ebit is None or total_assets is None or curr_liab is None: return None
    invested = float(total_assets) - float(curr_liab)
    if invested <= 0: return None
    return (float(ebit) * 0.79 / invested) * 100.0


def compute_metrics_v2(ticker: str, use_fmp_fallback: bool = True, *, fmp_mode: str = "full",
                       use_yf_cache: Optional[bool] = None) -> Dict[str, Any]:
    tkr = yf.Ticker(ticker)
    _cache_enabled, _cache_ttl = yf_cache_settings()
    if use_yf_cache is not None: _cache_enabled = bool(use_yf_cache)
    yf_cache = DiskCache("yf", ttl_hours=_cache_ttl, enabled=_cache_enabled)

    def _cached_df(key: str, fn):
        v = yf_cache.get_pickle(key)
        if v is not None: return v
        try:
            v2 = fn()
        except:
            v2 = None
        if v2 is not None: yf_cache.set_pickle(key, v2)
        return v2

    info = yf_cache.get_json(f"info:{ticker}") or yf_call(lambda: tkr.get_info() or {}) or {}
    yf_cache.set_json(f"info:{ticker}", info)

    fmp_data = {}
    if (use_fmp_fallback or FORCE_FMP_FALLBACK):
        fmp = FMPClient()
        if fmp.enabled: fmp_data = fmp.fetch_bundle(ticker, mode=fmp_mode) or {}

    h10 = ensure_df(_cached_df(f"hist:{ticker}:10y:1d", lambda: _history_retry(ticker, tkr, period="10y")))
    try:
        h3y = h10.loc[h10.index >= (h10.index.max() - pd.DateOffset(years=3))].copy()
    except:
        h3y = pd.DataFrame()

    price = safe_get(info, "currentPrice") or _fmp_get_num(fmp_data, "quote", "price")
    if price is None and not h10.empty: price = float(h10["Close"].iloc[-1])

    mcap = safe_get(info, "marketCap") or _fmp_get_num(fmp_data, "quote", "marketCap")
    ev = safe_get(info, "enterpriseValue") or _fmp_get_num(fmp_data, "key_metrics_ttm", "enterpriseValue")

    annual_income = ensure_df(_cached_df(f"stmt:{ticker}:income_stmt", lambda: yf_call(lambda: tkr.income_stmt)))
    annual_bs = ensure_df(_cached_df(f"stmt:{ticker}:balance_sheet", lambda: yf_call(lambda: tkr.balance_sheet)))
    annual_cf = ensure_df(_cached_df(f"stmt:{ticker}:cashflow", lambda: yf_call(lambda: tkr.cashflow)))
    q_income = ensure_df(_cached_df(f"stmt:{ticker}:q_income", lambda: yf_call(lambda: tkr.quarterly_income_stmt)))
    q_cf = ensure_df(_cached_df(f"stmt:{ticker}:q_cf", lambda: yf_call(lambda: tkr.quarterly_cashflow)))

    ttm_rev = last_n_quarters_sum(q_income, ["Total Revenue", "TotalRevenue"])
    ttm_ebit = last_n_quarters_sum(q_income, ["Operating Income", "OperatingIncome", "EBIT"])
    ttm_ni = last_n_quarters_sum(q_income, ["Net Income", "NetIncome"])
    ttm_gp = last_n_quarters_sum(q_income, ["Gross Profit", "GrossProfit"])
    ttm_fcf = None
    ocf = last_n_quarters_sum(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"])
    cap = last_n_quarters_sum(q_cf, ["Capital Expenditure", "CapitalExpenditure"])
    if ocf is not None and cap is not None: ttm_fcf = ocf + cap

    pe = safe_get(info, "trailingPE")
    if pe is None and mcap and ttm_ni and ttm_ni > 0: pe = mcap / ttm_ni

    ev_ebit = (ev / ttm_ebit) if (ev and ttm_ebit and ttm_ebit > 0) else None
    fcf_yield = (ttm_fcf / mcap * 100) if (ttm_fcf and mcap) else None
    gross_m = (ttm_gp / ttm_rev * 100) if (ttm_gp and ttm_rev) else None
    op_m = (ttm_ebit / ttm_rev * 100) if (ttm_ebit and ttm_rev) else None
    roic = approx_roic_percent(annual_income, annual_bs)

    net_debt = None
    if not annual_bs.empty:
        cash = None;
        debt = None
        for k in ["Cash And Cash Equivalents", "CashAndCashEquivalents"]:
            if k in annual_bs.index: cash = _to_float(annual_bs.loc[k].iloc[0]); break
        for k in ["Total Debt", "TotalDebt", "Long Term Debt"]:
            if k in annual_bs.index: debt = _to_float(annual_bs.loc[k].iloc[0]); break
        if cash is not None and debt is not None: net_debt = debt - cash

    nd_ebitda = None
    ebitda = safe_get(info, "ebitda") or (ttm_ebit * 1.15 if ttm_ebit else None)
    if net_debt is not None and ebitda and ebitda > 0: nd_ebitda = net_debt / ebitda

    int_cov = None
    int_exp = None
    for k in ["Interest Expense", "InterestExpense"]:
        if k in annual_income.index: int_exp = _to_float(annual_income.loc[k].iloc[0]); break
    if int_exp and ttm_ebit: int_cov = ttm_ebit / abs(int_exp)

    shares = safe_get(info, "sharesOutstanding")
    revps_cagr = None
    fcfps_cagr = None
    if shares:
        revs = annual_series(annual_income, ["Total Revenue", "TotalRevenue"], 5)
        if len(revs) > 1: revps_cagr = cagr([r / shares if r else None for r in revs])

        ocfs = annual_series(annual_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], 5)
        caps = annual_series(annual_cf, ["Capital Expenditure", "CapitalExpenditure"], 5)
        if len(ocfs) == len(caps) and len(ocfs) > 1:
            fcfs = [(o + c) for o, c in zip(ocfs, caps) if o is not None and c is not None]
            fcfps_cagr = cagr([f / shares for f in fcfs])

    max_dd = None
    if not h3y.empty:
        c = h3y["Close"]
        max_dd = ((c / c.cummax() - 1).min()) * 100

    metrics = {
        "Ticker": ticker, "Price": price, "Market Cap": mcap,
        "Sector Bucket": map_sector(info.get("sector"), info.get("industry")),
        "P/E (TTM, positive EPS)": pe, "EV/EBIT": ev_ebit, "FCF Yield (TTM FCF / Market Cap)": fcf_yield,
        "Gross Margin %": gross_m, "Operating Margin %": op_m, "ROIC % (standardized)": roic,
        "Net Debt / EBITDA": nd_ebitda, "Interest Coverage (EBIT / Interest)": int_cov,
        "Revenue per Share CAGR (5Y)": revps_cagr, "FCF per Share CAGR (5Y)": fcfps_cagr,
        "Max Drawdown (3–5Y)": max_dd,
        "__yf_bundle__": {"info": info, "h10": h10}, "__notes__": {}
    }
    return metrics