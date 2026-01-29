import math
import os
import re
import time
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import yfinance as yf

from cache_utils import DiskCache, yf_call, yf_cache_settings

from config import FORCE_FMP_FALLBACK
from value_matrix_extras import compute_value_matrix_extras

from sector_map import map_sector
from fmp_provider import FMPClient


# -------------------------------------------------------------------
# Small helpers
# -------------------------------------------------------------------
_MULTIPLIERS = {
    "K": 1e3,
    "M": 1e6,
    "B": 1e9,
    "T": 1e12,
}

def davf_protection_label(mos_pct: Optional[float], confidence: str) -> str:
    """
    Interpretation label for DAVF MOS vs Floor (%).
    Conservative by design:
      - If LOW confidence → return "NA" (forces manual review)
      - Else:
          >= 40%  -> "GREEN"
          15-40%  -> "YELLOW"
          < 15%   -> "RED"
    """
    if mos_pct is None:
        return "NA"
    if (confidence or "NA").upper() == "LOW":
        return "NA"

    try:
        m = float(mos_pct)
    except Exception:
        return "NA"

    if m >= 40.0:
        return "GREEN"
    if m >= 15.0:
        return "YELLOW"
    return "RED"


# -------------------------------------------------------------------
# DAVF (Downside-Anchored Value Floor)
# -------------------------------------------------------------------
# Conservative value floor (not a fair-value model):
#   - uses normalized historical cashflows/earnings (no forecasts)
#   - applies only downside haircuts (durability / fragility proxies)
#   - subtracts net debt (or adds net cash) to get an equity floor
# Output:
#   DAVF Value Floor (per share), DAVF MOS % (vs floor), confidence + diagnostics

SECTOR_FLOOR_MULTIPLES = {
    # keep these conservative and stable; tune later if needed
    "Defensive": 10.0,
    "Cyclical": 6.0,
    "Capital Intensive": 5.0,
    "Asset Heavy": 0.8,     # intended for P/B style fallback; we won't use unless needed
    "High Growth": 8.0,
    "Default (All)": 7.0,
}

def _winsorize(vals, p_lo=0.10, p_hi=0.90):
    xs = [float(v) for v in vals if v is not None and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))]
    if len(xs) < 3:
        return xs
    lo = float(np.quantile(xs, p_lo))
    hi = float(np.quantile(xs, p_hi))
    return [min(max(x, lo), hi) for x in xs]

def _median_pos(vals):
    xs = [float(v) for v in vals if v is not None and float(v) > 0 and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))]
    if not xs:
        return None
    return float(np.median(xs))

def _fcf_series_annual(annual_cf: pd.DataFrame, n: int = 8):
    """Oldest->newest annual FCF series using OCF + CapEx (CapEx is usually negative)."""
    if annual_cf is None or annual_cf.empty:
        return []
    ocf_s = annual_series(annual_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], n)
    cap_s = annual_series(annual_cf, ["Capital Expenditure", "CapitalExpenditure"], n)
    if not ocf_s or not cap_s:
        return []
    out = []
    for o, c in zip(ocf_s, cap_s):
        out.append((o + c) if (o is not None and c is not None) else None)
    return out

def _ebit_series_annual(annual_income: pd.DataFrame, n: int = 8):
    if annual_income is None or annual_income.empty:
        return []
    return annual_series(annual_income, ["EBIT", "Ebit", "Operating Income", "OperatingIncome"], n)

def _ni_series_annual(annual_income: pd.DataFrame, n: int = 8):
    if annual_income is None or annual_income.empty:
        return []
    return annual_series(annual_income, ["Net Income", "NetIncome"], n)

def _interest_series_annual(annual_income: pd.DataFrame, n: int = 8):
    if annual_income is None or annual_income.empty:
        return []
    return annual_series(annual_income, ["Interest Expense", "InterestExpense"], n)

def _net_debt_latest(annual_bs: pd.DataFrame) -> Optional[float]:
    if annual_bs is None or annual_bs.empty:
        return None

    cash = None
    debt = None
    for k in ["Cash And Cash Equivalents", "CashAndCashEquivalents"]:
        if k in annual_bs.index:
            cash = _to_float(annual_bs.loc[k].iloc[0])
            break
    for k in ["Total Debt", "TotalDebt", "Long Term Debt", "LongTermDebt"]:
        if k in annual_bs.index:
            debt = _to_float(annual_bs.loc[k].iloc[0])
            break

    if cash is None or debt is None:
        return None
    return float(debt) - float(cash)

def compute_davf(
    sector_bucket: str,
    price: Optional[float],
    shares_out: Optional[float],
    annual_income: pd.DataFrame,
    annual_cf: pd.DataFrame,
    annual_bs: pd.DataFrame,
    roic_pct: Optional[float],
    nd_ebitda: Optional[float],
    interest_cov: Optional[float],
) -> Dict[str, Any]:
    """
    Returns a dict with:
      - davf_floor_ps
      - davf_mos_pct
      - davf_confidence
      - davf_base_type
      - davf_multiple
      - davf_haircut_pct
      - davf_note
    """
    out: Dict[str, Any] = {
        "davf_floor_ps": None,
        "davf_mos_pct": None,
        "davf_confidence": "NA",
        "davf_base_type": None,
        "davf_multiple": None,
        "davf_haircut_pct": None,
        "davf_note": "",
        "davf_label": "NA",

    }

    if price in (None, 0) or shares_out in (None, 0):
        out["davf_note"] = "Missing price or sharesOutstanding."
        return out

    mult = float(SECTOR_FLOOR_MULTIPLES.get(sector_bucket, SECTOR_FLOOR_MULTIPLES["Default (All)"]))
    out["davf_multiple"] = mult

    # --- Choose normalized base: FCF (preferred) → EBIT → Net Income ---
    fcf_series = _fcf_series_annual(annual_cf, n=8)
    ebit_series = _ebit_series_annual(annual_income, n=8)
    ni_series = _ni_series_annual(annual_income, n=8)

    # Use winsorized, positive medians
    fcf_norm = _median_pos(_winsorize(fcf_series))
    ebit_norm = _median_pos(_winsorize(ebit_series))
    ni_norm = _median_pos(_winsorize(ni_series))

    base = None
    base_type = None
    years_used = 0

    if fcf_norm is not None:
        base = fcf_norm
        base_type = "FCF (median, winsorized)"
        years_used = len([v for v in fcf_series if v is not None])
    elif ebit_norm is not None:
        base = ebit_norm
        base_type = "EBIT (median, winsorized)"
        years_used = len([v for v in ebit_series if v is not None])
    elif ni_norm is not None:
        base = ni_norm
        base_type = "Net Income (median, winsorized)"
        years_used = len([v for v in ni_series if v is not None])
    else:
        out["davf_note"] = "No usable positive history for FCF/EBIT/Net Income."
        return out

    out["davf_base_type"] = base_type

    # Confidence from coverage depth
    if years_used >= 5:
        out["davf_confidence"] = "HIGH"
    elif years_used >= 3:
        out["davf_confidence"] = "MEDIUM"
    else:
        out["davf_confidence"] = "LOW"

    # --- Haircuts (durability / fragility proxies) ---
    haircut = 0.0
    hc_reasons = []

    # Leverage proxy
    if nd_ebitda is not None and nd_ebitda > 3:
        haircut += 0.20
        hc_reasons.append("Net Debt/EBITDA > 3")

    # Coverage proxy
    if interest_cov is not None and interest_cov < 3:
        haircut += 0.20
        hc_reasons.append("Interest coverage < 3")

    # FCF volatility proxy (only if we used FCF and have enough years)
    if base_type.startswith("FCF") and len([v for v in fcf_series if v is not None]) >= 4:
        xs = [float(v) for v in fcf_series if v is not None]
        xs = _winsorize(xs)
        if len(xs) >= 4:
            mu = float(np.mean(xs))
            sig = float(np.std(xs))
            if mu != 0 and abs(sig / mu) > 0.75:
                haircut += 0.15
                hc_reasons.append("FCF volatility high")

    # ROIC proxy (if available)
    if roic_pct is not None and roic_pct < 6:
        haircut += 0.15
        hc_reasons.append("ROIC weak (<6%)")

    haircut = min(haircut, 0.50)  # cap at 50%
    out["davf_haircut_pct"] = haircut * 100.0

    # --- Equity floor calculation ---
    # Enterprise-like floor from base × multiple, then subtract net debt (add net cash if negative net debt).
    raw_value = float(base) * mult

    nd = _net_debt_latest(annual_bs)
    if nd is not None:
        raw_value = raw_value - float(nd)

    adj_value = raw_value * (1.0 - haircut)

    # Per share floor
    floor_ps = adj_value / float(shares_out)
    out["davf_floor_ps"] = float(floor_ps)

    mos = (float(floor_ps) / float(price) - 1.0) * 100.0
    out["davf_mos_pct"] = float(mos)
    out["davf_label"] = davf_protection_label(out["davf_mos_pct"], out["davf_confidence"])

    # Note
    note_bits = [f"Base={base_type}", f"Mult={mult:.1f}"]
    if nd is not None:
        note_bits.append(f"NetDebtAdj={'yes' if nd > 0 else 'net cash'}")
    if haircut > 0:
        note_bits.append(f"Haircut={haircut*100:.0f}% ({', '.join(hc_reasons)})")
    out["davf_note"] = " | ".join(note_bits)

    return out


def _to_float(value: Any) -> Optional[float]:
    """
    Robust numeric coercion:
      - accepts Python int/float
      - accepts numpy numbers (np.integer + np.floating)
      - accepts strings like '10B', '$250M', '1,234,567', '(123.4)', '15.2%'
    Returns None if it can't parse or is NaN/Inf.
    """
    if value is None:
        return None

    # Fast path: numeric types (including numpy)
    try:
        if isinstance(value, (int, float, np.number)):
            v = float(value)
            if math.isnan(v) or math.isinf(v):
                return None
            return v
    except Exception:
        pass

    # Strings
    if isinstance(value, str):
        s = value.strip()
        if not s or s.lower() in {"n/a", "na", "none", "null", "-"}:
            return None

        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1].strip()

        s = s.replace(" ", "")

        # Strip leading currency/symbols
        s = re.sub(r"^[^\d\.\-\+]+", "", s)

        # Percent
        is_percent = s.endswith("%")
        if is_percent:
            s = s[:-1]

        # Remove thousands separators
        s = s.replace(",", "").replace("_", "")

        # Optional suffix multiplier
        m = re.match(r"^([+-]?\d*\.?\d+)([KMBT])?$", s, re.IGNORECASE)
        if not m:
            return None

        try:
            num = float(m.group(1))
        except Exception:
            return None

        suffix = m.group(2)
        if suffix:
            num *= _MULTIPLIERS.get(suffix.upper(), 1.0)

        if neg:
            num = -num
        if is_percent:
            num /= 100.0

        if math.isnan(num) or math.isinf(num):
            return None
        return num

    return None


def safe_get(d: Dict[str, Any], key: str) -> Optional[float]:
    """
    Extract numeric value from a dict safely.

    IMPORTANT FIX:
    Previous version rejected np.integer (e.g., numpy.int64) which yfinance frequently
    returns for marketCap, enterpriseValue, sharesOutstanding, etc. That cascades into
    lots of N/As.

    This version:
      - accepts any numpy number via np.number
      - parses compact strings like '10B' if they appear (fallback sources)
    """
    try:
        v = d.get(key, None)
        out = _to_float(v)
        if out is None:
            return None
        if pd.isna(out):
            return None
        return float(out)
    except Exception:
        return None


def _fmp_first_record(payload: Any) -> Optional[Dict[str, Any]]:
    """FMP endpoints sometimes return list[dict]; sometimes dict."""
    if payload is None:
        return None
    if isinstance(payload, dict):
        return payload
    if isinstance(payload, list) and payload:
        if isinstance(payload[0], dict):
            return payload[0]
    return None


def _fmp_get_num(bundle: Dict[str, Any], section: str, *keys: str) -> Optional[float]:
    """
    Extract a numeric value from an FMP bundle section.
    Example: _fmp_get_num(fmp_data, "quote", "price") or ("marketCap").
    """
    rec = _fmp_first_record(bundle.get(section))
    if not rec:
        return None
    for k in keys:
        if k in rec:
            v = _to_float(rec.get(k))
            if v is not None:
                return v
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



def _fast_info_dict(tkr: yf.Ticker) -> Dict[str, Any]:
    """Best-effort fast_info dict extraction (yfinance versions vary)."""
    try:
        fi = getattr(tkr, "fast_info", None)
        if fi is None:
            return {}
        if isinstance(fi, dict):
            return fi
        return dict(fi)
    except Exception:
        return {}


def _download_history(symbol: str, period: str, interval: str = "1d") -> pd.DataFrame:
    """Fallback history fetch using yf.download (sometimes more stable than Ticker.history)."""
    try:
        return yf_call(lambda: yf.download(
            symbol,
            period=period,
            interval=interval,
            auto_adjust=False,
            progress=False,
            threads=False,
        ))
    except Exception:
        return pd.DataFrame()


def _history_retry(symbol: str, tkr: yf.Ticker, period: str, interval: str = "1d", tries: int = 3) -> pd.DataFrame:
    """Fetch history with small retries + yf.download fallback."""
    last_err = None
    for i in range(max(1, tries)):
        try:
            h = yf_call(lambda: tkr.history(period=period, interval=interval, auto_adjust=False))
            if h is not None and not h.empty:
                return h
        except Exception as e:
            last_err = e
        # small backoff (kept short to avoid UI feeling frozen)
        try:
            time.sleep(0.8 * (i + 1))
        except Exception:
            pass

    h2 = _download_history(symbol, period, interval)
    if h2 is not None and not h2.empty:
        return h2
    return pd.DataFrame()


# -------------------------------------------------------------------
# History slicing helpers (PERF)
# -------------------------------------------------------------------
def _slice_history_from(hist: pd.DataFrame, years: float) -> pd.DataFrame:
    """Slice a daily history df to last N years (no extra API calls)."""
    if hist is None or hist.empty:
        return pd.DataFrame()
    try:
        end = hist.index.max()
        cutoff = end - pd.DateOffset(days=int(years * 365.25))
        return hist.loc[hist.index >= cutoff].copy()
    except Exception:
        return hist


def _slice_history_days(hist: pd.DataFrame, days: int) -> pd.DataFrame:
    """Slice a daily history df to last N calendar days."""
    if hist is None or hist.empty:
        return pd.DataFrame()
    try:
        end = hist.index.max()
        cutoff = end - pd.Timedelta(days=int(days))
        return hist.loc[hist.index >= cutoff].copy()
    except Exception:
        return hist


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
def compute_metrics_v2(
    ticker: str,
    use_fmp_fallback: bool = True,
    *,
    fmp_mode: str = "full",
    use_yf_cache: Optional[bool] = None,
) -> Dict[str, Any]:
    """Compute checklist metrics + improved compound Stock NUPL v2.

    Adds a special key '__notes__' which maps metric name -> note string.
    report_writer.py will display these notes in the Notes column so that
    'not meaningful' ratios are explicitly explained.
    """
    tkr = yf.Ticker(ticker)


    # --- Yahoo cache + concurrency cap ---
    _cache_enabled, _cache_ttl = yf_cache_settings()
    if use_yf_cache is not None:
        _cache_enabled = bool(use_yf_cache)
    yf_cache = DiskCache("yf", ttl_hours=_cache_ttl, enabled=_cache_enabled)

    def _cached_json(key: str, fn):
        v = yf_cache.get_json(key)
        if v is not None:
            return v
        try:
            v2 = fn()
        except Exception:
            v2 = None
        if v2 is not None:
            yf_cache.set_json(key, v2)
        return v2

    def _cached_df(key: str, fn):
        v = yf_cache.get_pickle(key)
        if v is not None:
            return v
        try:
            v2 = fn()
        except Exception:
            v2 = None
        if v2 is not None:
            yf_cache.set_pickle(key, v2)
        return v2

    info = _cached_json(f"info:{ticker}", lambda: yf_call(lambda: tkr.get_info() or {})) or {}
    fast_info = _fast_info_dict(tkr)

    notes: Dict[str, str] = {}

    # --- Optional FMP fallback bundle ---
    # FMP is a *fallback* provider for missing/misaligned Yahoo fields (shares, EV, margins, etc).
    # If FORCE_FMP_FALLBACK=True and FMP_API_KEY is present, we use it even if the UI toggle is off.
    fmp_data: Dict[str, Any] = {}
    fmp = FMPClient()

    requested = bool(use_fmp_fallback)
    forced = bool(FORCE_FMP_FALLBACK)
    use_fmp_effective = requested or forced

    env_mode = (os.environ.get("FMP_MODE") or "").strip().lower()
    fmp_mode_effective = (env_mode or (fmp_mode or "full")).strip().lower()

    if not use_fmp_effective:
        notes["FMP"] = "enabled=no (disabled by user)"
    elif fmp.enabled:
        # Partial FMP to control time/budget per stage (conditional|minimal|full)
        fmp_data = fmp.fetch_bundle(ticker, mode=fmp_mode_effective) or {}
        if forced and not requested:
            notes["FMP"] = f"enabled=yes (FORCED, mode={fmp_mode_effective}); requests={fmp.request_count}; last_status={fmp.last_status}; last_error={fmp.last_error}"
        else:
            notes["FMP"] = f"enabled=yes (mode={fmp_mode_effective}); requests={fmp.request_count}; last_status={fmp.last_status}; last_error={fmp.last_error}"
    else:
        notes["FMP"] = "enabled=no (set FMP_API_KEY env var to enable fallback)"

    sector = info.get("sector")
    industry = info.get("industry")
    sector_bucket = map_sector(sector, industry)

    # Price histories
    # PERF: one network call (10y) + local slicing instead of 4 separate calls.
    h10 = _cached_df(f"hist:{ticker}:10y:1d", lambda: _history_retry(ticker, tkr, period="10y", interval="1d"))
    if h10 is None:
        h10 = pd.DataFrame()
    h5 = _slice_history_from(h10, 5.0)
    h3y = _slice_history_from(h10, 3.0)
    h2y = _slice_history_from(h10, 2.0)
    h1y = _slice_history_from(h10, 1.0)

    p10 = _adj_close(h10)
    p5 = _adj_close(h5)
    p1y = _adj_close(h1y)
    p3y = _adj_close(h3y)

    # Price: prefer history -> fast_info -> Yahoo info -> FMP quote
    fi_price = _to_float(fast_info.get("last_price") or fast_info.get("lastPrice") or fast_info.get("regularMarketPrice"))
    price = float(p5.iloc[-1]) if not p5.empty else fi_price
    if price is None:
        price = safe_get(info, "currentPrice")
    if price is None:
        price = safe_get(info, "regularMarketPrice")
    if price is None:
        price = _fmp_get_num(fmp_data, "quote", "price")

    # Market cap / EV early (needed for share fallback + yield metrics)
    mcap = safe_get(info, "marketCap")
    if mcap is None:
        mcap = _to_float(fast_info.get("market_cap") or fast_info.get("marketCap"))
    if mcap is None:
        mcap = _fmp_get_num(fmp_data, "quote", "marketCap")

    ev = safe_get(info, "enterpriseValue")
    if ev is None:
        ev = _fmp_get_num(fmp_data, "profile", "enterpriseValue")
    if ev is None:
        ev = _fmp_get_num(fmp_data, "key_metrics_ttm", "enterpriseValue")
    if ev is None:
        ev = _fmp_get_num(fmp_data, "quote", "enterpriseValue")
    if ev is None:
        ev = mcap

    shares_out = safe_get(info, "sharesOutstanding")
    if shares_out is None:
        shares_out = _to_float(fast_info.get("shares") or fast_info.get("shares_outstanding"))
    if shares_out is None:
        shares_out = _fmp_get_num(fmp_data, "profile", "sharesOutstanding")
    if shares_out is None and (mcap is not None) and (price not in (None, 0)):
        shares_out = float(mcap) / float(price)


    # --- NUPL legacy (Adj Close) ---
    twap10 = twap(p10)
    twap5 = twap(p5)
    nupl10 = (price - twap10) / price if (price and twap10) else None
    nupl5 = (price - twap5) / price if (price and twap5) else None

    ath_adj = float(p10.max()) if not p10.empty else None
    cycle_proxy = normalize_cycle_proxy(price, ath_adj) if (price and ath_adj) else None

    # --- NUPL v2 components ---
    vwap10_series = anchored_vwap(p10, h10["Volume"]) if (h10 is not None and not h10.empty and "Volume" in h10.columns) else pd.Series(dtype=float)
    vwap10 = float(vwap10_series.iloc[-1]) if not vwap10_series.empty else None
    nupl_vwap10 = (price - vwap10) / price if (price and vwap10) else None

    # shares_out already derived above (with multiple fallbacks)

    realized_series = turnover_realized_price(p10, h10["Volume"], shares_out) if (h10 is not None and not h10.empty and "Volume" in h10.columns) else pd.Series(dtype=float)
    realized_price = float(realized_series.iloc[-1]) if not realized_series.empty else None
    nupl_turnover = (price - realized_price) / price if (price and realized_price) else None

    vol_1y = realized_vol_1y(h1y)
    dd_sigma = None
    if cycle_proxy is not None and vol_1y not in (None, 0):
        dd_sigma = float(cycle_proxy) / (float(vol_1y) / 100.0)

    composite_nupl = None
    weights = {"turn": 0.50, "vwap": 0.30, "cycle": 0.20}
    vals = {"turn": nupl_turnover, "vwap": nupl_vwap10, "cycle": cycle_proxy}
    wsum = sum(weights[k] for k in weights if vals[k] is not None)
    if wsum > 0:
        composite_nupl = sum((weights[k] / wsum) * vals[k] for k in weights if vals[k] is not None)

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
    annual_income = _cached_df(f"stmt:{ticker}:income_stmt", lambda: yf_call(lambda: tkr.income_stmt))
    if annual_income is None:
        annual_income = pd.DataFrame()
    annual_bs = _cached_df(f"stmt:{ticker}:balance_sheet", lambda: yf_call(lambda: tkr.balance_sheet))
    if annual_bs is None:
        annual_bs = pd.DataFrame()
    annual_cf = _cached_df(f"stmt:{ticker}:cashflow", lambda: yf_call(lambda: tkr.cashflow))
    if annual_cf is None:
        annual_cf = pd.DataFrame()

    # Valuation basics (Yahoo → FMP fallback)
    trailing_pe = safe_get(info, "trailingPE")
    ev_ebitda = safe_get(info, "enterpriseToEbitda")

    beta = safe_get(info, "beta")
    if beta is None:
        beta = _fmp_get_num(fmp_data, "profile", "beta")

    short_pct = safe_get(info, "shortPercentOfFloat")
    short_ratio = safe_get(info, "shortRatio")

    if ev is None:
        ev = safe_get(info, "enterpriseValue")
    if ev is None:
        ev = _fmp_get_num(fmp_data, "enterprise_value", "enterpriseValue")

    if mcap is None:
        mcap = safe_get(info, "marketCap")
    if mcap is None:
        mcap = _fmp_get_num(fmp_data, "quote", "marketCap")

    # TTM revenue & FCF
    q_income = _cached_df(f"stmt:{ticker}:q_income", lambda: yf_call(lambda: tkr.quarterly_income_stmt))
    if q_income is None:
        q_income = pd.DataFrame()
    q_cf = _cached_df(f"stmt:{ticker}:q_cf", lambda: yf_call(lambda: tkr.quarterly_cashflow))
    if q_cf is None:
        q_cf = pd.DataFrame()

    ttm_rev = last_n_quarters_sum(q_income, ["Total Revenue", "TotalRevenue"], n=4)

    ttm_ocf = last_n_quarters_sum(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"], n=4)
    ttm_capex = last_n_quarters_sum(q_cf, ["Capital Expenditure", "CapitalExpenditure"], n=4)
    ttm_fcf = (ttm_ocf + ttm_capex) if (ttm_ocf is not None and ttm_capex is not None) else None

    # P/S
    ps = safe_get(info, "priceToSalesTrailing12Months")
    if ps is None:
        ps = _fmp_get_num(fmp_data, "ratios_ttm", "priceToSalesRatioTTM", "priceToSalesRatio")
    if ps is None and mcap is not None and ttm_rev not in (None, 0):
        ps = float(mcap) / float(ttm_rev)

    # -------- EV/EBIT (stability patch) --------
    ttm_ebit = last_n_quarters_sum(q_income, ["Operating Income", "OperatingIncome"], n=4)

    annual_ebit = None
    if annual_income is not None and not annual_income.empty:
        for k in ["EBIT", "Ebit", "Operating Income", "OperatingIncome"]:
            if k in annual_income.index:
                try:
                    annual_ebit = float(annual_income.loc[k].iloc[0])
                except Exception:
                    annual_ebit = None
                break

    ebit_used = ttm_ebit if ttm_ebit not in (None, 0) else annual_ebit

    ev_ebit = None
    earnings_yield = None

    if ev in (None, 0) or ebit_used is None:
        notes["EV/EBIT"] = "No meaningful data to calculate this ratio (missing EV or EBIT proxy)."
        notes["Earnings Yield (EBIT / EV)"] = "No meaningful data to calculate this ratio (missing EV or EBIT proxy)."
    elif ebit_used <= 0:
        notes["EV/EBIT"] = "No meaningful data to calculate this ratio (EBIT <= 0)."
        notes["Earnings Yield (EBIT / EV)"] = "No meaningful data to calculate this ratio (EBIT <= 0)."
    else:
        if ttm_rev not in (None, 0) and (float(ebit_used) / float(ttm_rev)) < 0.005:
            notes["EV/EBIT"] = "No meaningful data to calculate this ratio (EBIT margin < 0.5% → ratio becomes unstable)."
            notes["Earnings Yield (EBIT / EV)"] = "No meaningful data to calculate this ratio (EBIT margin < 0.5% → inverse becomes unstable)."
        else:
            ev_ebit = float(ev) / float(ebit_used)
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

    if gross_m is None:
        gm = _fmp_get_num(fmp_data, "ratios_ttm", "grossProfitMarginTTM", "grossProfitMargin")
        if gm is not None:
            gross_m = gm if abs(gm) <= 1.5 else (gm / 100.0)
    if op_m is None:
        om = _fmp_get_num(fmp_data, "ratios_ttm", "operatingProfitMarginTTM", "operatingProfitMargin")
        if om is not None:
            op_m = om if abs(om) <= 1.5 else (om / 100.0)
    if net_m is None:
        nm = _fmp_get_num(fmp_data, "ratios_ttm", "netProfitMarginTTM", "netProfitMargin")
        if nm is not None:
            net_m = nm if abs(nm) <= 1.5 else (nm / 100.0)
    if roe is None:
        r = _fmp_get_num(fmp_data, "ratios_ttm", "returnOnEquityTTM", "returnOnEquity")
        if r is not None:
            roe = r if abs(r) <= 1.5 else (r / 100.0)

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
    if ebitda is None:
        ebitda = _fmp_get_num(fmp_data, "key_metrics_ttm", "ebitdaTTM", "ebitda")

    nd_ebitda = float(net_debt) / float(ebitda) if (net_debt is not None and ebitda not in (None, 0)) else None
    nd_fcf = float(net_debt) / float(ttm_fcf) if (net_debt is not None and ttm_fcf not in (None, 0)) else None

    # Interest coverage
    interest_cov = None
    fcf_interest = None
    if annual_income is not None and not annual_income.empty:
        interest_exp = None
        for k in ["Interest Expense", "InterestExpense"]:
            if k in annual_income.index:
                interest_exp = annual_income.loc[k].iloc[0]
                break
        if interest_exp is not None and not pd.isna(interest_exp):
            denom = abs(float(interest_exp))
            if denom > 0 and ebit_used is not None:
                interest_cov = float(ebit_used) / denom
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
    if shares is None:
        shares = shares_out

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
        # PERF: reuse existing history slice (no extra Yahoo call)
        h3m = _slice_history_days(h10, 110)  # ~3-4 months of calendar days
        if h3m is not None and not h3m.empty and "Volume" in h3m.columns:
            px = _adj_close(h3m)
            dv = (px * pd.to_numeric(h3m["Volume"], errors="coerce")).dropna()
            if len(dv) > 10:
                avg_dollar_vol = float(dv.mean())
    except Exception:
        pass
    worst_week_3y = worst_weekly_return_3y(h3y)

    # --- DAVF (Downside-Anchored Value Floor) ---
    davf = compute_davf(
        sector_bucket=sector_bucket,
        price=price,
        shares_out=shares_out,
        annual_income=annual_income,
        annual_cf=annual_cf,
        annual_bs=annual_bs,
        roic_pct=roic_pct,
        nd_ebitda=nd_ebitda,
        interest_cov=interest_cov,
    )

    # Add to notes so it shows up in report (Notes column + header block)
    if davf.get("davf_note"):
        notes["DAVF"] = davf["davf_note"]
    if davf.get("davf_label") and davf.get("davf_label") != "NA":
        notes["DAVF Downside Protection"] = f"{davf['davf_label']} (MOS={davf.get('davf_mos_pct'):.1f}%, conf={davf.get('davf_confidence')})"



    # --- Value-matrix extras (fills missing checklist metrics + adds stability metrics) ---
    extra_metrics: Dict[str, Any] = {}
    try:
        extra_metrics, extra_notes = compute_value_matrix_extras(
            info=info,
            fmp_bundle=fmp_data,
            price=price,
            market_cap=mcap,
            enterprise_value=ev,
            annual_income=annual_income,
            annual_cashflow=annual_cf,
            annual_balance_sheet=annual_bs,
            quarterly_cashflow=q_cf,
        )
        if extra_notes:
            notes.update(extra_notes)
    except Exception as e:
        notes["Value Matrix Extras"] = f"Failed to compute extras: {e}"

    metrics = {
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

        # NUPL legacy keys
        "NUPL10Y": nupl10,
        "NUPL5Y": nupl5,
        "CycleProxy": cycle_proxy,

        # NUPL v2 components
        "NUPL VWAP10 Proxy": nupl_vwap10,
        "NUPL Turnover Proxy": nupl_turnover,
        "Realized Price (Turnover Proxy)": realized_price,
        "VWAP10 (Anchored)": vwap10,
        "Drawdown Sigma (ATH/Vol)": dd_sigma,

        # Composite
        "Composite NUPL": composite_nupl,
        "Composite NUPL Z": composite_z,
        "NUPL Regime": nupl_regime,

        # DAVF (Downside-Anchored Value Floor)
        "DAVF Value Floor (per share)": davf.get("davf_floor_ps"),
        "DAVF MOS vs Floor (%)": davf.get("davf_mos_pct"),
        "DAVF Confidence": davf.get("davf_confidence"),
        "DAVF Base Type": davf.get("davf_base_type"),
        "DAVF Multiple Used": davf.get("davf_multiple"),
        "DAVF Haircut (%)": davf.get("davf_haircut_pct"),
        "DAVF Downside Protection": davf.get("davf_label"),

    }

    # Merge extras last so they can overwrite base placeholders (e.g., previously-NA checklist metrics).
    if extra_metrics:
        metrics.update(extra_metrics)


    # --- Bundle Yahoo objects so reversal scoring can reuse them without extra calls ---
    metrics["__yf_bundle__"] = {
        "info": info,
        "q_income": q_income,
        "q_cf": q_cf,
        "annual_bs": annual_bs,
        "h10": h10,
        "h1y": h1y,
        "h2y": h2y,
    }

    metrics["__notes__"] = notes
    return metrics
