import math
from typing import Any, Dict, Optional

import numpy as np
import pandas as pd
import yfinance as yf

from cache_utils import DiskCache, yf_call, yf_cache_settings
from metrics import anchored_vwap, turnover_realized_price, zscore_last, stock_nupl_regime_from_z


def _safe_float(x) -> Optional[float]:
    try:
        if x is None:
            return None
        if isinstance(x, (float, int)):
            if isinstance(x, float) and (math.isnan(x) or math.isinf(x)):
                return None
            return float(x)
        return float(str(x))
    except Exception:
        return None


def _max_drawdown(prices: pd.Series) -> Optional[float]:
    p = pd.to_numeric(prices, errors="coerce").dropna()
    if len(p) < 30:
        return None
    roll_max = p.cummax()
    dd = (p / roll_max) - 1.0
    return float(dd.min()) * 100.0


def _realized_vol(prices: pd.Series, window: int = 252) -> Optional[float]:
    p = pd.to_numeric(prices, errors="coerce").dropna()
    if len(p) < window // 2:
        return None
    rets = p.pct_change().dropna()
    if rets.empty:
        return None
    # annualized stdev
    vol = float(rets.std(ddof=0)) * math.sqrt(252.0) * 100.0
    return vol


def compute_metrics_scan(
    ticker: str,
    *,
    use_yf_cache: Optional[bool] = None,
    history_years: str = "10y",
) -> Dict[str, Any]:
    """Cheap broad-scan metrics: Yahoo-only, one history call, no statements.

    Goal: rank/filter large universes quickly.
    """
    tkr = yf.Ticker(ticker)

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

    # info is slower than fast_info but still useful for sector/currency; cache it.
    info = _cached_json(f"info:{ticker}", lambda: yf_call(lambda: tkr.get_info() or {})) or {}

    # One history call (10y) then slice for 1y/3y stats.
    hist = _cached_df(f"hist:{history_years}:{ticker}", lambda: yf_call(lambda: tkr.history(period=history_years, auto_adjust=False)))
    if hist is None or hist.empty:
        return {"Ticker": ticker, "__notes__": {"scan": "No price history"}}

    close = hist.get("Close") if "Close" in hist.columns else hist.iloc[:, 0]
    vol = hist.get("Volume") if "Volume" in hist.columns else pd.Series(index=hist.index, dtype=float)

    # Basic price / size / liquidity
    fast = {}
    try:
        fi = getattr(tkr, "fast_info", None)
        if fi:
            # yfinance fast_info behaves like dict-ish
            fast = dict(fi)
    except Exception:
        fast = {}

    price = _safe_float(fast.get("last_price") or info.get("currentPrice") or close.dropna().iloc[-1])
    mcap = _safe_float(fast.get("market_cap") or info.get("marketCap"))
    shares = _safe_float(info.get("sharesOutstanding"))
    beta = _safe_float(info.get("beta"))

    # Avg $ volume (approx using last 63 trading days)
    tail_63 = hist.tail(63)
    if tail_63 is not None and not tail_63.empty and price is not None:
        adv = _safe_float(tail_63["Volume"].mean()) if "Volume" in tail_63.columns else None
        avg_dollar_vol = (adv * price) if (adv is not None and price is not None) else None
    else:
        avg_dollar_vol = None

    # Risk
    hist_3y = hist.tail(252 * 3)
    hist_1y = hist.tail(252)

    maxdd_3y = _max_drawdown(hist_3y["Close"] if "Close" in hist_3y.columns else close.tail(252 * 3))
    vol_1y = _realized_vol(hist_1y["Close"] if "Close" in hist_1y.columns else close.tail(252), window=252)

    # NUPL-like proxy (same core ingredients as compute_metrics_v2, simplified)
    # realized price via turnover + anchored VWAP
    realized = turnover_realized_price(close, vol, shares)
    vwap_anch = anchored_vwap(close, vol)

    nupl_turnover = (close - realized) / realized
    nupl_vwap = (close - vwap_anch) / vwap_anch

    # Z-score on latest for regime
    z = zscore_last(nupl_turnover, window=2520)
    regime = stock_nupl_regime_from_z(z)

    out: Dict[str, Any] = {
        "Ticker": ticker,
        "Price": price,
        "Market Cap": mcap,
        "Avg $ Volume (63d)": avg_dollar_vol,
        "Beta": beta,
        "Max Drawdown 3Y %": maxdd_3y,
        "Realized Vol 1Y %": vol_1y,
        "NUPL Turnover10 Proxy": float(nupl_turnover.dropna().iloc[-1]) if len(nupl_turnover.dropna()) else None,
        "NUPL VWAP10 Proxy": float(nupl_vwap.dropna().iloc[-1]) if len(nupl_vwap.dropna()) else None,
        "Stock NUPL Z (10Y)": z,
        "Stock NUPL Regime": regime,
        "Sector": info.get("sector"),
        "Currency": info.get("currency") or info.get("financialCurrency"),
        "__notes__": {
            "scan": "Broad scan metrics (Yahoo-only). No statements / no DAVF / no reversal."
        }
    }
    return out
