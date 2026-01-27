import re
from typing import Optional
import yfinance as yf

TICKER_RE = re.compile(r"^[A-Z0-9\.\-\^=]+$")

def looks_like_ticker(s: str) -> bool:
    s = s.strip().upper()
    return bool(TICKER_RE.match(s)) and any(c.isalpha() for c in s)

def try_validate_ticker(ticker: str) -> bool:
    try:
        t = yf.Ticker(ticker)
        h = t.history(period="5d", interval="1d")
        return h is not None and not h.empty
    except Exception:
        return False

def resolve_to_ticker(query: str) -> Optional[str]:
    q = query.strip()
    if not q:
        return None

    if looks_like_ticker(q):
        t = q.upper()
        if try_validate_ticker(t):
            return t

    try:
        lk = yf.Lookup(q)
        stocks = lk.get_stock(count=5) if hasattr(lk, "get_stock") else lk.stock
        if isinstance(stocks, list) and len(stocks) > 0:
            candidate = stocks[0].get("symbol") or stocks[0].get("ticker")
            if candidate and try_validate_ticker(candidate.upper()):
                return candidate.upper()
    except Exception:
        pass

    try:
        sr = yf.Search(q, max_results=10)
        quotes = getattr(sr, "quotes", None)
        if isinstance(quotes, list) and len(quotes) > 0:
            sym = quotes[0].get("symbol")
            if sym and try_validate_ticker(sym.upper()):
                return sym.upper()
    except Exception:
        pass

    return None
