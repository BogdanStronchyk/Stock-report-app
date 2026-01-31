import os
import json
import time
import hashlib
import urllib.parse
import urllib.request
import urllib.error
from datetime import datetime
from typing import Any, Dict, Optional


class FMPClient:
    """Financial Modeling Prep fallback client.

    IMPORTANT (post Aug 31, 2025): FMP's legacy `/api/v3/...` endpoints are
    blocked for non-legacy accounts.

    This client uses the **stable** API exclusively:

        https://financialmodelingprep.com/stable/<endpoint>?symbol=SYMBOL&apikey=...

    Performance/budget features:
      - Optional on-disk cache (default ON, TTL 24h) to avoid repeat calls.
      - Optional request budget per run (FMP_MAX_REQUESTS).
      - STRICT MINIMALISM: No retries, no fallbacks, early exit on bad tickers.
    """

    def __init__(self, api_key: Optional[str] = None, timeout: int = 25):
        self.api_key = (api_key or os.environ.get("FMP_API_KEY") or "").strip()
        self.timeout = timeout
        self.enabled = bool(self.api_key)

        self.request_count = 0
        self.debug = os.environ.get("FMP_DEBUG", "").strip().lower() in ("1", "true", "yes")
        self.last_error = None
        self.last_status = None

        # Budget controls
        self.max_requests = self._safe_int(os.environ.get("FMP_MAX_REQUESTS", ""), default=0)

        # Cache controls
        self.cache_enabled = (os.environ.get("FMP_USE_CACHE", "1") or "1").strip().lower() not in ("0", "false", "no")
        self.cache_ttl_hours = self._safe_float(os.environ.get("FMP_CACHE_TTL_HOURS", "24"), default=24.0)

        # Retry controls: Default to 0 to protect quota
        self.max_retries = self._safe_int(os.environ.get("FMP_MAX_RETRIES", "0"), default=0)
        self.backoff_base = self._safe_float(os.environ.get("FMP_BACKOFF_BASE", "1.5"), default=1.5)

        # Plan restrictions
        self.use_quarterly = (os.environ.get("FMP_USE_QUARTERLY", "0") or "0").strip().lower() in ("1", "true", "yes")
        self.statement_limit = self._safe_int(os.environ.get("FMP_STATEMENT_LIMIT", "5"), default=5)
        self.statement_limit = max(0, min(5, self.statement_limit))

    def clean_ticker(self, symbol: str) -> str:
        """Align Yahoo-style tickers to FMP format to avoid 400/404 errors."""
        s = (symbol or "").strip().upper()
        # 1. Forex: 'EURUSD=X' -> 'EURUSD'
        if s.endswith("=X"):
            s = s.replace("=X", "")
        # 2. Crypto: 'BTC-USD' -> 'BTCUSD'
        if s.endswith("-USD"):
            s = s.replace("-", "")
        # 3. US Share Classes: 'BRK.B' -> 'BRK-B'
        if "." in s:
            parts = s.split(".")
            if len(parts) == 2 and parts[0].isalpha() and len(parts[1]) == 1:
                return s.replace(".", "-")
        return s

    def _cap_limit(self, desired: int) -> Optional[int]:
        try:
            desired_i = int(desired)
        except Exception:
            desired_i = 0
        if self.statement_limit <= 0:
            return None
        return min(max(0, desired_i), self.statement_limit)

    @staticmethod
    def _safe_int(v: str, default: int = 0) -> int:
        try:
            return int(float((v or "").strip()))
        except Exception:
            return default

    @staticmethod
    def _safe_float(v: str, default: float = 0.0) -> float:
        try:
            return float((v or "").strip())
        except Exception:
            return default

    # -----------------------
    # URL / Cache
    # -----------------------
    def _stable_url(self, endpoint: str, params: Optional[Dict[str, Any]] = None) -> str:
        params = dict(params or {})
        params["apikey"] = self.api_key
        qs = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
        base = (os.environ.get("FMP_BASE_URL") or "https://financialmodelingprep.com/stable/").strip().rstrip("/")
        endpoint = (endpoint or "").lstrip("/")
        return f"{base}/{endpoint}?{qs}"

    def _cache_dir(self) -> str:
        return os.path.join(os.getcwd(), ".cache", "fmp")

    def _cache_key(self, url: str) -> str:
        return hashlib.sha256(url.encode("utf-8", errors="ignore")).hexdigest() + ".json"

    def _cache_get(self, url: str) -> Optional[Any]:
        if not self.cache_enabled: return None
        try:
            path = os.path.join(self._cache_dir(), self._cache_key(url))
            st = os.stat(path)
            if (datetime.now().timestamp() - st.st_mtime) > (self.cache_ttl_hours * 3600):
                return None
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _cache_set(self, url: str, payload: Any) -> None:
        if not self.cache_enabled: return
        try:
            cdir = self._cache_dir()
            os.makedirs(cdir, exist_ok=True)
            path = os.path.join(cdir, self._cache_key(url))
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f)
        except Exception:
            pass

    # -----------------------
    # HTTP
    # -----------------------
    def _get_json(self, url: str) -> Optional[Any]:
        if not self.enabled:
            return None
        if self.max_requests > 0 and self.request_count >= self.max_requests:
            self.last_error = "Budget limit reached."
            return None

        cached = self._cache_get(url)
        if cached is not None:
            self.last_status = 200
            self.last_error = None
            return cached

        headers = {"User-Agent": "StockReportApp/1.0", "Connection": "close"}

        # Max retries loop (default 0 loops = 1 attempt total)
        for attempt in range(max(1, self.max_retries + 1)):
            req = urllib.request.Request(url, headers=headers, method="GET")
            try:
                self.request_count += 1
                if self.debug:
                    print(f"[FMP] GET {url}")

                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    self.last_status = getattr(resp, "status", 200)
                    raw = resp.read().decode("utf-8", errors="ignore")
                    if not raw: return None
                    payload = json.loads(raw)
                    self._cache_set(url, payload)
                    self.last_error = None
                    return payload

            except urllib.error.HTTPError as e:
                self.last_status = e.code
                # 429 Rate Limit
                if e.code == 429 and attempt < self.max_retries:
                    time.sleep(1.0)  # Simple sleep
                    continue
                # All other errors (400, 403, 404, 500) -> Abort immediately to save quota
                self.last_error = f"HTTP {e.code}"
                return None
            except Exception as e:
                self.last_error = str(e)
                if attempt < self.max_retries:
                    time.sleep(0.5)
                    continue
                return None
        return None

    # -----------------------
    # Fetch (Strict Minimal)
    # -----------------------
    def fetch_minimal(
            self,
            symbol: str,
            *,
            need_profile: bool = True,
            need_quote: bool = True,
            need_income_annual: bool = False,
            need_income_quarter: bool = False,
            need_cashflow_quarter: bool = False,
            need_balance_annual: bool = False,
            need_enterprise_value: bool = False,
            need_ratios_ttm: bool = False,
            need_key_metrics_ttm: bool = True,
    ) -> Dict[str, Any]:
        """Fetch endpoints with strict call minimization."""
        if not self.enabled:
            return {}

        # 1. Clean ticker (prevents 400s)
        symbol = self.clean_ticker(symbol)
        if not symbol:
            return {}

        data: Dict[str, Any] = {}

        def _get(endpoint: str, params: Dict[str, Any]) -> Optional[Any]:
            return self._get_json(self._stable_url(endpoint, params))

        # 2. Fetch Profile First
        # Optimization: If profile fails (invalid ticker), abort immediately to save calls.
        if need_profile:
            p = _get("profile", {"symbol": symbol})
            data["profile"] = p or []

            # EARLY EXIT: If we expected a profile but got nothing, the ticker is likely bad/delisted.
            # Stop here to save the other ~6 calls.
            if not data["profile"]:
                if self.debug:
                    print(f"[FMP] Aborting {symbol} (no profile found) to save quota.")
                return data
        else:
            data["profile"] = []

        # 3. Fetch Quote (Single Attempt)
        if need_quote:
            data["quote"] = _get("quote", {"symbol": symbol}) or []
        else:
            data["quote"] = []

        # 4. Fetch Statements (Single Attempt - No Fallbacks)
        # We assume strict compliance with FMP plans (e.g. Annual default).
        if need_income_annual:
            data["income_a"] = _get("income-statement",
                                    {"symbol": symbol, "period": "annual", "limit": self._cap_limit(6)}) or []
        else:
            data["income_a"] = []

        if need_income_quarter:
            data["income_q"] = _get("income-statement",
                                    {"symbol": symbol, "period": "quarter", "limit": self._cap_limit(8)}) or []
        else:
            data["income_q"] = []

        if need_cashflow_quarter:
            data["cashflow_q"] = _get("cash-flow-statement",
                                      {"symbol": symbol, "period": "quarter", "limit": self._cap_limit(8)}) or []
        else:
            data["cashflow_q"] = []

        if need_balance_annual:
            data["balance_a"] = _get("balance-sheet-statement",
                                     {"symbol": symbol, "period": "annual", "limit": self._cap_limit(2)}) or []
        else:
            data["balance_a"] = []

        if need_enterprise_value:
            data["enterprise_value"] = _get("enterprise-values", {"symbol": symbol, "limit": self._cap_limit(2)}) or []
        else:
            data["enterprise_value"] = []

        # TTM endpoints
        data["ratios_ttm"] = _get("ratios-ttm", {"symbol": symbol}) or [] if need_ratios_ttm else []
        data["key_metrics_ttm"] = _get("key-metrics-ttm", {"symbol": symbol}) or [] if need_key_metrics_ttm else []

        return data

    def fetch_bundle(self, symbol: str, *, mode: str = "full") -> Dict[str, Any]:
        """Budget-aware fetch wrapper."""
        mode = (mode or "full").strip().lower()
        if mode in ("off", "0", "false", "no"):
            return {}
        if mode == "conditional":
            return self.fetch_minimal(
                symbol,
                need_profile=True,
                need_quote=True,
                need_enterprise_value=False,
                need_ratios_ttm=False,
                need_key_metrics_ttm=True,
            )
        if mode == "minimal":
            return self.fetch_minimal(
                symbol,
                need_profile=True,
                need_quote=True,
                need_enterprise_value=True,
                need_ratios_ttm=True,
                need_key_metrics_ttm=True,
            )
        # default full
        return self.fetch_all(symbol)

    def fetch_all(self, symbol: str) -> Dict[str, Any]:
        return self.fetch_minimal(
            symbol,
            need_profile=True,
            need_quote=True,
            need_income_annual=True,
            need_income_quarter=self.use_quarterly,
            need_cashflow_quarter=self.use_quarterly,
            need_balance_annual=True,
            need_enterprise_value=True,
            need_ratios_ttm=True,
            need_key_metrics_ttm=True,
        )