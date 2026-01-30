import os
import json
import time
import threading
import hashlib
import urllib.parse
import urllib.request
import urllib.error
from datetime import datetime
from typing import Any, Dict, Optional, Tuple


class FMPClient:
    """Financial Modeling Prep fallback client.

    Features:
      - Thread-safe global stats (Success/Error counts) for UI tracking.
      - Stable API endpoints only.
      - Caching and Budgeting.
    """

    # --- Global Stats (Thread-Safe) ---
    _stats_lock = threading.Lock()
    _stat_success = 0
    _stat_error = 0

    @classmethod
    def get_stats(cls) -> Tuple[int, int]:
        """Returns (success_count, error_count)"""
        with cls._stats_lock:
            return cls._stat_success, cls._stat_error

    @classmethod
    def _inc_success(cls):
        with cls._stats_lock:
            cls._stat_success += 1

    @classmethod
    def _inc_error(cls):
        with cls._stats_lock:
            cls._stat_error += 1

    def __init__(self, api_key: Optional[str] = None, timeout: int = 25):
        self.api_key = (api_key or os.environ.get("FMP_API_KEY") or "").strip()
        self.timeout = timeout
        self.enabled = bool(self.api_key)

        self.max_retries = int(self._safe_int(os.environ.get("FMP_MAX_RETRIES", "2")))
        self.backoff_base = self._safe_float(os.environ.get("FMP_BACKOFF_BASE", "1.2"), default=1.2)

        # Budget controls
        self.request_count = 0
        self.max_requests = self._safe_int(os.environ.get("FMP_MAX_REQUESTS", "0"), default=0)

        # Debug / Status tracking
        self.debug = os.environ.get("FMP_DEBUG", "").strip().lower() in ("1", "true", "yes")
        self.last_error = None
        self.last_status = None

        # Cache settings
        self.use_cache = (os.environ.get("FMP_USE_CACHE", "1") or "1").strip().lower() not in ("0", "false", "no")
        self.cache_ttl_hours = self._safe_float(os.environ.get("FMP_CACHE_TTL_HOURS", "24"), default=24.0)
        self.cache_dir = os.path.join(os.getcwd(), ".cache", "fmp")

        # Quarterly vs Annual preference
        self.use_quarterly = (os.environ.get("FMP_USE_QUARTERLY", "0") or "0").strip().lower() in ("1", "true", "yes")
        self.statement_limit = self._safe_int(os.environ.get("FMP_STATEMENT_LIMIT", "5"), default=5)
        self.statement_limit = max(0, min(5, self.statement_limit))

    @staticmethod
    def _safe_int(v: Any, default: int = 0) -> int:
        try:
            return int(float(v))
        except Exception:
            return default

    @staticmethod
    def _safe_float(v: Any, default: float = 0.0) -> float:
        try:
            return float(v)
        except Exception:
            return default

    def _cap_limit(self, desired: int) -> Optional[int]:
        if self.statement_limit <= 0:
            return None
        return min(max(0, desired), self.statement_limit)

    # -----------------------
    # Fetch Methods
    # -----------------------

    def fetch_bundle(self, symbol: str, *, mode: str = "full") -> Dict[str, Any]:
        """Budget-aware fetch wrapper.

        Required by metrics.py.
        mode: off | conditional | minimal | full
        """
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
        """Smart fetcher that aggregates multiple endpoints."""
        if not self.enabled:
            return {}

        # Helper to simplify calls
        def _get(endpoint: str, params: Dict[str, Any]) -> Optional[Any]:
            return self._get_data(endpoint, symbol, params)

        data: Dict[str, Any] = {}

        # 1. Profile & Quote
        if need_profile:
            p = _get("profile", {})
            data["profile"] = p[0] if isinstance(p, list) and p else p

        if need_quote:
            q = _get("quote", {})
            if q is None:  # Fallback
                q = _get("quote-short", {})
            data["quote"] = q[0] if isinstance(q, list) and q else q

        # 2. Key Metrics & Ratios (TTM)
        if need_key_metrics_ttm:
            km = _get("key-metrics-ttm", {})
            data["key_metrics_ttm"] = km[0] if isinstance(km, list) and km else km

        if need_ratios_ttm:
            rt = _get("ratios-ttm", {})
            data["ratios_ttm"] = rt[0] if isinstance(rt, list) and rt else rt

        # 3. Enterprise Value
        if need_enterprise_value:
            ev = _get("enterprise-values", {"limit": self._cap_limit(2)})
            data["enterprise_value"] = ev[0] if isinstance(ev, list) and ev else ev

        # 4. Statements
        if need_income_annual:
            data["income_a"] = _get("income-statement", {"period": "annual", "limit": self._cap_limit(6)}) or []

        if need_income_quarter:
            data["income_q"] = _get("income-statement", {"period": "quarter", "limit": self._cap_limit(8)}) or []

        if need_cashflow_quarter:
            data["cashflow_q"] = _get("cash-flow-statement", {"period": "quarter", "limit": self._cap_limit(8)}) or []

        if need_balance_annual:
            data["balance_a"] = _get("balance-sheet-statement", {"period": "annual", "limit": self._cap_limit(2)}) or []

        return data

    def fetch_all(self, symbol: str) -> Dict[str, Any]:
        """Full bundle: Quote, Profile, Metrics, Ratios, and basic Statements."""
        return self.fetch_minimal(
            symbol,
            need_profile=True,
            need_quote=True,
            need_enterprise_value=True,
            need_ratios_ttm=True,
            need_key_metrics_ttm=True,
            need_income_annual=True,
            need_income_quarter=self.use_quarterly,
            need_cashflow_quarter=self.use_quarterly,
            need_balance_annual=True
        )

    # -----------------------
    # Internal Logic
    # -----------------------

    def _get_data(self, endpoint: str, symbol: str, params: Dict[str, Any]) -> Any:
        """Internal requester with Retry, Cache, and Stats counting."""
        if not self.enabled:
            return None

        # Budget check
        if self.max_requests > 0 and self.request_count >= self.max_requests:
            self.last_error = "Budget Exceeded"
            return None

        # Build URL
        base = "https://financialmodelingprep.com/stable/"
        p_copy = params.copy()
        p_copy["symbol"] = symbol
        p_copy["apikey"] = self.api_key
        qs = urllib.parse.urlencode({k: v for k, v in p_copy.items() if v is not None})
        url = f"{base}{endpoint}?{qs}"

        # 1. Check Cache
        cache_key = hashlib.sha256(url.encode("utf-8")).hexdigest()
        cached = self._read_cache(cache_key)
        if cached is not None:
            return cached

        # 2. Network Request
        for attempt in range(self.max_retries + 1):
            try:
                self.request_count += 1
                with urllib.request.urlopen(url, timeout=self.timeout) as response:
                    if response.status == 200:
                        data = json.loads(response.read().decode())

                        # Empty list check (soft error)
                        if isinstance(data, list) and not data:
                            pass

                        self._write_cache(cache_key, data)

                        # --- STATS: SUCCESS ---
                        self._inc_success()
                        return data

            except urllib.error.HTTPError as e:
                self.last_status = e.code
                if e.code in (403, 404):
                    self._inc_error()
                    print(f"FMP HTTP {e.code}: {symbol} {endpoint}")
                    return None
                if e.code == 429:
                    # Rate limit -> Sleep and retry
                    retry_after = 1.0
                    try:
                        retry_after = float(e.headers.get("Retry-After", 1.0))
                    except:
                        pass
                    time.sleep(retry_after * (attempt + 1))
                    continue

            except Exception as e:
                # Network/Timeout -> Retry
                time.sleep(0.5 * (attempt + 1))

        # --- STATS: FAILURE ---
        self._inc_error()
        return None

    def _read_cache(self, key: str) -> Optional[Any]:
        if not self.use_cache: return None
        try:
            path = os.path.join(self.cache_dir, f"{key}.json")
            if not os.path.exists(path): return None

            # TTL Check
            if time.time() - os.path.getmtime(path) > (self.cache_ttl_hours * 3600):
                return None

            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return None

    def _write_cache(self, key: str, data: Any):
        if not self.use_cache or data is None: return
        try:
            os.makedirs(self.cache_dir, exist_ok=True)
            path = os.path.join(self.cache_dir, f"{key}.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f)
        except:
            pass