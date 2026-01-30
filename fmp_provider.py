import os
import json
import time
import threading
import hashlib
import requests
from typing import Any, Dict, Optional, Tuple


class FMPClient:
    """Financial Modeling Prep client (Safe Mode).

    CRITICAL FIXES:
      1. URL Structure: Fixes ".../quote?symbol=TICKER" -> ".../quote/TICKER".
      2. Library: Uses 'requests' so 'requests_cache' in main.py works.
      3. Circuit Breaker: Stops immediately if Profile/Quote fails.
    """

    # --- Global Stats (Thread-Safe) ---
    _stats_lock = threading.Lock()
    _stat_success = 0
    _stat_error = 0

    @classmethod
    def get_stats(cls) -> Tuple[int, int]:
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
        print("[SYSTEM] FMP Client v2 (Requests + URL Fix) Loaded.")  # CONFIRMATION LOG

        self.api_key = (api_key or os.environ.get("FMP_API_KEY") or "").strip()
        self.timeout = timeout
        self.enabled = bool(self.api_key)
        self.session = requests.Session()

        # RETRIES: Default to 1 to save quota
        self.max_retries = int(self._safe_int(os.environ.get("FMP_MAX_RETRIES", "1")))

        self.request_count = 0
        self.max_requests = self._safe_int(os.environ.get("FMP_MAX_REQUESTS", "0"), default=0)

        # Cache settings
        self.use_cache = (os.environ.get("FMP_USE_CACHE", "1") or "1").strip().lower() not in ("0", "false", "no")
        self.cache_ttl_hours = self._safe_float(os.environ.get("FMP_CACHE_TTL_HOURS", "24"), default=24.0)
        self.cache_dir = os.path.join(os.getcwd(), ".cache", "fmp")

        self.use_quarterly = (os.environ.get("FMP_USE_QUARTERLY", "0") or "0").strip().lower() in ("1", "true", "yes")
        self.statement_limit = max(0, min(5, self._safe_int(os.environ.get("FMP_STATEMENT_LIMIT", "5"), default=5)))

        if self.use_cache:
            os.makedirs(self.cache_dir, exist_ok=True)

    @staticmethod
    def _safe_int(v: Any, default: int = 0) -> int:
        try:
            return int(float(v))
        except:
            return default

    @staticmethod
    def _safe_float(v: Any, default: float = 0.0) -> float:
        try:
            return float(v)
        except:
            return default

    def _cap_limit(self, desired: int) -> Optional[int]:
        if self.statement_limit <= 0: return None
        return min(max(0, desired), self.statement_limit)

    # -----------------------
    # Fetch Methods
    # -----------------------

    def fetch_bundle(self, symbol: str, *, mode: str = "full") -> Dict[str, Any]:
        mode = (mode or "full").strip().lower()
        if mode in ("off", "0", "false", "no"): return {}

        if mode == "conditional":
            return self.fetch_minimal(symbol, need_profile=True, need_quote=True, need_key_metrics_ttm=True)

        if mode == "minimal":
            return self.fetch_minimal(
                symbol,
                need_profile=True, need_quote=True, need_enterprise_value=True,
                need_ratios_ttm=True, need_key_metrics_ttm=True,
            )

        return self.fetch_all(symbol)

    def fetch_minimal(self, symbol: str, *,
                      need_profile: bool = True, need_quote: bool = True,
                      need_income_annual: bool = False, need_income_quarter: bool = False,
                      need_cashflow_quarter: bool = False, need_balance_annual: bool = False,
                      need_enterprise_value: bool = False, need_ratios_ttm: bool = False,
                      need_key_metrics_ttm: bool = True) -> Dict[str, Any]:

        if not self.enabled: return {}

        def _get(endpoint: str, params: Dict[str, Any]) -> Optional[Any]:
            return self._get_data(endpoint, symbol, params)

        data: Dict[str, Any] = {}

        # --- CIRCUIT BREAKER ---
        # If Profile or Quote fails, we stop IMMEDIATELY.
        failed_critical = False

        if need_profile:
            p = _get("profile", {})
            if p is None:
                failed_critical = True
            else:
                data["profile"] = p[0] if isinstance(p, list) and p else p

        if need_quote and not failed_critical:
            q = _get("quote", {})
            if q is None:
                failed_critical = True
            else:
                data["quote"] = q[0] if isinstance(q, list) and q else q

        if failed_critical:
            print(f"[FMP PROTECTION] Aborting fetch for {symbol} (Profile/Quote failed).")
            return {}
            # -----------------------

        if need_key_metrics_ttm:
            km = _get("key-metrics-ttm", {})
            data["key_metrics_ttm"] = km[0] if isinstance(km, list) and km else km

        if need_ratios_ttm:
            rt = _get("ratios-ttm", {})
            data["ratios_ttm"] = rt[0] if isinstance(rt, list) and rt else rt

        if need_enterprise_value:
            ev = _get("enterprise-values", {"limit": self._cap_limit(2)})
            data["enterprise_value"] = ev[0] if isinstance(ev, list) and ev else ev

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
        return self.fetch_minimal(
            symbol,
            need_profile=True, need_quote=True, need_enterprise_value=True,
            need_ratios_ttm=True, need_key_metrics_ttm=True,
            need_income_annual=True, need_income_quarter=self.use_quarterly,
            need_cashflow_quarter=self.use_quarterly, need_balance_annual=True
        )

    # -----------------------
    # Internal Logic (FIXED URL STRUCTURE)
    # -----------------------

    def _get_data(self, endpoint: str, symbol: str, params: Dict[str, Any]) -> Any:
        if not self.enabled: return None

        if self.max_requests > 0 and self.request_count >= self.max_requests:
            return None

        # --- FIX: PUT SYMBOL IN PATH ---
        # Correct: .../quote/AAPL?apikey=...
        # Wrong:   .../quote?symbol=AAPL&apikey=...
        base_url = f"https://financialmodelingprep.com/api/v3/{endpoint}/{symbol}"

        req_params = params.copy()
        req_params["apikey"] = self.api_key
        # Note: We do NOT put 'symbol' in req_params anymore.

        # Cache Key
        cache_str = f"{endpoint}_{symbol}_{json.dumps(params, sort_keys=True)}"
        cache_key = hashlib.sha256(cache_str.encode("utf-8")).hexdigest()

        # 1. READ CACHE
        cached = self._read_cache(cache_key)
        if cached is not None:
            if isinstance(cached, dict) and "_error" in cached:
                return None
            return cached

        # 2. NETWORK REQUEST
        for attempt in range(self.max_retries + 1):
            try:
                self.request_count += 1
                resp = self.session.get(base_url, params=req_params, timeout=self.timeout)

                # --- FATAL ERRORS ---
                if resp.status_code == 400:
                    self._write_cache(cache_key, {"_error": 400})
                    self._inc_error()
                    return None

                if resp.status_code in (403, 404):
                    self._write_cache(cache_key, {"_error": resp.status_code})
                    self._inc_error()
                    return None

                if resp.status_code == 401:
                    self._inc_error()
                    return None

                    # --- RETRYABLE ---
                if resp.status_code == 429:
                    time.sleep(1.0 * (attempt + 1))
                    continue

                if resp.status_code >= 500:
                    time.sleep(0.5 * (attempt + 1))
                    continue

                # --- SUCCESS ---
                if resp.status_code == 200:
                    data = resp.json()
                    self._write_cache(cache_key, data)
                    self._inc_success()
                    return data

            except Exception:
                time.sleep(0.5 * (attempt + 1))

        self._inc_error()
        return None

    def _read_cache(self, key: str) -> Optional[Any]:
        if not self.use_cache: return None
        try:
            path = os.path.join(self.cache_dir, f"{key}.json")
            if not os.path.exists(path): return None
            if time.time() - os.path.getmtime(path) > (self.cache_ttl_hours * 3600):
                return None
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return None

    def _write_cache(self, key: str, data: Any):
        if not self.use_cache or data is None: return
        try:
            path = os.path.join(self.cache_dir, f"{key}.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f)
        except:
            pass