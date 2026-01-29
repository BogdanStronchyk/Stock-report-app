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

    IMPORTANT: FMP has *two* URL styles:
      - Legacy v3 style: /api/v3/<endpoint>/<SYMBOL>?period=quarter&limit=...&apikey=...
      - Stable style:     /stable/<endpoint>?symbol=SYMBOL&apikey=...

    This client primarily uses the v3 style for statements and enterprise values,
    because those endpoints are clearly documented with the symbol in the path.
    For TTM ratios/key-metrics, we try v3 first and then stable as a fallback.

    Performance/budget features:
      - Optional on-disk cache (default ON, TTL 24h) to avoid repeat calls.
      - Optional request budget per run (FMP_MAX_REQUESTS) to prevent exploding call counts.
      - Simple retry/backoff on transient network errors and HTTP 429.
    """

    def __init__(self, api_key: Optional[str] = None, timeout: int = 25):
        self.api_key = (api_key or os.environ.get("FMP_API_KEY") or "").strip()
        self.timeout = timeout
        self.enabled = bool(self.api_key)

        self.request_count = 0
        self.debug = os.environ.get("FMP_DEBUG", "").strip().lower() in ("1", "true", "yes")
        self.last_error = None
        self.last_status = None

        # Budget controls (per run)
        # If set (int > 0), hard-stops further calls once request_count reaches the limit.
        self.max_requests = self._safe_int(os.environ.get("FMP_MAX_REQUESTS", ""), default=0)

        # Cache controls
        self.cache_enabled = (os.environ.get("FMP_USE_CACHE", "1") or "1").strip().lower() not in ("0", "false", "no")
        self.cache_ttl_hours = self._safe_float(os.environ.get("FMP_CACHE_TTL_HOURS", "24"), default=24.0)

        # Retry controls
        self.max_retries = self._safe_int(os.environ.get("FMP_MAX_RETRIES", "2"), default=2)  # extra attempts
        self.backoff_base = self._safe_float(os.environ.get("FMP_BACKOFF_BASE", "1.2"), default=1.2)

    @staticmethod
    def _safe_int(v: str, default: int = 0) -> int:
        try:
            v = (v or "").strip()
            if not v:
                return default
            return int(float(v))
        except Exception:
            return default

    @staticmethod
    def _safe_float(v: str, default: float = 0.0) -> float:
        try:
            v = (v or "").strip()
            if not v:
                return default
            return float(v)
        except Exception:
            return default

    # -----------------------
    # URL builders
    # -----------------------
    def _v3_url(self, endpoint: str, symbol: Optional[str] = None, params: Optional[Dict[str, Any]] = None) -> str:
        params = dict(params or {})
        params["apikey"] = self.api_key
        qs = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
        if symbol:
            return f"https://financialmodelingprep.com/api/v3/{endpoint}/{symbol}?{qs}"
        return f"https://financialmodelingprep.com/api/v3/{endpoint}?{qs}"

    def _stable_url(self, endpoint: str, params: Optional[Dict[str, Any]] = None) -> str:
        params = dict(params or {})
        params["apikey"] = self.api_key
        qs = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
        return f"https://financialmodelingprep.com/stable/{endpoint}?{qs}"

    # -----------------------
    # Cache (disk)
    # -----------------------
    def _cache_dir(self) -> str:
        return os.path.join(os.getcwd(), ".cache", "fmp")

    def _cache_ttl_seconds(self) -> int:
        return int(max(0.0, self.cache_ttl_hours) * 3600)

    def _cache_key(self, url: str) -> str:
        # Hash full URL (includes apikey) to a stable filename.
        return hashlib.sha256(url.encode("utf-8", errors="ignore")).hexdigest() + ".json"

    def _cache_get(self, url: str) -> Optional[Any]:
        if not self.cache_enabled:
            return None
        ttl = self._cache_ttl_seconds()
        if ttl <= 0:
            return None
        cdir = self._cache_dir()
        try:
            os.makedirs(cdir, exist_ok=True)
        except Exception:
            return None
        path = os.path.join(cdir, self._cache_key(url))
        try:
            st = os.stat(path)
            age = datetime.now().timestamp() - st.st_mtime
            if age > ttl:
                return None
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _cache_set(self, url: str, payload: Any) -> None:
        if not self.cache_enabled:
            return
        cdir = self._cache_dir()
        try:
            os.makedirs(cdir, exist_ok=True)
            path = os.path.join(cdir, self._cache_key(url))
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f)
        except Exception:
            # Cache must never break data fetch
            return

    # -----------------------
    # HTTP
    # -----------------------
    def _budget_allows(self) -> bool:
        return (self.max_requests <= 0) or (self.request_count < self.max_requests)

    def _sleep_backoff(self, attempt: int, retry_after: Optional[float] = None) -> None:
        if retry_after is not None and retry_after > 0:
            time.sleep(retry_after)
            return
        # Exponential-ish backoff
        delay = max(0.2, self.backoff_base ** max(0, attempt))
        time.sleep(min(15.0, delay))

    def _get_json(self, url: str) -> Optional[Any]:
        if not self.enabled:
            self.last_error = "FMP_API_KEY not set."
            return None

        if not self._budget_allows():
            self.last_error = "FMP request budget reached (FMP_MAX_REQUESTS)."
            return None

        cached = self._cache_get(url)
        if cached is not None:
            self.last_status = 200
            return cached

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) StockReportApp/1.0",
            "Accept": "application/json,text/plain,*/*",
            "Accept-Language": "en-US,en;q=0.9",
            "Connection": "close",
        }

        for attempt in range(0, max(0, self.max_retries) + 1):
            req = urllib.request.Request(url, headers=headers, method="GET")
            try:
                self.request_count += 1
                if self.debug:
                    print(f"[FMP] GET {url}")

                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    self.last_status = getattr(resp, "status", None)
                    raw = resp.read().decode("utf-8", errors="ignore")
                    if not raw:
                        return None
                    payload = json.loads(raw)
                    self._cache_set(url, payload)
                    return payload

            except urllib.error.HTTPError as e:
                self.last_status = getattr(e, "code", None)

                retry_after = None
                try:
                    ra = e.headers.get("Retry-After")
                    if ra is not None:
                        retry_after = float(str(ra).strip())
                except Exception:
                    retry_after = None

                body = ""
                try:
                    body = e.read().decode("utf-8", errors="ignore")
                except Exception:
                    pass

                if e.code == 403:
                    self.last_error = (
                        "HTTP 403 Forbidden from FMP. Possible causes: invalid/expired API key, "
                        "plan restriction for this endpoint, or access blocked."
                    )
                    return None
                if e.code == 401:
                    self.last_error = "HTTP 401 Unauthorized from FMP (API key invalid)."
                    return None

                if e.code == 429:
                    self.last_error = "HTTP 429 Too Many Requests (rate-limited)."
                    if attempt < self.max_retries:
                        if self.debug:
                            print(f"[FMP] 429 rate limit; sleeping (Retry-After={retry_after}) and retrying...")
                        self._sleep_backoff(attempt, retry_after=retry_after)
                        continue
                    return None

                # Other 5xx can be transient; retry a couple times
                if e.code in (500, 502, 503, 504) and attempt < self.max_retries:
                    self.last_error = f"HTTP {e.code} from FMP (transient). Retrying..."
                    if self.debug:
                        print(f"[FMP] {self.last_error}")
                    self._sleep_backoff(attempt)
                    continue

                self.last_error = f"HTTP {e.code} error from FMP. Response: {body[:200]}"
                if self.debug:
                    print(f"[FMP] ERROR {e.code}: {self.last_error}")
                return None

            except urllib.error.URLError as e:
                self.last_error = f"Network error: {e}"
                if attempt < self.max_retries:
                    if self.debug:
                        print(f"[FMP] {self.last_error} (retrying...)")
                    self._sleep_backoff(attempt)
                    continue
                if self.debug:
                    print(f"[FMP] ERROR: {self.last_error}")
                return None

            except Exception as e:
                self.last_error = str(e)
                if attempt < self.max_retries:
                    if self.debug:
                        print(f"[FMP] ERROR: {self.last_error} (retrying...)")
                    self._sleep_backoff(attempt)
                    continue
                if self.debug:
                    print(f"[FMP] ERROR: {self.last_error}")
                return None

        return None

    # -----------------------
    # Fetch bundle(s)
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
        """Fetch only the endpoints you need (saves calls + time).

        Returned keys match fetch_all() so downstream code can keep working.
        Missing sections are returned as empty lists.
        """
        if not self.enabled:
            return {}

        data: Dict[str, Any] = {}

        data["profile"] = self._get_json(self._v3_url("profile", symbol, {})) or [] if need_profile else []
        data["quote"] = self._get_json(self._v3_url("quote", symbol, {})) or [] if need_quote else []

        data["income_a"] = (
            self._get_json(self._v3_url("income-statement", symbol, {"period": "annual", "limit": 6})) or []
            if need_income_annual else []
        )
        data["income_q"] = (
            self._get_json(self._v3_url("income-statement", symbol, {"period": "quarter", "limit": 8})) or []
            if need_income_quarter else []
        )
        data["cashflow_q"] = (
            self._get_json(self._v3_url("cash-flow-statement", symbol, {"period": "quarter", "limit": 8})) or []
            if need_cashflow_quarter else []
        )
        data["balance_a"] = (
            self._get_json(self._v3_url("balance-sheet-statement", symbol, {"period": "annual", "limit": 2})) or []
            if need_balance_annual else []
        )
        data["enterprise_value"] = (
            self._get_json(self._v3_url("enterprise-values", symbol, {"limit": 2})) or []
            if need_enterprise_value else []
        )

        if need_ratios_ttm:
            ratios = self._get_json(self._v3_url("ratios-ttm", symbol, {}))
            if ratios is None:
                ratios = self._get_json(self._stable_url("ratios-ttm", {"symbol": symbol}))
            data["ratios_ttm"] = ratios or []
        else:
            data["ratios_ttm"] = []

        if need_key_metrics_ttm:
            km = self._get_json(self._v3_url("key-metrics-ttm", symbol, {}))
            if km is None:
                km = self._get_json(self._stable_url("key-metrics-ttm", {"symbol": symbol}))
            data["key_metrics_ttm"] = km or []
        else:
            data["key_metrics_ttm"] = []

        return data

    def fetch_bundle(self, symbol: str, *, mode: str = "full") -> Dict[str, Any]:
        """Budget-aware fetch wrapper.

        mode:
          - off         : returns {}
          - conditional : quote + profile (+ key_metrics_ttm). No statements.
          - minimal     : quote + profile + key_metrics_ttm + ratios_ttm + enterprise_value. No statements.
          - full        : original fetch_all() bundle (statements + ratios + key metrics).
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

    def fetch_all(self, symbol: str) -> Dict[str, Any]:
        """Original 'full' bundle (kept for backwards compatibility)."""
        if not self.enabled:
            return {}

        data: Dict[str, Any] = {}

        # Profile/Quote (v3) for market cap/volume/price fallbacks
        data["profile"] = self._get_json(self._v3_url("profile", symbol, {})) or []
        data["quote"] = self._get_json(self._v3_url("quote", symbol, {})) or []

        # Statements (v3 format with symbol in path)
        data["income_a"] = self._get_json(self._v3_url("income-statement", symbol, {"period": "annual", "limit": 6})) or []
        data["income_q"] = self._get_json(self._v3_url("income-statement", symbol, {"period": "quarter", "limit": 8})) or []
        data["cashflow_q"] = self._get_json(self._v3_url("cash-flow-statement", symbol, {"period": "quarter", "limit": 8})) or []
        data["balance_a"] = self._get_json(self._v3_url("balance-sheet-statement", symbol, {"period": "annual", "limit": 2})) or []

        # Enterprise values (v3)
        data["enterprise_value"] = self._get_json(self._v3_url("enterprise-values", symbol, {"limit": 2})) or []

        # TTM ratios + key metrics:
        # Try v3 first (symbol in path) then stable (symbol query param).
        ratios_v3 = self._get_json(self._v3_url("ratios-ttm", symbol, {}))
        if ratios_v3 is None:
            ratios_v3 = self._get_json(self._stable_url("ratios-ttm", {"symbol": symbol}))
        data["ratios_ttm"] = ratios_v3 or []

        km_v3 = self._get_json(self._v3_url("key-metrics-ttm", symbol, {}))
        if km_v3 is None:
            km_v3 = self._get_json(self._stable_url("key-metrics-ttm", {"symbol": symbol}))
        data["key_metrics_ttm"] = km_v3 or []

        return data
