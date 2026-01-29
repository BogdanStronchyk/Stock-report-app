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

    If you see HTTP 400/403 responses, the JSON body usually explains whether the
    issue is a plan restriction, an invalid symbol, or a bad/expired key.

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

        # Many plans restrict quarterly statements. Default to annual-only unless explicitly enabled.
        self.use_quarterly = (os.environ.get("FMP_USE_QUARTERLY", "0") or "0").strip().lower() in ("1", "true", "yes")

        # Statement pagination limits
        #
        # Some subscriptions restrict the allowed values for the `limit` parameter (commonly <= 5).
        # Keep this configurable but capped to 5 by default to avoid HTTP 402 "Premium Query Parameter" errors.
        self.statement_limit = self._safe_int(os.environ.get("FMP_STATEMENT_LIMIT", "5"), default=5)
        # Hard cap at 5 unless you change code (matches your current plan restriction).
        self.statement_limit = max(0, min(5, self.statement_limit))

    def _cap_limit(self, desired: int) -> Optional[int]:
        """Return a subscription-safe `limit` value.

        If statement_limit <= 0, omit the `limit` parameter entirely (API default).
        """
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
    def _stable_url(self, endpoint: str, params: Optional[Dict[str, Any]] = None) -> str:
        params = dict(params or {})
        params["apikey"] = self.api_key
        qs = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
        base = (os.environ.get("FMP_BASE_URL") or "https://financialmodelingprep.com/stable/").strip()
        base = base.rstrip("/")
        endpoint = (endpoint or "").lstrip("/")
        return f"{base}/{endpoint}?{qs}"

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
            # Avoid confusing state like last_status=200 with a stale last_error from earlier.
            self.last_error = None
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
                    # Successful call clears previous errors.
                    self.last_error = None
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

                # Best-effort parse of JSON error bodies (FMP often returns: {"Error Message": "..."}).
                msg = None
                if body:
                    try:
                        j = json.loads(body)
                        if isinstance(j, dict):
                            msg = j.get("Error Message") or j.get("error") or j.get("message")
                    except Exception:
                        msg = None

                if e.code in (401, 403):
                    base = f"HTTP {e.code} {'Unauthorized' if e.code == 401 else 'Forbidden'} from FMP"
                    if msg:
                        self.last_error = f"{base}: {msg}"
                    else:
                        self.last_error = (
                            f"{base}. Possible causes: invalid/expired API key, plan restriction, "
                            "or access blocked."
                        )
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

                if msg:
                    self.last_error = f"HTTP {e.code} error from FMP: {msg}"
                else:
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

        def _get(endpoint: str, params: Dict[str, Any]) -> Optional[Any]:
            return self._get_json(self._stable_url(endpoint, params))

        # Profile / Quote
        data["profile"] = _get("profile", {"symbol": symbol}) or [] if need_profile else []
        if need_quote:
            q = _get("quote", {"symbol": symbol})
            if q is None:
                # Some plans expose quote-short but not full quote.
                q = _get("quote-short", {"symbol": symbol})
            data["quote"] = q or []
        else:
            data["quote"] = []

        # Statements / enterprise values
        if need_income_annual:
            ia = _get("income-statement", {"symbol": symbol, "period": "annual", "limit": self._cap_limit(6)})
            if ia is None:
                ia = _get("income-statement", {"symbol": symbol, "limit": self._cap_limit(6)})
            data["income_a"] = ia or []
        else:
            data["income_a"] = []

        if need_income_quarter:
            iq = _get("income-statement", {"symbol": symbol, "period": "quarter", "limit": self._cap_limit(8)})
            if iq is None:
                # Fallback to annual if quarterly is restricted.
                iq = _get("income-statement", {"symbol": symbol, "period": "annual", "limit": self._cap_limit(8)})
            data["income_q"] = iq or []
        else:
            data["income_q"] = []

        if need_cashflow_quarter:
            cq = _get("cash-flow-statement", {"symbol": symbol, "period": "quarter", "limit": self._cap_limit(8)})
            if cq is None:
                cq = _get("cash-flow-statement", {"symbol": symbol, "period": "annual", "limit": self._cap_limit(8)})
            data["cashflow_q"] = cq or []
        else:
            data["cashflow_q"] = []

        if need_balance_annual:
            ba = _get("balance-sheet-statement", {"symbol": symbol, "period": "annual", "limit": self._cap_limit(2)})
            if ba is None:
                ba = _get("balance-sheet-statement", {"symbol": symbol, "limit": self._cap_limit(2)})
            data["balance_a"] = ba or []
        else:
            data["balance_a"] = []

        if need_enterprise_value:
            ev = _get("enterprise-values", {"symbol": symbol, "limit": self._cap_limit(2)})
            if ev is None:
                ev = _get("enterprise-values", {"symbol": symbol})
            data["enterprise_value"] = ev or []
        else:
            data["enterprise_value"] = []

        # TTM endpoints
        data["ratios_ttm"] = _get("ratios-ttm", {"symbol": symbol}) or [] if need_ratios_ttm else []
        data["key_metrics_ttm"] = _get("key-metrics-ttm", {"symbol": symbol}) or [] if need_key_metrics_ttm else []

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
        """Full bundle (kept for backwards compatibility).

        NOTE: This is stable-only (no legacy `/api/v3/...` calls).
        """
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
