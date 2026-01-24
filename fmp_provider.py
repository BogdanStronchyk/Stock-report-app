
import os
import json
import urllib.parse
import urllib.request
import urllib.error
from typing import Any, Dict, Optional


class FMPClient:
    """
    Financial Modeling Prep fallback client.

    IMPORTANT: FMP has *two* URL styles:
      - Legacy v3 style: /api/v3/<endpoint>/<SYMBOL>?period=quarter&limit=...&apikey=...
      - Stable style:     /stable/<endpoint>?symbol=SYMBOL&apikey=...

    This client primarily uses the v3 style for statements and enterprise values,
    because those endpoints are clearly documented with the symbol in the path.
    For TTM ratios/key-metrics, we try v3 first and then stable as a fallback.
    """


    def __init__(self, api_key: Optional[str] = None, timeout: int = 25):
        self.api_key = (api_key or os.environ.get("FMP_API_KEY") or "").strip()
        self.timeout = timeout
        self.enabled = bool(self.api_key)

        self.request_count = 0
        self.debug = os.environ.get("FMP_DEBUG", "").strip().lower() in ("1", "true", "yes")
        self.last_error = None
        self.last_status = None

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
    # HTTP
    # -----------------------
    def _get_json(self, url: str) -> Optional[Any]:
        if not self.enabled:
            self.last_error = "FMP_API_KEY not set."
            return None

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) StockReportApp/1.0",
            "Accept": "application/json,text/plain,*/*",
            "Accept-Language": "en-US,en;q=0.9",
            "Connection": "close",
        }
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
                return json.loads(raw)

        except urllib.error.HTTPError as e:
            self.last_status = getattr(e, "code", None)
            body = ""
            try:
                body = e.read().decode("utf-8", errors="ignore")
            except Exception:
                pass

            if e.code == 403:
                self.last_error = (
                    "HTTP 403 Forbidden from FMP. Possible causes: invalid/expired API key, "
                    "plan restriction for this endpoint, or access blocked. "
                    "Try the same URL in a browser to confirm."
                )
            elif e.code == 401:
                self.last_error = "HTTP 401 Unauthorized from FMP (API key invalid)."
            elif e.code == 429:
                self.last_error = "HTTP 429 Too Many Requests (rate-limited)."
            else:
                self.last_error = f"HTTP {e.code} error from FMP. Response: {body[:200]}"

            if self.debug:
                print(f"[FMP] ERROR {e.code}: {self.last_error}")
            return None

        except Exception as e:
            self.last_error = str(e)
            if self.debug:
                print(f"[FMP] ERROR: {self.last_error}")
            return None

    # -----------------------
    # Fetch bundle
    # -----------------------
    def fetch_all(self, symbol: str) -> Dict[str, Any]:
        """
        fetch a small bundle of endpoints used for fallback fills.
        """
        if not self.enabled:
            return {}

        data: Dict[str, Any] = {}

        # Statements (v3 format with symbol in path)
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
