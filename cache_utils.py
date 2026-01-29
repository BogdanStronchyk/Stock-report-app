"""
cache_utils.py

Small, dependency-light disk cache + concurrency limiter for IO-bound finance APIs.

Goals:
- Minimize repeat Yahoo/FMP calls within and across runs.
- Keep implementation simple and robust (cache must never break a run).
- Allow you to tune behavior via env vars.

Env vars (optional):
  YF_USE_CACHE=1|0
  YF_CACHE_TTL_HOURS=12
  YF_HTTP_CONCURRENCY=6          # semaphore limiting concurrent Yahoo requests
"""
from __future__ import annotations

import hashlib
import json
import os
import pickle
import threading
from datetime import datetime
from typing import Any, Callable, Optional, Tuple


def _safe_float(v: str, default: float) -> float:
    try:
        v = (v or "").strip()
        return default if not v else float(v)
    except Exception:
        return default


def _safe_int(v: str, default: int) -> int:
    try:
        v = (v or "").strip()
        return default if not v else int(float(v))
    except Exception:
        return default


class DiskCache:
    """Tiny on-disk cache with TTL.

    - JSON for dict/list payloads.
    - Pickle for arbitrary Python objects (pandas DataFrames, etc.)
    """

    def __init__(self, namespace: str, *, ttl_hours: float = 12.0, enabled: bool = True):
        self.namespace = (namespace or "cache").strip()
        self.ttl_hours = float(ttl_hours or 0.0)
        self.enabled = bool(enabled) and self.ttl_hours > 0

    def _dir(self) -> str:
        return os.path.join(os.getcwd(), ".cache", self.namespace)

    def _ttl_seconds(self) -> float:
        return max(0.0, float(self.ttl_hours)) * 3600.0

    @staticmethod
    def _key_to_fname(key: str) -> str:
        h = hashlib.sha256(key.encode("utf-8", errors="ignore")).hexdigest()
        return h

    def _path(self, key: str, ext: str) -> str:
        return os.path.join(self._dir(), self._key_to_fname(key) + ext)

    def _fresh(self, path: str) -> bool:
        try:
            age = datetime.now().timestamp() - os.stat(path).st_mtime
            return age <= self._ttl_seconds()
        except Exception:
            return False

    def get_json(self, key: str) -> Optional[Any]:
        if not self.enabled:
            return None
        path = self._path(key, ".json")
        if not self._fresh(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def set_json(self, key: str, payload: Any) -> None:
        if not self.enabled:
            return
        try:
            os.makedirs(self._dir(), exist_ok=True)
            with open(self._path(key, ".json"), "w", encoding="utf-8") as f:
                json.dump(payload, f)
        except Exception:
            return

    def get_pickle(self, key: str) -> Optional[Any]:
        if not self.enabled:
            return None
        path = self._path(key, ".pkl")
        if not self._fresh(path):
            return None
        try:
            with open(path, "rb") as f:
                return pickle.load(f)
        except Exception:
            return None

    def set_pickle(self, key: str, payload: Any) -> None:
        if not self.enabled:
            return
        try:
            os.makedirs(self._dir(), exist_ok=True)
            with open(self._path(key, ".pkl"), "wb") as f:
                pickle.dump(payload, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception:
            return


# -----------------------
# Yahoo IO concurrency cap
# -----------------------
_YF_HTTP_CONCURRENCY = _safe_int(os.environ.get("YF_HTTP_CONCURRENCY", "6"), default=6)
_YF_SEM = threading.BoundedSemaphore(max(1, _YF_HTTP_CONCURRENCY))


def yf_call(fn: Callable[[], Any]) -> Any:
    """Run a Yahoo/yfinance network call under a concurrency semaphore."""
    with _YF_SEM:
        return fn()


def yf_cache_settings() -> Tuple[bool, float]:
    enabled = (os.environ.get("YF_USE_CACHE", "1") or "1").strip().lower() not in ("0", "false", "no")
    ttl = _safe_float(os.environ.get("YF_CACHE_TTL_HOURS", "12"), default=12.0)
    return enabled, ttl
