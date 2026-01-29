"""Lightweight .env loader (no external deps).

Usage (add at the very top of main.py, before other imports that read env vars):
    from env_loader import load_env
    load_env()  # loads .env into os.environ (does not override existing env vars)

It supports simple KEY=VALUE lines and ignores blank lines / comments (# ...).
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional


def load_env(path: Optional[str] = None, *, override: bool = False) -> None:
    env_path = Path(path) if path else Path(os.getcwd()) / ".env"
    if not env_path.exists() or not env_path.is_file():
        return

    try:
        text = env_path.read_text(encoding="utf-8")
    except Exception:
        return

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if not key:
            continue

        if override or (key not in os.environ):
            os.environ[key] = value
