"""
eligibility.py

Data-driven eligibility gating.
Loads rules from 'eligibility_rules.json' to determine PASS/WATCH/FAIL status.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Any

# Map legacy internal mode names to new JSON keys
MODE_ALIASES = {
    "buy": "strict",
    "portfolio": "strict",
    "shortlist": "permissible",
    "screen": "permissible",
    "broad": "loose",
}

DEFAULT_RULES_FILE = "eligibility_rules.json"

# Fallback in case JSON is missing
FALLBACK_RULES = {
    "strict": {
        "min_overall_coverage": 70.0,
        "davf_policy": "enforce_allowed",
        "davf_list": ["GREEN", "YELLOW"]
    }
}

@dataclass(frozen=True)
class EligibilityResult:
    status: str          # PASS | WATCH | FAIL
    label: str           # ELIGIBLE | WATCH | INELIGIBLE
    overall_coverage_pct: float
    reasons: List[str]

    def reasons_text(self, max_items: int = 3) -> str:
        if not self.reasons:
            return ""
        items = self.reasons[: max_items if max_items > 0 else len(self.reasons)]
        suffix = "" if len(self.reasons) <= len(items) else f" (+{len(self.reasons) - len(items)} more)"
        return " | ".join(items) + suffix


def _load_rules() -> Dict[str, Any]:
    """Load JSON rules from disk or return fallback."""
    path = os.path.join(os.getcwd(), DEFAULT_RULES_FILE)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {DEFAULT_RULES_FILE}: {e}")
    return FALLBACK_RULES


# Cache rules in memory so we don't re-read file for every ticker
_CACHED_RULES = {}

def get_rule_set(mode: str) -> Dict[str, Any]:
    global _CACHED_RULES
    if not _CACHED_RULES:
        _CACHED_RULES = _load_rules()

    # Resolve alias (e.g. 'buy' -> 'strict')
    clean_mode = mode.strip().lower()
    target_key = MODE_ALIASES.get(clean_mode, clean_mode)

    # Return rules or strict fallback if key missing
    return _CACHED_RULES.get(target_key) or _CACHED_RULES.get("strict") or FALLBACK_RULES["strict"]


def evaluate_eligibility(
    *,
    mode: str,
    cat_adj: Dict[str, Optional[float]],
    cat_cov: Dict[str, float],
    category_ratings: Dict[str, Dict[str, str]],  # Unused by generic logic but kept for interface compat
    sector_bucket: Optional[str] = None,
    fund_adj: Optional[float] = None,
    reversal_total: Optional[float] = None,
    davf_label: Optional[str] = None,
) -> EligibilityResult:

    rules = get_rule_set(mode)
    reasons: List[str] = []

    # --- Helper to safely get floats ---
    def _f(x) -> Optional[float]:
        try:
            return None if x is None else float(x)
        except Exception:
            return None

    # --- 1. Calculate Overall Coverage ---
    # (Simple average of category coverages if explicit weights unavailable)
    cov_vals = [float(v) for v in (cat_cov or {}).values() if v is not None]
    overall_cov = sum(cov_vals) / max(1, len(cov_vals)) if cov_vals else 0.0

    # --- 2. Check Coverage Gates ---
    min_cov = float(rules.get("min_overall_coverage", 0.0))
    if overall_cov < min_cov:
        reasons.append(f"Low overall coverage ({overall_cov:.0f}% < {min_cov:.0f}%)")

    for cat, limit in rules.get("min_category_coverage", {}).items():
        curr = float(cat_cov.get(cat, 0.0))
        if curr < float(limit):
            reasons.append(f"Thin {cat} data ({curr:.0f}% < {limit:.0f}%)")

    # --- 3. Check Score Gates (Fundamental & Reversal) ---
    for score_key, limit in rules.get("min_total_scores", {}).items():
        val = None
        if score_key == "fund_adj": val = _f(fund_adj)
        elif score_key == "reversal_total": val = _f(reversal_total)

        if val is not None and val < float(limit):
            reasons.append(f"{score_key} too low ({val:.0f} < {limit})")

    # --- 4. Check Structural Category Scores ---
    # e.g., Reject if Balance Sheet score is < 15
    for cat, limit in rules.get("min_category_score", {}).items():
        val = _f(cat_adj.get(cat))
        if val is not None and val < float(limit):
            reasons.append(f"Weak {cat} structure ({val:.0f} < {limit})")

    # --- 5. DAVF (Downside Protection) Logic ---
    davf = (davf_label or "NA").upper().strip()
    policy = rules.get("davf_policy", "ignore")
    davf_list = set(rules.get("davf_list", []))

    if policy == "enforce_allowed":
        # Strict: Must be in the allowed list (e.g. GREEN/YELLOW)
        if davf not in davf_list:
            reasons.append(f"DAVF protection insufficient ({davf})")

    elif policy == "flag_watch":
        # Permissible: If in list (e.g. RED/NA), flag it but don't auto-fail yet
        if davf in davf_list:
            reasons.append(f"DAVF weak ({davf})")

    # --- Final Decision ---
    if not reasons:
        return EligibilityResult(status="PASS", label="ELIGIBLE", overall_coverage_pct=overall_cov, reasons=[])

    # Determining FAIL vs WATCH
    # If policy is strict, any reason is a FAIL.
    # If policy is permissible, we might allow 'watch' for soft failures.

    # We treat Coverage failures as hard FAILS usually,
    # but Score failures might be WATCH in looser modes.
    # For simplicity/safety:
    # Strict mode -> Fail closed (FAIL)
    # Permissible -> Fail closed on Coverage, Watch on Scores?

    # Implementation: If mode is Strict ('buy'), everything is a FAIL.
    if mode in ("strict", "buy", "portfolio"):
        return EligibilityResult(status="FAIL", label="INELIGIBLE", overall_coverage_pct=overall_cov, reasons=reasons)

    # In Permissible/Loose:
    # If coverage is the issue -> FAIL (bad data)
    # If score is the issue -> WATCH (bad stock, but readable)
    is_coverage_issue = any("coverage" in r.lower() or "data" in r.lower() for r in reasons)

    if is_coverage_issue:
        return EligibilityResult(status="FAIL", label="INELIGIBLE", overall_coverage_pct=overall_cov, reasons=reasons)

    return EligibilityResult(status="WATCH", label="WATCH", overall_coverage_pct=overall_cov, reasons=reasons)