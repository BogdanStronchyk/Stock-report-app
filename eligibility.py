"""Eligibility gating for index-scale screening.

This module adds a *decision eligibility layer* on top of your existing:
- per-metric ratings (GREEN/YELLOW/RED/NA)
- category adjusted scores (already coverage-adjusted)
- sector buckets (already checklist-sector aware)
- DAVF downside protection
- reversal confirmation

Why this exists:
Coverage-adjusted scores reduce inflation, but they do not *prevent* a ticker with
thin critical coverage from ranking highly in a large universe. Eligibility
gates enforce "fail-closed" behavior for automated screening.

You can tune thresholds in DEFAULT_RULES below without changing other code.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


@dataclass(frozen=True)
class EligibilityResult:
    """Eligibility decision for screening/shortlisting."""
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


# ---- Tunable rules ----
DEFAULT_RULES = {
    # For continuous screening: keep the funnel wide, but fail-closed on thin *core* coverage.
    "shortlist": {
        "min_overall_cov": 60.0,
        "min_cat_cov": {
            "Valuation": 45.0,
            "Profitability": 45.0,
            "Balance Sheet": 45.0,
            "Risk": 35.0,
        },
        # Structural red flags: if the model can score them and they're terrible, do not shortlist.
        "min_cat_adj": {
            "Balance Sheet": 15.0,
            "Risk": 15.0,
        },
        # DAVF: allow, but degrade to WATCH if uncertain/no protection.
        "davf_watch": {"NA", "RED"},
    },

    # For a stricter "buy candidate" filter: narrower funnel.
    "buy": {
        "min_overall_cov": 70.0,
        "min_cat_cov": {
            "Valuation": 55.0,
            "Profitability": 55.0,
            "Balance Sheet": 55.0,
            "Risk": 45.0,
        },
        "min_scores": {
            "fund_adj": 60.0,
            "reversal_total": 55.0,
        },
        "davf_allowed": {"GREEN", "YELLOW"},
    },
}



# Keep in sync with scoring.CATEGORY_WEIGHTS (duplicated here to avoid import cycles).
CATEGORY_WEIGHTS = {
    "Valuation": 0.20,
    "Profitability": 0.25,
    "Balance Sheet": 0.25,
    "Growth": 0.15,
    "Risk": 0.15,
}

# ---- Optional sector anchor checks (non-breaking) ----
# These only activate when a matching anchor metric exists in the checklist for that ticker.
SECTOR_ANCHORS = {
    # Typical bank/financial anchor concepts.
    "FINANCIALS": {
        "Valuation": ["P/B", "PRICE/BOOK", "P/TBV", "TANGIBLE BOOK"],
        "Profitability": ["ROE", "NET INTEREST", "NIM"],
    }
}


def _norm_sector_bucket(sector_bucket: Optional[str]) -> str:
    s = (sector_bucket or "").upper()
    if "FINANCIAL" in s or "BANK" in s:
        return "FINANCIALS"
    return ""


def _anchor_missing(
    sector_key: str,
    category: str,
    category_ratings: Dict[str, str],
) -> bool:
    patterns = SECTOR_ANCHORS.get(sector_key, {}).get(category, [])
    if not patterns:
        return False

    # Only enforce if the checklist *contains* at least one anchor.
    present = []
    scorable = []
    for metric, rating in (category_ratings or {}).items():
        m = (metric or "").upper()
        if any(p in m for p in patterns):
            present.append(metric)
            if (rating or "NA").upper() != "NA":
                scorable.append(metric)

    return bool(present) and not bool(scorable)


def evaluate_eligibility(
    *,
    mode: str,
    cat_adj: Dict[str, Optional[float]],
    cat_cov: Dict[str, float],
    category_ratings: Dict[str, Dict[str, str]],
    sector_bucket: Optional[str] = None,
    fund_adj: Optional[float] = None,
    reversal_total: Optional[float] = None,
    davf_label: Optional[str] = None,
) -> EligibilityResult:
    """Return an EligibilityResult for the given ticker.

    - mode: "shortlist" (default screener) or "buy" (strict candidate filter)
    - cat_adj: adjusted category scores (0-100) or None
    - cat_cov: category coverage % (0-100)
    - category_ratings: {category -> {metric -> rating}}
    """
    rules = DEFAULT_RULES.get(mode) or DEFAULT_RULES["shortlist"]

    def _f(x) -> Optional[float]:
        try:
            return None if x is None else float(x)
        except Exception:
            return None

    # Weighted overall coverage by category importance.
    # Uses CATEGORY_WEIGHTS where available; falls back to equal-weight if not.
    total_w = 0.0
    weighted = 0.0
    for cat, cov in (cat_cov or {}).items():
        try:
            c = float(cov or 0.0)
        except Exception:
            c = 0.0
        w = float(CATEGORY_WEIGHTS.get(cat, 0.0) or 0.0)
        if w <= 0:
            # If a category isn't recognized, treat it with a small equal weight later.
            continue
        weighted += c * w
        total_w += w

    if total_w > 0:
        overall_cov = weighted / total_w
    else:
        cov_vals = [float(v) for v in (cat_cov or {}).values() if v is not None]
        overall_cov = sum(cov_vals) / len(cov_vals) if cov_vals else 0.0

    reasons: List[str] = []

    # Overall coverage gate
    if overall_cov < float(rules.get("min_overall_cov", 0.0)):
        reasons.append(f"Low overall scorable coverage ({overall_cov:.0f}%)")

    # Category coverage gates
    for cat, thr in (rules.get("min_cat_cov") or {}).items():
        cov = float(cat_cov.get(cat, 0.0) or 0.0)
        if cov < float(thr):
            reasons.append(f"{cat} coverage {cov:.0f}% < {float(thr):.0f}%")

    # Category structural red flag gates (only if score exists)
    for cat, thr in (rules.get("min_cat_adj") or {}).items():
        s = _f(cat_adj.get(cat))
        if s is not None and s < float(thr):
            reasons.append(f"{cat} adjusted score {s:.0f}% < {float(thr):.0f}%")

    # DAVF handling
    d = (davf_label or "NA").upper().strip()
    if mode == "buy":
        allowed = set(rules.get("davf_allowed") or set())
        if allowed and d not in allowed:
            reasons.append(f"DAVF not acceptable for buy (DAVF={d})")
    else:
        watch = set(rules.get("davf_watch") or set())
        if d in watch:
            reasons.append(f"DAVF weak/uncertain (DAVF={d})")

    # Buy-mode minimum score gates
    if mode == "buy":
        mins = rules.get("min_scores") or {}
        fmin = float(mins.get("fund_adj", 0.0))
        rmin = float(mins.get("reversal_total", 0.0))

        f = _f(fund_adj)
        r = _f(reversal_total)

        if f is not None and f < fmin:
            reasons.append(f"Fund checklist {f:.0f}% < {fmin:.0f}%")
        if r is not None and r < rmin:
            reasons.append(f"Reversal {r:.0f}% < {rmin:.0f}%")

    # Optional sector anchors (only when anchors exist in the checklist)
    sector_key = _norm_sector_bucket(sector_bucket)
    if sector_key:
        for cat in ("Valuation", "Profitability"):
            if _anchor_missing(sector_key, cat, (category_ratings or {}).get(cat, {}) or {}):
                reasons.append(f"{cat}: missing sector anchor metric for {sector_key}")

    # Decide PASS/WATCH/FAIL
    if not reasons:
        return EligibilityResult(status="PASS", label="ELIGIBLE", overall_coverage_pct=float(overall_cov), reasons=[])

    # Fail vs watch:
    # - If overall coverage is low OR a core category coverage fails -> FAIL
    # - Otherwise -> WATCH
    core_fail = False
    if overall_cov < float(rules.get("min_overall_cov", 0.0)):
        core_fail = True
    for cat, thr in (rules.get("min_cat_cov") or {}).items():
        cov = float(cat_cov.get(cat, 0.0) or 0.0)
        if cov < float(thr):
            core_fail = True

    if mode == "buy":
        # For buy candidates, any reason is a FAIL.
        core_fail = True

    if core_fail:
        return EligibilityResult(status="FAIL", label="INELIGIBLE", overall_coverage_pct=float(overall_cov), reasons=reasons)

    return EligibilityResult(status="WATCH", label="WATCH", overall_coverage_pct=float(overall_cov), reasons=reasons)
