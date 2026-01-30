"""
scoring.py

Central Scoring Engine.
Calculates coverage-adjusted scores, applies metric weights, detects outliers,
and handles sector-specific nuances (Banks/REITs).
"""

import math
from typing import Any, Dict, List, Optional, Tuple
from openpyxl.styles import PatternFill

from checklist_loader import parse_range_cell, get_threshold_set
from config import FILL_GREEN, FILL_YELLOW, FILL_RED, FILL_GRAY

# ----------------------------
# Configuration
# ----------------------------
POINTS = {"GREEN": 2, "YELLOW": 1, "RED": 0, "NA": 0}

CATEGORY_WEIGHTS = {
    "Valuation": 0.20,
    "Profitability": 0.25,
    "Balance Sheet": 0.25,
    "Growth": 0.15,
    "Risk": 0.15,
}

# ----------------------------
# Helpers
# ----------------------------
def _metric_weight(category: str, metric: str) -> float:
    """Internal weights for specific metrics (Higher weight = more critical)."""
    m = (metric or "")
    cat = (category or "")
    w = 1.0

    if cat == "Valuation":
        if any(x in m for x in ["EV/FCF", "FCF Yield", "EV/EBIT"]): w = 1.6
        elif any(x in m for x in ["EV/EBITDA", "P/E"]): w = 1.3
        elif any(x in m for x in ["P/S", "EV/Gross Profit"]): w = 1.1

    elif cat == "Profitability":
        if "ROIC" in m: w = 1.8
        elif any(x in m for x in ["Operating Margin", "FCF Margin"]): w = 1.4
        elif any(x in m for x in ["Gross Margin", "Net Margin"]): w = 1.2

    elif cat == "Balance Sheet":
        if any(x in m for x in ["Net Debt / EBITDA", "Interest Coverage"]): w = 1.7
        elif any(x in m for x in ["Net Debt / FCF", "FCF / Interest"]): w = 1.4
        elif "Cash / Total Assets" in m: w = 0.8

    elif cat == "Growth":
        if any(x in m for x in ["FCF per Share", "Revenue per Share"]): w = 1.4
        elif "Revenue CAGR" in m: w = 1.2
        else: w = 0.9

    elif cat == "Risk":
        if any(x in m for x in ["Max Drawdown", "Realized Volatility"]): w = 1.4
        elif "Worst Weekly" in m: w = 1.3
        elif "Avg Daily" in m: w = 1.2
        elif "Market Cap" in m: w = 0.7

    return float(w)

def _extract_numeric_bounds(green_txt: Any, yellow_txt: Any, red_txt: Any) -> Tuple[Optional[float], Optional[float]]:
    rules = []
    for txt in (green_txt, yellow_txt, red_txt):
        if txt:
            r = parse_range_cell(str(txt))
            if r: rules.append(r)
    if not rules:
        return (None, None)
    lows = [lo for lo, hi in rules if lo is not None]
    highs = [hi for lo, hi in rules if hi is not None]
    return (min(lows) if lows else None, max(highs) if highs else None)

def _is_aberrant(metric_name: str, value: Any, low: Optional[float], high: Optional[float]) -> bool:
    """Detect data artifacts (e.g. P/E of 1,000,000 due to near-zero earnings)."""
    try:
        if value is None or isinstance(value, str): return False
        v = float(value)
        if math.isnan(v) or math.isinf(v): return True
    except Exception:
        return False

    # Skip magnitude checks for "Size" metrics
    if any(k in (metric_name or "") for k in ["Market Cap", "Enterprise Value", "Volume"]):
        return False

    # Logic: if value is > 50x the logic range, assume data artifact
    if low is None and high is None:
        if abs(v) > 1e15: return True
        return False

    ref = max([abs(x) for x in [low, high] if x is not None] + [1.0])
    MULT = 50.0

    # Calculate hard limits
    min_limit = (low if low is not None else -ref) - (MULT * ref)
    max_limit = (high if high is not None else ref) + (MULT * ref)

    return v < min_limit or v > max_limit

def _adjust_metric_requirements(category: str, metric: str, sector: str) -> bool:
    """
    Returns False if a metric should be IGNORED for this sector to avoid false failures.
    """
    s = (sector or "").upper()
    m = (metric or "")

    # 1. Financials (Banks/Insurance): No EBITDA, No Operating Margin (usually)
    if "FINANCIAL" in s or "BANK" in s:
        if "EBITDA" in m or "EV/" in m or "OPERATING MARGIN" in m.upper():
            return False

    # 2. REITs: No P/E (use FFO), EPS is misleading
    if "REIT" in s or "REAL ESTATE" in s:
        if "P/E" in m or "EPS" in m:
            return False

    return True

# ----------------------------
# Core Scoring
# ----------------------------
def score_with_threshold_txt(value: Optional[float], green_txt: Any, yellow_txt: Any, red_txt: Any) -> Tuple[str, PatternFill]:
    if value is None or (isinstance(value, float) and (math.isnan(value) or math.isinf(value))):
        return "NA", FILL_GRAY

    g = parse_range_cell(str(green_txt)) if green_txt else None
    y = parse_range_cell(str(yellow_txt)) if yellow_txt else None
    r = parse_range_cell(str(red_txt)) if red_txt else None

    def in_rule(rule, val):
        if not rule: return False
        lo, hi = rule
        if lo is None and hi is not None: return val < hi
        if lo is not None and hi is None: return val > lo
        return lo <= val <= hi

    if in_rule(g, value): return "GREEN", FILL_GREEN
    if in_rule(y, value): return "YELLOW", FILL_YELLOW
    if in_rule(r, value): return "RED", FILL_RED

    return "NA", FILL_GRAY

def compute_category_score_and_coverage(ratings: Dict[str, str], weights: Dict[str, float]) -> Tuple[Optional[float], float]:
    possible_w = 0.0
    scorable_w = 0.0
    earned = 0.0
    max_earned = 0.0

    for metric, w in weights.items():
        if w <= 0: continue
        possible_w += w
        rating = (ratings.get(metric) or "NA").upper()
        if rating == "NA": continue

        scorable_w += w
        pts = POINTS.get(rating, 0)
        earned += pts * w
        max_earned += 2 * w

    coverage = (scorable_w / possible_w * 100.0) if possible_w > 0 else 0.0
    score = (earned / max_earned * 100.0) if max_earned > 0 else None
    return score, coverage

def adjusted_from_raw_and_coverage(raw_score: Optional[float], coverage: float) -> Optional[float]:
    if raw_score is None: return None
    return float(raw_score) * (float(coverage) / 100.0)

def score_ticker(metrics: Dict[str, Any], thresholds: Dict[str, Any]) -> Dict[str, Any]:
    """
    Orchestrator: Takes raw metrics + checklist thresholds -> Returns fully calculated scores.
    """
    sector_bucket = metrics.get("Sector Bucket", "Default (All)")

    # Structure to return
    out = {
        "cat_adj": {},      # {Category: Score 0-100}
        "cat_cov": {},      # {Category: Coverage 0-100}
        "ratings": {},      # {Category: {Metric: Rating}}
        "fund_adj": 0.0,    # Overall Fundamental Score
        "flags": []         # Any structural red flags found
    }

    category_maps = ["Valuation", "Profitability", "Balance Sheet", "Growth", "Risk"]

    # Temp storage for weighted blend
    cat_adj_scores = {}

    for cat in category_maps:
        if cat not in thresholds: continue

        cat_ratings = {}
        cat_weights = {}

        for metric, thset in thresholds[cat].items():
            # --- PATCH: SECTOR FILTER ---
            # If this metric is nonsense for this sector (e.g. EBITDA for Banks), skip it entirely.
            if not _adjust_metric_requirements(cat, metric, sector_bucket):
                continue

            raw_val = metrics.get(metric)
            th = get_threshold_set(thresholds, cat, metric, sector_bucket)

            # 1. Outlier Check
            lo, hi = _extract_numeric_bounds(th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))
            if _is_aberrant(metric, raw_val, lo, hi):
                rating = "NA"
            else:
                rating, _ = score_with_threshold_txt(raw_val, th.get("green_txt"), th.get("yellow_txt"), th.get("red_txt"))

            cat_ratings[metric] = rating
            cat_weights[metric] = _metric_weight(cat, metric)

        # 2. Compute Category Score
        raw, cov = compute_category_score_and_coverage(cat_ratings, cat_weights)

        # 3. Apply Category Caps (Red Flags)
        final_raw = raw
        if cat == "Balance Sheet" and raw is not None:
             if any(("Net Debt" in k or "Interest" in k) and v == "RED" for k,v in cat_ratings.items()):
                 final_raw = min(raw, 60.0)

        # --- PATCH: DILUTION PENALTY ---
        # If Share Count growing > 3% CAGR, cap Growth Score
        if cat == "Growth":
             share_cagr = metrics.get("Share Count CAGR (3Y)")
             if share_cagr is not None and share_cagr > 3.0:
                 final_raw = min(final_raw or 100.0, 50.0)
                 out["flags"].append(f"⚠ Dilution Risk ({share_cagr:.1f}% CAGR)")

        # If High SBC, cap Profitability
        if cat == "Profitability":
             sbc_pct = metrics.get("SBC % of Market Cap (TTM)")
             if sbc_pct is not None and sbc_pct > 3.0:
                 final_raw = min(final_raw or 100.0, 60.0)
                 out["flags"].append(f"⚠ High SBC ({sbc_pct:.1f}%)")

        adj_score = adjusted_from_raw_and_coverage(final_raw, cov)

        out["ratings"][cat] = cat_ratings
        out["cat_cov"][cat] = cov
        out["cat_adj"][cat] = adj_score
        cat_adj_scores[cat] = adj_score

    # 4. Weighted Blend for Final Score
    acc = 0.0
    wsum = 0.0
    for cat, w in CATEGORY_WEIGHTS.items():
        v = cat_adj_scores.get(cat, 0.0) # Treat None as 0.0 for penalty
        v = 0.0 if v is None else v
        acc += v * w
        wsum += w

    out["fund_adj"] = (acc / wsum) if wsum > 0 else 0.0
    return out