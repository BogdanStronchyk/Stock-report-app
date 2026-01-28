
import math
from typing import Any, Dict, Optional, Tuple
from openpyxl.styles import PatternFill

from checklist_loader import parse_range_cell
from config import FILL_GREEN, FILL_YELLOW, FILL_RED, FILL_GRAY

# ----------------------------
# Threshold rating (existing)
# ----------------------------
def score_with_threshold_txt(
    value: Optional[float],
    green_txt: Any,
    yellow_txt: Any,
    red_txt: Any
) -> Tuple[str, PatternFill]:
    if value is None or (isinstance(value, float) and (math.isnan(value) or math.isinf(value))):
        return "NA", FILL_GRAY

    g = parse_range_cell(str(green_txt)) if green_txt is not None else None
    y = parse_range_cell(str(yellow_txt)) if yellow_txt is not None else None
    r = parse_range_cell(str(red_txt)) if red_txt is not None else None

    def in_rule(rule):
        if rule is None:
            return False
        lo, hi = rule
        if lo is None and hi is not None:
            return value < hi
        if lo is not None and hi is None:
            return value > lo
        if lo is not None and hi is not None:
            return lo <= value <= hi
        return False

    if in_rule(g):
        return "GREEN", FILL_GREEN
    if in_rule(y):
        return "YELLOW", FILL_YELLOW
    if in_rule(r):
        return "RED", FILL_RED

    return "NA", FILL_GRAY


# ----------------------------
# Scoring helpers (new)
# ----------------------------
POINTS = {"GREEN": 2, "YELLOW": 1, "RED": 0, "NA": 0}

CATEGORY_WEIGHTS = {
    "Valuation": 0.20,
    "Profitability": 0.25,
    "Balance Sheet": 0.25,
    "Growth": 0.15,
    "Risk": 0.15,
}


def rating_to_points(rating: str) -> int:
    return POINTS.get((rating or "NA").upper(), 0)


def compute_category_score_and_coverage(
    ratings_by_metric: Dict[str, str],
    weights_by_metric: Dict[str, float],
) -> Tuple[Optional[float], float]:
    # Returns (raw_score_pct or None if no scorable metrics, coverage_pct)
    #
    # - Raw score uses ONLY scorable metrics (rating != NA) in denominator.
    # - Coverage is % of total category weight that is scorable.
    possible_w = 0.0
    scorable_w = 0.0
    earned = 0.0
    max_earned = 0.0

    for metric, w in weights_by_metric.items():
        w = float(w or 0.0)
        if w <= 0:
            continue
        possible_w += w
        rating = (ratings_by_metric.get(metric) or "NA").upper()
        if rating == "NA":
            continue
        scorable_w += w
        pts = rating_to_points(rating)
        earned += pts * w
        max_earned += 2 * w

    coverage = 0.0 if possible_w == 0 else (scorable_w / possible_w) * 100.0
    if max_earned == 0:
        return None, coverage
    return (earned / max_earned) * 100.0, coverage


def adjusted_from_raw_and_coverage(raw_score_pct: Optional[float], coverage_pct: float) -> Optional[float]:
    if raw_score_pct is None:
        return None
    return float(raw_score_pct) * (float(coverage_pct) / 100.0)
