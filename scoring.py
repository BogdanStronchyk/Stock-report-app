import math
from typing import Any, Optional, Tuple
from openpyxl.styles import PatternFill

from checklist_loader import parse_range_cell
from config import FILL_GREEN, FILL_YELLOW, FILL_RED, FILL_GRAY

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
