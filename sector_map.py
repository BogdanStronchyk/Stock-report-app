from typing import Optional
from typing import Optional

def map_sector(sector: Optional[str], industry: Optional[str]) -> str:
    """
    Maps yfinance sector/industry strings to our specific checklist buckets.
    Buckets:
      - Software/Tech
      - Financials (Banks)
      - REITs
      - Energy/Materials
      - Industrials
      - Healthcare/Pharma
      - Consumer Staples
      - Consumer Discretionary
      - Utilities/Telecom
      - Default (All)
    """
    s = (sector or "").lower()
    i = (industry or "").lower()

    # 1. REITs (High priority: often tagged as Real Estate sector)
    if "reit" in s or "reit" in i or "real estate" in s:
        return "REITs"

    # 2. Software/Tech
    # "Technology", "Communication Services" often overlap for tech stocks
    if s in ["technology", "tech"] or "software" in i or "semicon" in i or "information" in s:
        return "Software/Tech"
    # Communication services can be tech (Google/Meta) or Telecom.
    # We'll try to distinguish via industry, but often mapped to Tech if not legacy telecom.
    if s == "communication services":
        if "telecom" in i or "telephone" in i:
            return "Utilities/Telecom"
        return "Software/Tech"

    # 3. Financials
    if s == "financial services" or s == "financials":
        if "bank" in i or "credit" in i or "capital" in i or "insurance" in i:
            return "Financials (Banks)"
        # Fintech often falls here too, could map to Tech if desired, but default to Financials
        return "Financials (Banks)"

    # 4. Healthcare
    if s == "healthcare" or "pharm" in i or "biotech" in i or "medical" in i:
        return "Healthcare/Pharma"

    # 5. Energy/Materials
    if s in ["energy", "basic materials"] or "oil" in i or "gas" in i or "mining" in i or "metal" in i:
        return "Energy/Materials"

    # 6. Industrials
    if s == "industrials" or "aerospace" in i or "defense" in i or "machinery" in i or "transport" in i:
        return "Industrials"

    # 7. Utilities
    if s == "utilities":
        return "Utilities/Telecom"

    # 8. Consumer Staples
    if s == "consumer defensive" or "staples" in s:
        return "Consumer Staples"

    # 9. Consumer Discretionary
    if s == "consumer cyclical" or "discretionary" in s or "retail" in i or "auto" in i:
        return "Consumer Discretionary"

    # Default
    return "Default (All)"


# ==========================================
# NEW: Benchmark Mapping for Relative Return
# ==========================================
BENCHMARK_MAP = {
    "Software/Tech": "XLK",
    "Financials (Banks)": "XLF",
    "REITs": "XLRE",
    "Energy/Materials": "XLE",  # Or XLB for Materials, but XLE is a good proxy for 'old economy' resources
    "Industrials": "XLI",
    "Healthcare/Pharma": "XLV",
    "Consumer Staples": "XLP",
    "Consumer Discretionary": "XLY",
    "Utilities/Telecom": "XLU",
    "Default (All)": "SPY"
}

def get_sector_benchmark(sector_bucket: str) -> str:
    """Returns the ETF ticker for the given sector bucket to measure relative performance."""
    return BENCHMARK_MAP.get(sector_bucket, "SPY")
