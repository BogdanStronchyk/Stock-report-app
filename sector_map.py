from typing import Optional

def map_sector(yahoo_sector: Optional[str], industry: Optional[str]) -> str:
    s = (yahoo_sector or "").lower()
    ind = (industry or "").lower()

    if "reit" in ind or "real estate" in s or "reit" in s:
        return "REITs"
    if "financial" in s or "bank" in ind or "capital markets" in ind:
        return "Financials (Banks)"
    if "utilities" in s or "telecom" in ind or "telecommunication" in ind:
        return "Utilities/Telecom"
    if "technology" in s or "software" in ind or "semiconductor" in ind:
        return "Software/Tech"
    if "healthcare" in s or "pharma" in ind or "biotech" in ind:
        return "Healthcare/Pharma"
    if "energy" in s or "materials" in s or "oil" in ind or "gas" in ind or "mining" in ind:
        return "Energy/Materials"
    if "consumer defensive" in s or "consumer staples" in s or "staples" in ind:
        return "Consumer Staples"
    if "consumer cyclical" in s or "consumer discretionary" in s or "retail" in ind:
        return "Consumer Discretionary"
    if "industrials" in s or "aerospace" in ind or "machinery" in ind:
        return "Industrials"

    return "Default (All)"
