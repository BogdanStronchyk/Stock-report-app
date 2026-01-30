import os
import csv
import pandas as pd
import ssl

# Bypass SSL verification for Wikipedia scraping
ssl._create_default_https_context = ssl._create_unverified_context

OUTPUT_DIR = "universes"
os.makedirs(OUTPUT_DIR, exist_ok=True)

print(f"--- 🌍 Generating Global Indices in ./{OUTPUT_DIR}/ ---")

def save_csv(filename, tickers):
    """Saves a list of tickers to a standard CSV format."""
    path = os.path.join(OUTPUT_DIR, filename)
    tickers = sorted(list(set(tickers)))
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Ticker"])
        for t in tickers:
            writer.writerow([t])
    print(f"✅ Generated {filename} ({len(tickers)} tickers)")

# ==========================================
# 1. THE AMERICAS
# ==========================================
print("\n--- 🌎 THE AMERICAS ---")

# US: LIVE SCRAPE
try:
    print("⏳ Scraping S&P 500...")
    sp500 = pd.read_html('https://en.wikipedia.org/wiki/List_of_S%26P_500_companies')[0]
    save_csv("AMERICA_US_SP500.csv", sp500['Symbol'].str.replace('.', '-', regex=False).tolist())
except Exception as e: print(f"❌ SP500 Error: {e}")

try:
    print("⏳ Scraping Nasdaq 100...")
    tables = pd.read_html('https://en.wikipedia.org/wiki/Nasdaq-100')
    ndx_table = next(t for t in tables if 'Ticker' in t.columns or 'Symbol' in t.columns)
    col = 'Ticker' if 'Ticker' in ndx_table.columns else 'Symbol'
    save_csv("AMERICA_US_NASDAQ100.csv", ndx_table[col].str.replace('.', '-', regex=False).tolist())
except Exception as e: print(f"❌ Nasdaq Error: {e}")

# CANADA (TSX 60 - Top Constituents) - Suffix: .TO
tsx_top = [
    "RY.TO", "TD.TO", "SHOP.TO", "ENB.TO", "CNR.TO", "CP.TO", "CNQ.TO", "BMO.TO",
    "BNS.TO", "ATD.TO", "TRI.TO", "CSU.TO", "BCE.TO", "TRP.TO", "CM.TO", "MFC.TO",
    "GIB-A.TO", "PPL.TO", "POW.TO", "QSR.TO", "WCN.TO", "TECK-B.TO", "FTS.TO",
    "AEM.TO", "SLF.TO", "DOL.TO", "T.TO", "WPM.TO", "NA.TO", "MRU.TO", "L.TO"
]
save_csv("AMERICA_CANADA_TSX.csv", tsx_top)

# BRAZIL (Ibovespa Top) - Suffix: .SA
brazil_top = [
    "VALE3.SA", "PETR4.SA", "ITUB4.SA", "BBDC4.SA", "PETR3.SA", "BBAS3.SA", "ABEV3.SA",
    "WEGE3.SA", "RENT3.SA", "BPAC11.SA", "ITSA4.SA", "SUZB3.SA", "HAPV3.SA", "RDOR3.SA",
    "JBSS3.SA", "GGBR4.SA", "RAIL3.SA", "CSNA3.SA", "PRIO3.SA", "VBBR3.SA", "RADL3.SA"
]
save_csv("AMERICA_BRAZIL_BOVESPA.csv", brazil_top)


# ==========================================
# 2. EUROPE
# ==========================================
print("\n--- 🇪🇺 EUROPE ---")

# UK (FTSE 100 Top) - Suffix: .L
ftse_top = [
    "AZN.L", "SHEL.L", "HSBA.L", "ULVR.L", "BP.L", "DGE.L", "RIO.L", "REL.L", "GSK.L",
    "GLEN.L", "BATS.L", "LSEG.L", "CNA.L", "NG.L", "CPG.L", "LLOY.L", "BARC.L", "PRU.L",
    "VOD.L", "RR.L", "HLN.L", "AAL.L", "NWG.L", "STAN.L", "TSCO.L", "EXPN.L", "SSE.L"
]
save_csv("EUROPE_UK_FTSE100.csv", ftse_top)

# GERMANY (DAX 40) - Suffix: .DE
dax_top = [
    "SAP.DE", "SIE.DE", "AIR.DE", "ALV.DE", "DTE.DE", "VOW3.DE", "BMW.DE", "MBG.DE",
    "BAS.DE", "ADS.DE", "DHL.DE", "IFX.DE", "MUV2.DE", "DB1.DE", "EOAN.DE", "BEI.DE",
    "DBK.DE", "RWE.DE", "BAYN.DE", "HEN3.DE", "VNA.DE", "CON.DE", "DTG.DE", "HEI.DE"
]
save_csv("EUROPE_GERMANY_DAX.csv", dax_top)

# FRANCE (CAC 40 Top) - Suffix: .PA
cac_top = [
    "MC.PA", "OR.PA", "TTE.PA", "SAN.PA", "AIR.PA", "SU.PA", "AI.PA", "RMS.PA",
    "BNP.PA", "EL.PA", "KER.PA", "DG.PA", "CS.PA", "DSY.PA", "BN.PA", "STLAM.PA",
    "GLE.PA", "ACA.PA", "ORA.PA", "LR.PA", "CAP.PA", "SAF.PA", "VIE.PA", "ENGI.PA"
]
save_csv("EUROPE_FRANCE_CAC40.csv", cac_top)


# ==========================================
# 3. MIDDLE EAST & AFRICA
# ==========================================
print("\n--- 🌍 MIDDLE EAST & AFRICA ---")

# SOUTH AFRICA (JSE Top 40) - Suffix: .JO
# Note: Check these occasionally, JSE tickers change.
jse_top = [
    "NPN.JO", "PRX.JO", "FSR.JO", "SBK.JO", "AGL.JO", "SOL.JO", "GFI.JO", "MTN.JO",
    "VOD.JO", "ABG.JO", "NED.JO", "CPI.JO", "DSY.JO", "BTI.JO", "IMP.JO", "SHP.JO",
    "AMS.JO", "REM.JO", "WHL.JO", "EXX.JO", "NPH.JO", "ANG.JO", "SSW.JO", "KIO.JO"
]
save_csv("AFRICA_SOUTH_AFRICA_JSE.csv", jse_top)

# SAUDI ARABIA (Tadawul Top) - Suffix: .SR
tadawul_top = [
    "1120.SR", "2222.SR", "1180.SR", "2010.SR", "7010.SR", "1150.SR", "2280.SR",
    "1211.SR", "2290.SR", "4190.SR", "1140.SR", "1010.SR", "2350.SR", "4031.SR",
    "2310.SR", "1060.SR", "4200.SR", "1302.SR", "2020.SR", "7200.SR"
]
save_csv("MIDEAST_SAUDI_TADAWUL.csv", tadawul_top)

# ISRAEL (TA-35 Top) - Suffix: .TA
ta35_top = [
    "NICE.TA", "TEVA.TA", "LUMI.TA", "POLI.TA", "DSCT.TA", "ICL.TA", "MTRS.TA",
    "MZTF.TA", "BEZQ.TA", "ILCO.TA", "ESLT.TA", "NVMI.TA", "TSEM.TA", "ALHE.TA",
    "PHOE.TA", "HARE.TA", "KEN.TA", "ENOG.TA", "BIG.TA", "AZRG.TA"
]
save_csv("MIDEAST_ISRAEL_TA35.csv", ta35_top)


# ==========================================
# 4. ASIA-PACIFIC
# ==========================================
print("\n--- 🌏 ASIA-PACIFIC ---")

# JAPAN (Nikkei 225 Major) - Suffix: .T
nikkei_top = [
    "7203.T", "6758.T", "6861.T", "8035.T", "9984.T", "9983.T", "8306.T", "9432.T",
    "6098.T", "4063.T", "6501.T", "7974.T", "8001.T", "8031.T", "6367.T", "4502.T",
    "6902.T", "6954.T", "7267.T", "7741.T", "3382.T", "8766.T", "6273.T", "6981.T",
    "8058.T", "4503.T", "5108.T", "4452.T", "4901.T", "6146.T"
]
save_csv("ASIA_JAPAN_NIKKEI.csv", nikkei_top)

# AUSTRALIA (ASX 200 Major) - Suffix: .AX
asx_top = [
    "BHP.AX", "CBA.AX", "CSL.AX", "NAB.AX", "WBC.AX", "ANZ.AX", "WDS.AX", "MQG.AX",
    "WES.AX", "TLS.AX", "RIO.AX", "WOW.AX", "FMG.AX", "GMG.AX", "STO.AX", "ALL.AX",
    "COH.AX", "NCM.AX", "S32.AX", "BXB.AX", "SCG.AX", "QBE.AX", "SUN.AX", "RMD.AX",
    "TCL.AX", "ORG.AX", "SHL.AX", "BSL.AX", "ASX.AX", "NST.AX"
]
save_csv("ASIA_AUSTRALIA_ASX.csv", asx_top)

# INDIA (Nifty 50 Top) - Suffix: .NS
nifty_top = [
    "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "INFY.NS", "ICICIBANK.NS", "HUL.NS",
    "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "KOTAKBANK.NS", "LICI.NS", "LT.NS",
    "AXISBANK.NS", "ASIANPAINT.NS", "HCLTECH.NS", "MARUTI.NS", "TITAN.NS",
    "BAJFINANCE.NS", "SUNPHARMA.NS", "TATASTEEL.NS", "NTPC.NS", "ULTRACEMCO.NS"
]
save_csv("ASIA_INDIA_NIFTY.csv", nifty_top)

# HONG KONG (Hang Seng Top) - Suffix: .HK
hsi_top = [
    "0700.HK", "0939.HK", "1299.HK", "9988.HK", "0941.HK", "1398.HK", "3988.HK",
    "3690.HK", "0005.HK", "0883.HK", "2318.HK", "0388.HK", "1211.HK", "0857.HK",
    "9618.HK", "2020.HK", "0016.HK", "1109.HK", "0669.HK", "0001.HK", "0027.HK",
    "1093.HK", "0003.HK", "2269.HK", "0267.HK"
]
save_csv("ASIA_HK_HANGSENG.csv", hsi_top)

print("\n✨ Done! Universe files generated.")