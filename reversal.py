from typing import Dict
import pandas as pd
import yfinance as yf

from metrics import get_row

def trend_reversal_scores(tkr: yf.Ticker) -> Dict[str, str]:
    out: Dict[str, str] = {}

    q_income = tkr.quarterly_income_stmt
    q_cf = tkr.quarterly_cashflow
    q_bs = tkr.quarterly_balance_sheet

    def stabilizing(series: pd.Series) -> str:
        s = pd.to_numeric(series, errors="coerce").dropna()
        if len(s) < 3:
            return "🔴"
        vals = s.iloc[:4].astype(float).values
        if len(vals) < 3:
            return "🔴"
        d1 = vals[0] - vals[1]
        d2 = vals[1] - vals[2]
        if d1 >= 0 and d2 >= 0:
            return "🟢"
        if d1 >= 0 or d2 >= 0:
            return "🟡"
        return "🔴"

    rev = get_row(q_income, ["Total Revenue", "TotalRevenue"])
    op_inc = get_row(q_income, ["Operating Income", "OperatingIncome"])
    if rev is not None and op_inc is not None:
        op_margin = (pd.to_numeric(op_inc, errors="coerce") / pd.to_numeric(rev, errors="coerce")) * 100
        out["1. Margin Stabilization"] = stabilizing(op_margin)
    else:
        out["1. Margin Stabilization"] = "🔴"

    ocf = get_row(q_cf, ["Operating Cash Flow", "Total Cash From Operating Activities"])
    capex = get_row(q_cf, ["Capital Expenditure", "CapitalExpenditure"])
    if ocf is not None and capex is not None:
        fcf_v = pd.to_numeric(ocf, errors="coerce") + pd.to_numeric(capex, errors="coerce")
        if len(fcf_v.dropna()) >= 3 and float(fcf_v.iloc[0]) >= 0:
            out["2. Cash-Flow Reversal"] = stabilizing(fcf_v)
        else:
            out["2. Cash-Flow Reversal"] = "🔴"
    else:
        out["2. Cash-Flow Reversal"] = "🔴"

    total_debt = get_row(q_bs, ["Total Debt", "TotalDebt", "Long Term Debt", "LongTermDebt"])
    cash = get_row(q_bs, ["Cash And Cash Equivalents", "CashAndCashEquivalents"])
    if total_debt is not None and cash is not None:
        nd = pd.to_numeric(total_debt, errors="coerce") - pd.to_numeric(cash, errors="coerce")
        out["3. Balance-Sheet Healing"] = stabilizing(-nd)
    else:
        out["3. Balance-Sheet Healing"] = "🔴"

    out["4. ROIC Inflection"] = stabilizing(op_inc) if op_inc is not None else "🔴"
    out["5. Capex Discipline"] = stabilizing(capex) if capex is not None else "🔴"

    gross_profit = get_row(q_income, ["Gross Profit", "GrossProfit"])
    if rev is not None and gross_profit is not None:
        gp_margin = (pd.to_numeric(gross_profit, errors="coerce") / pd.to_numeric(rev, errors="coerce")) * 100
        out["6. Revenue Quality"] = stabilizing(gp_margin)
    else:
        out["6. Revenue Quality"] = "🔴"

    try:
        hist = tkr.history(period="2y", interval="1d")
        if not hist.empty:
            close = hist["Close"]
            ma200 = close.rolling(200).mean().iloc[-1]
            price = float(close.iloc[-1])
            fund = out.get("1. Margin Stabilization", "🔴")
            if fund in ("🟢", "🟡") and price <= ma200:
                out["7. Fundamentals vs Valuation Divergence"] = "🟢"
            elif fund in ("🟢", "🟡"):
                out["7. Fundamentals vs Valuation Divergence"] = "🟡"
            else:
                out["7. Fundamentals vs Valuation Divergence"] = "🔴"
        else:
            out["7. Fundamentals vs Valuation Divergence"] = "🔴"
    except Exception:
        out["7. Fundamentals vs Valuation Divergence"] = "🔴"

    return out
