"""
DCF Engine Valuation

Enterprise Value and Equity Value bridges with reconciliation.
"""
from __future__ import annotations


def compute_enterprise_value(
    sum_pv_fcff: float,
    pv_terminal_value: float
) -> float:
    """
    Compute Enterprise Value from FCFF approach.
    
    EV = Sum(PV of FCFFs) + PV(Terminal Value)
    
    Returns:
        Enterprise Value
    """
    return sum_pv_fcff + pv_terminal_value


def compute_net_debt(
    debt_balance: float,
    cash_and_equivalents: float
) -> float:
    """
    Compute Net Debt at valuation date.
    
    NetDebt = Debt - Cash
    
    Returns:
        Net Debt (positive means net debt, negative means net cash)
    """
    return debt_balance - cash_and_equivalents


def compute_equity_from_ev(
    enterprise_value: float,
    net_debt: float
) -> float:
    """
    Compute Equity Value from Enterprise Value (FCFF approach).
    
    EquityValue = EV - NetDebt
    
    Returns:
        Equity Value
    """
    return enterprise_value - net_debt


def compute_equity_direct(
    sum_pv_fcfe: float,
    pv_terminal_value_fcfe: float
) -> float:
    """
    Compute Equity Value directly from FCFE approach.
    
    EquityValue = Sum(PV of FCFEs) + PV(Terminal Value)
    
    Returns:
        Equity Value
    """
    return sum_pv_fcfe + pv_terminal_value_fcfe


def reconcile_valuations(
    equity_from_ev: float,
    equity_direct: float,
    tolerance: float = 0.01  # 1% tolerance
) -> tuple[float, list[str]]:
    """
    Reconcile equity values from FCFF/WACC and FCFE/Ke approaches.
    
    Returns:
        Tuple of (difference, list of reconciliation notes)
    """
    difference = equity_from_ev - equity_direct
    pct_diff = abs(difference / equity_from_ev) * 100 if equity_from_ev != 0 else 0
    
    notes = []
    
    if abs(pct_diff) < tolerance * 100:
        notes.append(f"Values reconcile within {tolerance*100:.1f}% tolerance.")
    else:
        notes.append(f"Difference of {difference:,.2f} ({pct_diff:.2f}%) between methods.")
        notes.append("Potential sources of difference:")
        notes.append("  - Interest tax shield treatment differences")
        notes.append("  - Debt/equity weighting assumptions")
        notes.append("  - Terminal value discount rate differences (WACC vs Ke)")
        notes.append("  - Net borrowing timing assumptions")
    
    return difference, notes
