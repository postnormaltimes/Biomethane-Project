"""
DCF Engine Tax Calculations

Mode A: Tax on EBIT -> NOPAT (for FCFF)
Mode B: Profit tax -> Net Income (for FCFE)
"""
from __future__ import annotations

from dcf_engine.models import DCFInputs


def compute_tax_on_ebit(
    inputs: DCFInputs,
    ebit: dict[int, float]
) -> dict[int, float]:
    """
    Compute tax on EBIT for all forecast years (Mode A).
    
    TaxOnEBIT = EBIT * TaxRate
    
    Returns:
        dict mapping year -> tax on EBIT
    """
    forecast_years = inputs.timeline.forecast_years
    
    tax_on_ebit = {}
    for year in forecast_years:
        tax_rate = inputs.tax.get_rate(year)
        tax_on_ebit[year] = ebit[year] * tax_rate
    
    return tax_on_ebit


def compute_nopat(
    inputs: DCFInputs,
    ebit: dict[int, float]
) -> dict[int, float]:
    """
    Compute NOPAT for all forecast years (Mode A).
    
    NOPAT = EBIT * (1 - TaxRate)
    
    Returns:
        dict mapping year -> NOPAT
    """
    forecast_years = inputs.timeline.forecast_years
    
    nopat = {}
    for year in forecast_years:
        tax_rate = inputs.tax.get_rate(year)
        nopat[year] = ebit[year] * (1 - tax_rate)
    
    return nopat


def compute_ebt(
    ebit: dict[int, float],
    interest_expense: dict[int, float]
) -> dict[int, float]:
    """
    Compute Earnings Before Tax for all forecast years (Mode B).
    
    EBT = EBIT - InterestExpense
    
    Returns:
        dict mapping year -> EBT
    """
    ebt = {}
    for year in ebit:
        ebt[year] = ebit[year] - interest_expense.get(year, 0.0)
    
    return ebt


def compute_taxes_on_ebt(
    inputs: DCFInputs,
    ebt: dict[int, float]
) -> dict[int, float]:
    """
    Compute taxes on EBT for all forecast years (Mode B).
    
    Taxes = max(0, EBT) * TaxRate
    
    Returns:
        dict mapping year -> taxes
    """
    forecast_years = inputs.timeline.forecast_years
    
    taxes = {}
    for year in forecast_years:
        tax_rate = inputs.tax.get_rate(year)
        # Only positive EBT is taxed (loss carry-forward not modeled)
        taxes[year] = max(0.0, ebt[year]) * tax_rate
    
    return taxes


def compute_net_income(
    ebt: dict[int, float],
    taxes_on_ebt: dict[int, float]
) -> dict[int, float]:
    """
    Compute Net Income for all forecast years (Mode B).
    
    NetIncome = EBT - Taxes
    
    Returns:
        dict mapping year -> net income
    """
    net_income = {}
    for year in ebt:
        net_income[year] = ebt[year] - taxes_on_ebt.get(year, 0.0)
    
    return net_income
