"""
DCF Engine Cash Flow Calculations

FCFF and FCFE construction with both derivation methods.
"""
from __future__ import annotations

from dcf_engine.models import DCFInputs


def compute_interest_expense(
    inputs: DCFInputs
) -> dict[int, float]:
    """
    Compute interest expense for all forecast years.
    
    Uses end-of-period convention:
    InterestExpense_t = Debt_t * rd_t
    
    Returns:
        dict mapping year -> interest expense
    """
    forecast_years = inputs.timeline.forecast_years
    debt = inputs.debt
    
    interest = {}
    for year in forecast_years:
        debt_balance = debt.debt_balances[year]
        rd = debt.get_rd(year)
        interest[year] = debt_balance * rd
    
    return interest


def compute_net_borrowing(
    inputs: DCFInputs
) -> dict[int, float]:
    """
    Compute net borrowing for all forecast years.
    
    NetBorrowing_t = Debt_t - Debt_(t-1)
    Positive = new debt (cash inflow)
    
    Returns:
        dict mapping year -> net borrowing
    """
    base_year = inputs.timeline.base_year
    forecast_years = inputs.timeline.forecast_years
    all_years = [base_year] + forecast_years
    debt = inputs.debt
    
    net_borrowing = {}
    for i, year in enumerate(forecast_years):
        prev_year = all_years[i]
        net_borrowing[year] = debt.debt_balances[year] - debt.debt_balances[prev_year]
    
    return net_borrowing


def compute_fcff(
    nopat: dict[int, float],
    da: dict[int, float],
    delta_nwc: dict[int, float],
    capex: dict[int, float]
) -> dict[int, float]:
    """
    Compute Free Cash Flow to Firm for all forecast years.
    
    FCFF = NOPAT + D&A - ΔNWC - Capex
    
    Note on signs:
    - NOPAT: positive
    - D&A: add back (non-cash expense)
    - ΔNWC: subtract if positive (cash consumed), add if negative (cash released)
    - Capex: subtract (cash outflow, provided as positive number)
    
    Returns:
        dict mapping year -> FCFF
    """
    fcff = {}
    for year in nopat:
        fcff[year] = (
            nopat[year]
            + da.get(year, 0.0)
            - delta_nwc.get(year, 0.0)
            - capex.get(year, 0.0)
        )
    
    return fcff


def compute_fcfe_from_fcff(
    inputs: DCFInputs,
    fcff: dict[int, float],
    interest_expense: dict[int, float],
    net_borrowing: dict[int, float]
) -> dict[int, float]:
    """
    Compute FCFE from FCFF (Derivation 1, recommended).
    
    FCFE = FCFF - Interest*(1-TaxRate) + NetBorrowing
    
    Returns:
        dict mapping year -> FCFE
    """
    forecast_years = inputs.timeline.forecast_years
    
    fcfe = {}
    for year in forecast_years:
        tax_rate = inputs.tax.get_rate(year)
        after_tax_interest = interest_expense.get(year, 0.0) * (1 - tax_rate)
        fcfe[year] = (
            fcff[year]
            - after_tax_interest
            + net_borrowing.get(year, 0.0)
        )
    
    return fcfe


def compute_fcfe_from_net_income(
    net_income: dict[int, float],
    da: dict[int, float],
    delta_nwc: dict[int, float],
    capex: dict[int, float],
    net_borrowing: dict[int, float]
) -> dict[int, float]:
    """
    Compute FCFE from Net Income (Derivation 2).
    
    FCFE = NetIncome + D&A - ΔNWC - Capex + NetBorrowing
    
    Returns:
        dict mapping year -> FCFE
    """
    fcfe = {}
    for year in net_income:
        fcfe[year] = (
            net_income[year]
            + da.get(year, 0.0)
            - delta_nwc.get(year, 0.0)
            - capex.get(year, 0.0)
            + net_borrowing.get(year, 0.0)
        )
    
    return fcfe


def reconcile_fcfe(
    fcfe_from_fcff: dict[int, float],
    fcfe_from_net_income: dict[int, float],
    tolerance: float = 1e-6
) -> tuple[bool, dict[int, float]]:
    """
    Reconcile FCFE computed from both derivations.
    
    Returns:
        Tuple of (is_reconciled, differences_by_year)
    """
    differences = {}
    all_match = True
    
    for year in fcfe_from_fcff:
        diff = abs(fcfe_from_fcff[year] - fcfe_from_net_income.get(year, 0.0))
        differences[year] = diff
        if diff > tolerance:
            all_match = False
    
    return all_match, differences
