"""
DCF Engine Discounting Logic

Discount factor calculations and present value computations.
Supports both constant and year-specific-flat discounting modes.
"""
from __future__ import annotations

from dcf_engine.models import DiscountingMode


def compute_discount_factors(
    rates: dict[int, float] | float,
    years: list[int],
    base_year: int,
    mode: DiscountingMode
) -> dict[int, float]:
    """
    Compute discount factors for each forecast year.
    
    Constant mode (rate is a single float):
        DF_i = 1 / (1 + r)^i
        where i is the period number (1, 2, 3, ...)
    
    Year-specific-flat mode (rate is a dict by year):
        DF_i = 1 / (1 + r_i)^i
        where r_i is the rate for year i, applied for i periods
    
    Returns:
        dict mapping year -> discount factor
    """
    discount_factors = {}
    
    for year in years:
        period = year - base_year  # 1, 2, 3, ...
        
        if mode == DiscountingMode.CONSTANT:
            if isinstance(rates, dict):
                # Use first year's rate as constant
                r = list(rates.values())[0]
            else:
                r = rates
            discount_factors[year] = 1.0 / ((1 + r) ** period)
        
        elif mode == DiscountingMode.YEAR_SPECIFIC_FLAT:
            if isinstance(rates, dict):
                r = rates[year]
            else:
                r = rates
            # Use year i's rate applied for i periods
            discount_factors[year] = 1.0 / ((1 + r) ** period)
    
    return discount_factors


def compute_pv_series(
    cash_flows: dict[int, float],
    discount_factors: dict[int, float]
) -> dict[int, float]:
    """
    Compute present value of each cash flow in a series.
    
    PV(CF_i) = CF_i * DF_i
    
    Returns:
        dict mapping year -> PV of that year's cash flow
    """
    pv = {}
    for year, cf in cash_flows.items():
        df = discount_factors.get(year, 1.0)
        pv[year] = cf * df
    
    return pv


def compute_pv_single(
    value: float,
    rate: float,
    periods: int
) -> float:
    """
    Compute present value of a single future value.
    
    PV = FV / (1 + r)^n
    
    Args:
        value: Future value
        rate: Discount rate
        periods: Number of periods
    
    Returns:
        Present value
    """
    return value / ((1 + rate) ** periods)


def sum_pv(pv_series: dict[int, float]) -> float:
    """
    Sum all present values in a series.
    
    Returns:
        Sum of PVs
    """
    return sum(pv_series.values())
