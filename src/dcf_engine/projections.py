"""
DCF Engine Operating Projections

Revenue, EBITDA, EBIT, NWC computations.
"""
from __future__ import annotations

from dcf_engine.models import DCFInputs


def compute_revenue(inputs: DCFInputs) -> dict[int, float]:
    """
    Compute revenue for all years (base + forecast).
    
    Uses explicit_revenue if provided, otherwise computes from
    base_revenue and growth_rates.
    
    Returns:
        dict mapping year -> revenue
    """
    base_year = inputs.timeline.base_year
    forecast_years = inputs.timeline.forecast_years
    rev_inputs = inputs.revenue
    
    revenue = {}
    
    if rev_inputs.explicit_revenue:
        # Use explicit series
        revenue.update(rev_inputs.explicit_revenue)
        # Ensure base year has value if not in explicit
        if base_year not in revenue and rev_inputs.base_revenue:
            revenue[base_year] = rev_inputs.base_revenue
    else:
        # Compute from base + growth rates
        revenue[base_year] = rev_inputs.base_revenue
        prev_revenue = rev_inputs.base_revenue
        
        for year in forecast_years:
            g = rev_inputs.growth_rates[year]
            new_revenue = prev_revenue * (1 + g)
            revenue[year] = new_revenue
            prev_revenue = new_revenue
    
    return revenue


def compute_operating_costs(
    inputs: DCFInputs,
    revenue: dict[int, float]
) -> dict[int, float]:
    """
    Compute operating costs for forecast years.
    
    OpCosts = Revenue * CostRatio
    
    Returns:
        dict mapping year -> operating costs
    """
    forecast_years = inputs.timeline.forecast_years
    op = inputs.operating
    
    costs = {}
    for year in forecast_years:
        if op.cost_ratios and year in op.cost_ratios:
            costs[year] = revenue[year] * op.cost_ratios[year]
        else:
            # If EBITDA is explicit, compute costs as Revenue - EBITDA
            if op.explicit_ebitda and year in op.explicit_ebitda:
                costs[year] = revenue[year] - op.explicit_ebitda[year]
            else:
                costs[year] = 0.0  # Default
    
    return costs


def compute_ebitda(
    inputs: DCFInputs,
    revenue: dict[int, float],
    operating_costs: dict[int, float]
) -> dict[int, float]:
    """
    Compute EBITDA for forecast years.
    
    EBITDA = Revenue - OpCosts
    Or use explicit_ebitda if provided.
    
    Returns:
        dict mapping year -> EBITDA
    """
    forecast_years = inputs.timeline.forecast_years
    op = inputs.operating
    
    ebitda = {}
    for year in forecast_years:
        if op.explicit_ebitda and year in op.explicit_ebitda:
            ebitda[year] = op.explicit_ebitda[year]
        else:
            ebitda[year] = revenue[year] - operating_costs[year]
    
    return ebitda


def compute_ebit(
    inputs: DCFInputs,
    ebitda: dict[int, float]
) -> dict[int, float]:
    """
    Compute EBIT for forecast years.
    
    EBIT = EBITDA - D&A
    
    Returns:
        dict mapping year -> EBIT
    """
    forecast_years = inputs.timeline.forecast_years
    da = inputs.operating.depreciation_amortization
    
    ebit = {}
    for year in forecast_years:
        ebit[year] = ebitda[year] - da[year]
    
    return ebit


def compute_nwc(
    inputs: DCFInputs,
    revenue: dict[int, float]
) -> dict[int, float]:
    """
    Compute NWC for base and forecast years.
    
    NWC = Revenue * NWC_percent (if driver)
    Or use explicit_nwc if provided.
    
    Returns:
        dict mapping year -> NWC
    """
    base_year = inputs.timeline.base_year
    forecast_years = inputs.timeline.forecast_years
    all_years = [base_year] + forecast_years
    nwc_inputs = inputs.nwc
    
    nwc = {}
    for year in all_years:
        if nwc_inputs.explicit_nwc and year in nwc_inputs.explicit_nwc:
            nwc[year] = nwc_inputs.explicit_nwc[year]
        elif nwc_inputs.nwc_percent and year in nwc_inputs.nwc_percent:
            nwc[year] = revenue[year] * nwc_inputs.nwc_percent[year]
        elif nwc_inputs.base_nwc is not None and year == base_year:
            nwc[year] = nwc_inputs.base_nwc
    
    return nwc


def compute_delta_nwc(
    inputs: DCFInputs,
    nwc: dict[int, float]
) -> dict[int, float]:
    """
    Compute change in NWC for forecast years.
    
    ΔNWC = NWC_t - NWC_(t-1)
    
    Cash-flow sign rule:
    - ΔNWC > 0 means cash consumed (subtract in cash flow)
    - ΔNWC < 0 means cash released (add in cash flow)
    
    Returns:
        dict mapping year -> ΔNWC
    """
    base_year = inputs.timeline.base_year
    forecast_years = inputs.timeline.forecast_years
    all_years = [base_year] + forecast_years
    
    delta_nwc = {}
    for i, year in enumerate(forecast_years):
        prev_year = all_years[i]  # base_year or previous forecast year
        delta_nwc[year] = nwc[year] - nwc[prev_year]
    
    return delta_nwc
