"""
DCF Engine Discount Rate Calculations

CAPM-based Ke, WACC with target/book-value weights.
"""
from __future__ import annotations

from dcf_engine.models import DCFInputs, WeightingMode


def compute_ke(inputs: DCFInputs) -> float:
    """
    Compute cost of equity using CAPM.
    
    Ke = rf + beta * (rm - rf)
    
    Returns:
        float: cost of equity
    """
    capm = inputs.capm
    
    # Check for direct override
    if capm.ke_override is not None:
        return capm.ke_override
    
    # CAPM calculation
    ke = capm.risk_free_rate + capm.beta * (capm.market_return - capm.risk_free_rate)
    return ke


def compute_equity_book_rollforward(
    inputs: DCFInputs,
    net_income: dict[int, float]
) -> dict[int, float]:
    """
    Roll forward equity book value for all years.
    
    EquityBook_t = EquityBook_(t-1) + NetIncome_t - Dividends_t + NewEquity_t
    
    Returns:
        dict mapping year -> equity book value
    """
    base_year = inputs.timeline.base_year
    forecast_years = inputs.timeline.forecast_years
    equity_inputs = inputs.wacc.equity_book_inputs
    
    if equity_inputs is None:
        return {}
    
    equity_book = {base_year: equity_inputs.base_equity_book}
    
    all_years = [base_year] + forecast_years
    for i, year in enumerate(forecast_years):
        prev_year = all_years[i]
        
        dividends = 0.0
        if equity_inputs.dividends and year in equity_inputs.dividends:
            dividends = equity_inputs.dividends[year]
        
        new_equity = 0.0
        if equity_inputs.new_equity and year in equity_inputs.new_equity:
            new_equity = equity_inputs.new_equity[year]
        
        equity_book[year] = (
            equity_book[prev_year]
            + net_income.get(year, 0.0)
            - dividends
            + new_equity
        )
    
    return equity_book


def compute_wacc(
    inputs: DCFInputs,
    ke: float,
    equity_book: dict[int, float] | None = None
) -> dict[int, float]:
    """
    Compute WACC for all forecast years.
    
    WACC = wE * Ke + wD * rd * (1 - TaxRate)
    
    Weights can be:
    - Target: fixed wE, wD from inputs
    - Book-value: computed from debt and equity book values
    
    Returns:
        dict mapping year -> WACC
    """
    forecast_years = inputs.timeline.forecast_years
    wacc_inputs = inputs.wacc
    debt = inputs.debt
    
    wacc = {}
    
    for year in forecast_years:
        tax_rate = inputs.tax.get_rate(year)
        rd = debt.get_rd(year)
        debt_balance = debt.debt_balances[year]
        
        if wacc_inputs.weighting_mode == WeightingMode.TARGET:
            w_e = wacc_inputs.target_weight_equity
            w_d = wacc_inputs.target_weight_debt
        else:
            # Book-value weights
            equity_value = equity_book.get(year, 0.0) if equity_book else 0.0
            total_capital = debt_balance + equity_value
            if total_capital > 0:
                w_d = debt_balance / total_capital
                w_e = equity_value / total_capital
            else:
                w_d = 0.0
                w_e = 1.0
        
        wacc[year] = w_e * ke + w_d * rd * (1 - tax_rate)
    
    return wacc


def get_wacc_weights(
    inputs: DCFInputs,
    equity_book: dict[int, float] | None = None
) -> dict[int, tuple[float, float]]:
    """
    Get WACC weights (wE, wD) for all forecast years.
    
    Returns:
        dict mapping year -> (weight_equity, weight_debt)
    """
    forecast_years = inputs.timeline.forecast_years
    wacc_inputs = inputs.wacc
    debt = inputs.debt
    
    weights = {}
    
    for year in forecast_years:
        debt_balance = debt.debt_balances[year]
        
        if wacc_inputs.weighting_mode == WeightingMode.TARGET:
            w_e = wacc_inputs.target_weight_equity
            w_d = wacc_inputs.target_weight_debt
        else:
            # Book-value weights
            equity_value = equity_book.get(year, 0.0) if equity_book else 0.0
            total_capital = debt_balance + equity_value
            if total_capital > 0:
                w_d = debt_balance / total_capital
                w_e = equity_value / total_capital
            else:
                w_d = 0.0
                w_e = 1.0
        
        weights[year] = (w_e, w_d)
    
    return weights
