"""
DCF Engine Terminal Value Calculations

Perpetuity (Gordon) and Exit Multiple methods.
"""
from __future__ import annotations

from dcf_engine.models import (
    TerminalValueMethod,
    ExitMultipleMetric,
    DCFInputs,
)


class TerminalValueError(Exception):
    """Raised when terminal value calculation fails."""
    pass


def compute_terminal_value_perpetuity(
    final_cash_flow: float,
    discount_rate: float,
    growth_rate: float
) -> float:
    """
    Compute terminal value using perpetuity (Gordon) growth model.
    
    TV = CF_N * (1 + g) / (r - g)
    
    If g = 0:
        TV = CF_N / r
    
    Args:
        final_cash_flow: Cash flow in final forecast year (FCFF or FCFE)
        discount_rate: WACC or Ke
        growth_rate: Perpetuity growth rate (g)
    
    Returns:
        Terminal value
    
    Raises:
        TerminalValueError: If g >= discount_rate
    """
    if growth_rate >= discount_rate:
        raise TerminalValueError(
            f"Growth rate ({growth_rate:.4f}) must be less than discount rate ({discount_rate:.4f})"
        )
    
    if abs(growth_rate) < 1e-10:
        # Zero growth case
        return final_cash_flow / discount_rate
    else:
        return final_cash_flow * (1 + growth_rate) / (discount_rate - growth_rate)


def compute_terminal_value_exit_multiple(
    metric_value: float,
    multiple: float
) -> float:
    """
    Compute terminal value using exit multiple.
    
    TV = Metric_N * Multiple
    
    Args:
        metric_value: Value of the exit metric (EBITDA, EBIT, or Revenue)
        multiple: Exit multiple
    
    Returns:
        Terminal value
    """
    return metric_value * multiple


def get_exit_metric_value(
    inputs: DCFInputs,
    ebitda: dict[int, float],
    ebit: dict[int, float],
    revenue: dict[int, float]
) -> float:
    """
    Get the value of the exit metric for the final year.
    
    Returns:
        Metric value for terminal value calculation
    """
    final_year = inputs.timeline.forecast_years[-1]
    metric = inputs.terminal_value.exit_metric
    
    if metric == ExitMultipleMetric.EBITDA:
        return ebitda[final_year]
    elif metric == ExitMultipleMetric.EBIT:
        return ebit[final_year]
    elif metric == ExitMultipleMetric.REVENUE:
        return revenue[final_year]
    else:
        raise TerminalValueError(f"Unknown exit metric: {metric}")


def compute_terminal_value(
    inputs: DCFInputs,
    fcff: dict[int, float],
    fcfe: dict[int, float],
    wacc: dict[int, float],
    ke: float,
    ebitda: dict[int, float] | None = None,
    ebit: dict[int, float] | None = None,
    revenue: dict[int, float] | None = None
) -> tuple[float, float]:
    """
    Compute terminal values for both FCFF and FCFE valuation.
    
    Returns:
        Tuple of (TV for FCFF valuation, TV for FCFE valuation)
    """
    final_year = inputs.timeline.forecast_years[-1]
    tv_inputs = inputs.terminal_value
    final_wacc = wacc[final_year]
    
    if tv_inputs.method == TerminalValueMethod.PERPETUITY:
        g = tv_inputs.perpetuity_growth_rate
        
        tv_fcff = compute_terminal_value_perpetuity(
            fcff[final_year], final_wacc, g
        )
        tv_fcfe = compute_terminal_value_perpetuity(
            fcfe[final_year], ke, g
        )
    
    elif tv_inputs.method == TerminalValueMethod.EXIT_MULTIPLE:
        if ebitda is None or ebit is None or revenue is None:
            raise TerminalValueError(
                "Exit multiple method requires ebitda, ebit, and revenue"
            )
        
        metric_value = get_exit_metric_value(inputs, ebitda, ebit, revenue)
        multiple = tv_inputs.exit_multiple
        
        # For exit multiple, we typically use the same TV for both methods
        # (it's based on operating metrics, not cash flows)
        tv_fcff = compute_terminal_value_exit_multiple(metric_value, multiple)
        tv_fcfe = tv_fcff  # Same EV-based terminal value
    
    else:
        raise TerminalValueError(f"Unknown terminal value method: {tv_inputs.method}")
    
    return tv_fcff, tv_fcfe
