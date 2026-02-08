"""
DCF Engine Input Validation

Fail-fast validation for DCF inputs with precise error messages.
"""
from __future__ import annotations

from dcf_engine.models import (
    DCFInputs,
    TerminalValueMethod,
    WeightingMode,
)


class ValidationError(Exception):
    """Raised when input validation fails."""
    pass


def validate_inputs(inputs: DCFInputs) -> None:
    """
    Validate DCF inputs for completeness and consistency.
    
    Raises ValidationError with precise message on failure.
    """
    _validate_timeline_coverage(inputs)
    _validate_revenue_coverage(inputs)
    _validate_operating_coverage(inputs)
    _validate_nwc_coverage(inputs)
    _validate_investment_coverage(inputs)
    _validate_debt_coverage(inputs)
    _validate_terminal_value(inputs)
    _validate_wacc_inputs(inputs)


def _validate_timeline_coverage(inputs: DCFInputs) -> None:
    """Ensure timeline is well-formed."""
    base = inputs.timeline.base_year
    forecast = inputs.timeline.forecast_years
    
    # Forecast years should follow base year
    for year in forecast:
        if year <= base:
            raise ValidationError(
                f"Forecast year {year} must be after base year {base}"
            )


def _validate_revenue_coverage(inputs: DCFInputs) -> None:
    """Ensure revenue can be computed for all forecast years."""
    forecast = inputs.timeline.forecast_years
    rev = inputs.revenue
    
    if rev.explicit_revenue:
        missing = [y for y in forecast if y not in rev.explicit_revenue]
        if missing:
            raise ValidationError(
                f"Missing explicit revenue for years: {missing}"
            )
    elif rev.growth_rates:
        missing = [y for y in forecast if y not in rev.growth_rates]
        if missing:
            raise ValidationError(
                f"Missing growth rates for years: {missing}"
            )
        if rev.base_revenue is None:
            raise ValidationError(
                "base_revenue required when using growth_rates"
            )


def _validate_operating_coverage(inputs: DCFInputs) -> None:
    """Ensure EBITDA can be computed for all forecast years."""
    forecast = inputs.timeline.forecast_years
    op = inputs.operating
    
    # D&A is always required
    missing_da = [y for y in forecast if y not in op.depreciation_amortization]
    if missing_da:
        raise ValidationError(
            f"Missing depreciation_amortization for years: {missing_da}"
        )
    
    # Either explicit EBITDA or cost ratios
    if op.explicit_ebitda:
        missing = [y for y in forecast if y not in op.explicit_ebitda]
        if missing:
            raise ValidationError(
                f"Missing explicit_ebitda for years: {missing}"
            )
    elif op.cost_ratios:
        missing = [y for y in forecast if y not in op.cost_ratios]
        if missing:
            raise ValidationError(
                f"Missing cost_ratios for years: {missing}"
            )
    else:
        raise ValidationError(
            "Operating inputs underdetermined: provide either 'explicit_ebitda' or 'cost_ratios'"
        )


def _validate_nwc_coverage(inputs: DCFInputs) -> None:
    """Ensure NWC can be computed for base and all forecast years."""
    base = inputs.timeline.base_year
    forecast = inputs.timeline.forecast_years
    all_years = [base] + forecast
    nwc = inputs.nwc
    
    if nwc.explicit_nwc:
        missing = [y for y in all_years if y not in nwc.explicit_nwc]
        if missing:
            raise ValidationError(
                f"Missing explicit_nwc for years: {missing}"
            )
    elif nwc.nwc_percent:
        missing = [y for y in all_years if y not in nwc.nwc_percent]
        if missing:
            raise ValidationError(
                f"Missing nwc_percent for years: {missing}"
            )
    else:
        raise ValidationError(
            "NWC inputs underdetermined: provide either 'explicit_nwc' or 'nwc_percent'"
        )


def _validate_investment_coverage(inputs: DCFInputs) -> None:
    """Ensure Capex is provided for all forecast years."""
    forecast = inputs.timeline.forecast_years
    missing = [y for y in forecast if y not in inputs.investments.capex]
    if missing:
        raise ValidationError(
            f"Missing capex for years: {missing}"
        )


def _validate_debt_coverage(inputs: DCFInputs) -> None:
    """Ensure debt data covers base and all forecast years."""
    base = inputs.timeline.base_year
    forecast = inputs.timeline.forecast_years
    all_years = [base] + forecast
    debt = inputs.debt
    
    missing = [y for y in all_years if y not in debt.debt_balances]
    if missing:
        raise ValidationError(
            f"Missing debt_balances for years: {missing}"
        )
    
    # Cost of debt must cover forecast years
    if isinstance(debt.cost_of_debt, dict):
        missing = [y for y in all_years if y not in debt.cost_of_debt]
        if missing:
            raise ValidationError(
                f"Missing cost_of_debt for years: {missing}"
            )


def _validate_terminal_value(inputs: DCFInputs) -> None:
    """Validate terminal value configuration."""
    tv = inputs.terminal_value
    
    if tv.method == TerminalValueMethod.PERPETUITY:
        g = tv.perpetuity_growth_rate
        if g is None:
            raise ValidationError(
                "Perpetuity method requires perpetuity_growth_rate (g)"
            )
        # We'll validate g < discount_rate during computation when WACC is known


def _validate_wacc_inputs(inputs: DCFInputs) -> None:
    """Validate WACC weight configuration."""
    wacc = inputs.wacc
    
    if wacc.weighting_mode == WeightingMode.BOOK_VALUE:
        if wacc.equity_book_inputs is None:
            raise ValidationError(
                "Book-value weighting mode requires equity_book_inputs"
            )
