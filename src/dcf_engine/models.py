"""
DCF Engine Core Data Models

Pydantic models for all DCF inputs and outputs with validation.
"""
from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field, field_validator, model_validator


class DiscountingMode(str, Enum):
    """Discounting convention for present value calculations."""
    CONSTANT = "constant"
    YEAR_SPECIFIC_FLAT = "year_specific_flat"


class TerminalValueMethod(str, Enum):
    """Method for computing terminal value."""
    PERPETUITY = "perpetuity"
    EXIT_MULTIPLE = "exit_multiple"


class WeightingMode(str, Enum):
    """Method for computing WACC capital structure weights."""
    TARGET = "target"
    BOOK_VALUE = "book_value"


class ExitMultipleMetric(str, Enum):
    """Metric to use for exit multiple terminal value."""
    EBITDA = "ebitda"
    EBIT = "ebit"
    REVENUE = "revenue"


# ============================================================================
# INPUT MODELS
# ============================================================================

class TimelineInputs(BaseModel):
    """Timeline configuration."""
    base_year: int = Field(..., description="Base year (t0), valuation date is end of this year")
    forecast_years: list[int] = Field(..., min_length=1, description="List of forecast years [t1, ..., tN]")

    @field_validator("forecast_years")
    @classmethod
    def validate_forecast_years(cls, v: list[int], info) -> list[int]:
        if len(v) != len(set(v)):
            raise ValueError("Forecast years must be unique")
        return sorted(v)


class RevenueInputs(BaseModel):
    """Revenue projection inputs - either explicit series or growth drivers."""
    base_revenue: Optional[float] = Field(None, description="Revenue at base year (t0)")
    explicit_revenue: Optional[dict[int, float]] = Field(None, description="Explicit revenue by year")
    growth_rates: Optional[dict[int, float]] = Field(None, description="Growth rates by forecast year")

    @model_validator(mode="after")
    def validate_revenue_inputs(self) -> "RevenueInputs":
        has_explicit = self.explicit_revenue is not None
        has_drivers = self.base_revenue is not None and self.growth_rates is not None
        if not has_explicit and not has_drivers:
            raise ValueError(
                "Revenue inputs underdetermined: provide either 'explicit_revenue' "
                "or both 'base_revenue' and 'growth_rates'"
            )
        return self


class OperatingInputs(BaseModel):
    """Operating cost and EBITDA inputs."""
    cost_ratios: Optional[dict[int, float]] = Field(None, description="Operating cost as % of revenue by year")
    explicit_ebitda: Optional[dict[int, float]] = Field(None, description="Explicit EBITDA by year (overrides cost ratios)")
    depreciation_amortization: dict[int, float] = Field(..., description="D&A by year")


class NWCInputs(BaseModel):
    """Net Working Capital inputs."""
    base_nwc: Optional[float] = Field(None, description="NWC at base year (or compute from nwc_percent)")
    nwc_percent: Optional[dict[int, float]] = Field(None, description="NWC as % of revenue by year (including base year)")
    explicit_nwc: Optional[dict[int, float]] = Field(None, description="Explicit NWC by year (overrides nwc_percent)")


class InvestmentInputs(BaseModel):
    """Capital expenditure inputs."""
    capex: dict[int, float] = Field(..., description="Capex by forecast year (positive = cash outflow)")


class TaxInputs(BaseModel):
    """Tax rate inputs."""
    tax_rate: float | dict[int, float] = Field(..., description="Tax rate (constant or by year)")

    def get_rate(self, year: int) -> float:
        """Get tax rate for a specific year."""
        if isinstance(self.tax_rate, dict):
            return self.tax_rate[year]
        return self.tax_rate


class CAPMInputs(BaseModel):
    """CAPM inputs for cost of equity."""
    risk_free_rate: float = Field(..., alias="rf", description="Risk-free rate")
    market_return: float = Field(..., alias="rm", description="Expected market return")
    beta: float = Field(..., description="Equity beta")
    ke_override: Optional[float] = Field(None, description="Direct Ke override (optional)")

    model_config = {"populate_by_name": True}


class DebtInputs(BaseModel):
    """Debt and cost of debt inputs."""
    debt_balances: dict[int, float] = Field(..., description="Debt balance by year (including base year)")
    cost_of_debt: float | dict[int, float] = Field(..., alias="rd", description="Cost of debt (constant or by year)")

    model_config = {"populate_by_name": True}

    def get_rd(self, year: int) -> float:
        """Get cost of debt for a specific year."""
        if isinstance(self.cost_of_debt, dict):
            return self.cost_of_debt[year]
        return self.cost_of_debt


class EquityBookInputs(BaseModel):
    """Equity book value inputs for book-value weighted WACC."""
    base_equity_book: float = Field(..., description="Equity book value at base year")
    dividends: Optional[dict[int, float]] = Field(None, description="Dividends by forecast year")
    new_equity: Optional[dict[int, float]] = Field(None, description="New equity issuance by forecast year")


class WACCInputs(BaseModel):
    """WACC configuration inputs."""
    weighting_mode: WeightingMode = Field(WeightingMode.TARGET, description="How to compute capital structure weights")
    target_weight_equity: Optional[float] = Field(None, alias="wE", description="Target equity weight (for target mode)")
    target_weight_debt: Optional[float] = Field(None, alias="wD", description="Target debt weight (for target mode)")
    equity_book_inputs: Optional[EquityBookInputs] = Field(None, description="Required for book_value weighting mode")

    model_config = {"populate_by_name": True}

    @model_validator(mode="after")
    def validate_weights(self) -> "WACCInputs":
        if self.weighting_mode == WeightingMode.TARGET:
            if self.target_weight_equity is None or self.target_weight_debt is None:
                raise ValueError("Target weighting mode requires 'wE' and 'wD'")
            if abs(self.target_weight_equity + self.target_weight_debt - 1.0) > 1e-6:
                raise ValueError(f"Target weights must sum to 1.0, got {self.target_weight_equity + self.target_weight_debt}")
        elif self.weighting_mode == WeightingMode.BOOK_VALUE:
            if self.equity_book_inputs is None:
                raise ValueError("Book-value weighting mode requires 'equity_book_inputs'")
        return self


class TerminalValueInputs(BaseModel):
    """Terminal value configuration."""
    method: TerminalValueMethod = Field(..., description="Terminal value calculation method")
    perpetuity_growth_rate: Optional[float] = Field(None, alias="g", description="Perpetuity growth rate (for perpetuity method)")
    exit_multiple: Optional[float] = Field(None, description="Exit multiple (for exit_multiple method)")
    exit_metric: Optional[ExitMultipleMetric] = Field(None, description="Metric for exit multiple")

    model_config = {"populate_by_name": True}

    @model_validator(mode="after")
    def validate_tv_inputs(self) -> "TerminalValueInputs":
        if self.method == TerminalValueMethod.PERPETUITY:
            if self.perpetuity_growth_rate is None:
                raise ValueError("Perpetuity method requires 'g' (perpetuity_growth_rate)")
        elif self.method == TerminalValueMethod.EXIT_MULTIPLE:
            if self.exit_multiple is None or self.exit_metric is None:
                raise ValueError("Exit multiple method requires 'exit_multiple' and 'exit_metric'")
        return self


class NetDebtInputs(BaseModel):
    """Net debt components for equity bridge."""
    cash_and_equivalents: float = Field(..., description="Cash and cash equivalents at base year")
    # Debt comes from DebtInputs


class DCFInputs(BaseModel):
    """Complete DCF model inputs."""
    timeline: TimelineInputs
    revenue: RevenueInputs
    operating: OperatingInputs
    nwc: NWCInputs
    investments: InvestmentInputs
    tax: TaxInputs
    capm: CAPMInputs
    debt: DebtInputs
    wacc: WACCInputs
    terminal_value: TerminalValueInputs
    net_debt: NetDebtInputs
    discounting_mode: DiscountingMode = Field(
        DiscountingMode.YEAR_SPECIFIC_FLAT,
        description="Discounting convention"
    )


# ============================================================================
# OUTPUT MODELS
# ============================================================================

class YearlyProjection(BaseModel):
    """Single year's projection data."""
    year: int
    revenue: float
    operating_costs: float
    ebitda: float
    depreciation_amortization: float
    ebit: float
    tax_on_ebit: float
    nopat: float
    # Mode B (for FCFE)
    interest_expense: float
    ebt: float
    taxes_on_ebt: float
    net_income: float


class YearlyNWC(BaseModel):
    """Single year's NWC data."""
    year: int
    nwc: float
    delta_nwc: float


class YearlyCashFlow(BaseModel):
    """Single year's cash flow data."""
    year: int
    nopat: float
    depreciation_amortization: float
    delta_nwc: float
    capex: float
    fcff: float
    interest_expense: float
    interest_tax_shield: float
    net_borrowing: float
    fcfe: float


class YearlyDiscount(BaseModel):
    """Single year's discounting data."""
    year: int
    period: int  # 1, 2, 3, ...
    wacc: float
    ke: float
    discount_factor_wacc: float
    discount_factor_ke: float
    fcff: float
    fcfe: float
    pv_fcff: float
    pv_fcfe: float


class TerminalValueOutput(BaseModel):
    """Terminal value computation details."""
    method: TerminalValueMethod
    final_year: int
    # Perpetuity inputs
    growth_rate: Optional[float] = None
    # Exit multiple inputs
    exit_multiple: Optional[float] = None
    exit_metric: Optional[str] = None
    metric_value: Optional[float] = None
    # Results
    terminal_value_fcff: float
    terminal_value_fcfe: float
    discount_rate_wacc: float
    discount_rate_ke: float
    pv_terminal_value_fcff: float
    pv_terminal_value_fcfe: float


class ValuationBridge(BaseModel):
    """Valuation bridge from EV to Equity."""
    sum_pv_fcff: float
    sum_pv_fcfe: float
    pv_terminal_value_fcff: float
    pv_terminal_value_fcfe: float
    enterprise_value: float
    debt_at_base: float
    cash_at_base: float
    net_debt: float
    equity_value_from_ev: float
    equity_value_direct: float
    reconciliation_difference: float
    reconciliation_notes: list[str]


class WACCDetails(BaseModel):
    """WACC computation details by year."""
    year: int
    ke: float
    rd: float
    tax_rate: float
    debt: float
    equity_book: float
    weight_debt: float
    weight_equity: float
    wacc: float


class DCFOutputs(BaseModel):
    """Complete DCF model outputs."""
    # Input summary
    base_year: int
    forecast_years: list[int]
    discounting_mode: DiscountingMode

    # Projections
    projections: list[YearlyProjection]
    nwc_schedule: list[YearlyNWC]
    cash_flows: list[YearlyCashFlow]

    # Discount rates
    ke: float
    wacc_details: list[WACCDetails]

    # Discounting
    discount_schedule: list[YearlyDiscount]

    # Terminal value
    terminal_value: TerminalValueOutput

    # Valuation
    valuation_bridge: ValuationBridge

    # Equity book roll-forward (if applicable)
    equity_book_values: Optional[dict[int, float]] = None
