"""
Biometano Project Finance Input Schema

Pydantic models for all Biometano case inputs with validation.
Structured for audit-trail friendly project finance modeling.
"""
from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field, field_validator, model_validator


# ============================================================================
# ENUMS
# ============================================================================

class GrantAccountingPolicy(str, Enum):
    """OIC 15 compliant grant accounting policies."""
    REDUCE_ASSET = "A1"      # Grant offsets capitalized cost / reduces fixed assets
    DEFERRED_INCOME = "A2"   # Grant recognized as liability, released over asset life


class TaxCreditPolicy(str, Enum):
    """OIC 25 compliant tax credit accounting policies."""
    REDUCE_TAX_EXPENSE = "B1"  # Immediate offset against taxes payable
    TAX_RECEIVABLE = "B2"      # Recognize as deferred tax asset if not yet usable


class GrantRecognitionTrigger(str, Enum):
    """When to recognize grant in accounting."""
    AT_COD = "at_cod"              # Commercial operation date
    AT_DECREE = "at_decree"        # When grant decree is issued
    AT_CASH_RECEIPT = "at_receipt" # When cash is received


# ============================================================================
# CONSTANTS
# ============================================================================

# Revenue channel display order (consistent across tables, Excel, charts)
REVENUE_CHANNEL_ORDER = ["Gate Fee", "Tariff", "GO", "CO₂", "Compost", "Total"]

# Default ZES tax credit rate (14.6189% of total CAPEX)
DEFAULT_ZES_CREDIT_RATE = 0.146189

# Default PNRR grant rate (40% of eligible CAPEX)
DEFAULT_PNRR_GRANT_RATE = 0.40


# ============================================================================
# HORIZON AND TIMELINE
# ============================================================================

class HorizonInputs(BaseModel):
    """Project timeline configuration."""
    base_year: int = Field(..., description="Valuation date (end of year)")
    years_forecast: int = Field(..., ge=1, le=30, description="Number of forecast years post-COD")
    construction_years: int = Field(1, ge=0, le=5, description="Years of construction before COD")
    cod_year: Optional[int] = Field(None, description="Commercial operation date year (computed if not set)")
    
    @model_validator(mode="after")
    def compute_cod_year(self) -> "HorizonInputs":
        if self.cod_year is None:
            self.cod_year = self.base_year + self.construction_years + 1
        return self
    
    @property
    def construction_years_list(self) -> list[int]:
        """Years during construction phase."""
        return list(range(self.base_year + 1, self.cod_year))
    
    @property
    def operating_years_list(self) -> list[int]:
        """Years during operating phase (post-COD)."""
        return list(range(self.cod_year, self.cod_year + self.years_forecast))
    
    @property
    def all_forecast_years(self) -> list[int]:
        """All forecast years (construction + operating)."""
        return self.construction_years_list + self.operating_years_list


# ============================================================================
# PRODUCTION AND CAPACITY
# ============================================================================

class ProductionInputs(BaseModel):
    """Production capacity and output configuration."""
    # Feedstock
    forsu_throughput_tpy: float = Field(..., gt=0, description="FORSU throughput (tonnes/year at full capacity)")
    impurity_rate: float = Field(0.20, ge=0, le=0.5, description="Impurity/sovvalli rate (default 20%)")
    
    # Biomethane output - user provides one of these
    biomethane_smc_y: Optional[float] = Field(None, gt=0, description="Biomethane output (Smc/year at full capacity)")
    biomethane_mwh_y: Optional[float] = Field(None, gt=0, description="Biomethane output (MWh/year at full capacity)")
    kwh_per_smc: float = Field(10.0, gt=0, description="Conversion factor kWh per Smc (default 10)")
    
    # Availability ramp-up (list per operating year, starting from COD)
    availability_profile: list[float] = Field(
        default=[0.75, 0.90, 0.95], 
        description="Availability ramp-up per operating year (0-1)"
    )
    
    # Byproducts at full capacity
    compost_tpy: float = Field(0, ge=0, description="Compost output (tonnes/year at full capacity)")
    co2_tpy: float = Field(0, ge=0, description="CO2 output (tonnes/year at full capacity)")
    
    @model_validator(mode="after")
    def validate_biomethane_output(self) -> "ProductionInputs":
        if self.biomethane_smc_y is None and self.biomethane_mwh_y is None:
            raise ValueError("Must provide either biomethane_smc_y or biomethane_mwh_y")
        return self
    
    @field_validator("availability_profile")
    @classmethod
    def validate_availability(cls, v: list[float]) -> list[float]:
        for a in v:
            if not 0 <= a <= 1:
                raise ValueError(f"Availability must be between 0 and 1, got {a}")
        return v
    
    def get_biomethane_mwh(self) -> float:
        """Get biomethane output in MWh/year at full capacity."""
        if self.biomethane_mwh_y is not None:
            return self.biomethane_mwh_y
        return self.biomethane_smc_y * self.kwh_per_smc / 1000  # Convert kWh to MWh
    
    def get_availability(self, operating_year_index: int) -> float:
        """Get availability for a given operating year (0-indexed from COD)."""
        if operating_year_index < len(self.availability_profile):
            return self.availability_profile[operating_year_index]
        # After ramp-up, use last value (steady state)
        return self.availability_profile[-1] if self.availability_profile else 0.95


# ============================================================================
# REVENUE CHANNELS
# ============================================================================

class RevenueChannelInputs(BaseModel):
    """Single revenue channel configuration."""
    price: float = Field(..., description="Base price per unit")
    payment_delay_days: int = Field(30, ge=0, le=365, description="Payment delay in days")
    escalation_rate: float = Field(0.0, ge=-0.1, le=0.2, description="Annual escalation rate")
    enabled: bool = Field(True, description="Whether this revenue channel is active")


class RevenuesInputs(BaseModel):
    """All revenue channel configurations."""
    # Gate fee (€/tonne FORSU)
    gate_fee: RevenueChannelInputs = Field(
        default_factory=lambda: RevenueChannelInputs(price=190.0, payment_delay_days=45)
    )
    
    # Incentive tariff / GME (€/MWh biomethane)
    tariff: RevenueChannelInputs = Field(
        default_factory=lambda: RevenueChannelInputs(price=70.16, payment_delay_days=90)
    )
    
    # CO2 sales (€/tonne)
    co2: RevenueChannelInputs = Field(
        default_factory=lambda: RevenueChannelInputs(price=120.0, payment_delay_days=45)
    )
    
    # Garanzie d'Origine (€/MWh)
    go: RevenueChannelInputs = Field(
        default_factory=lambda: RevenueChannelInputs(price=0.3, payment_delay_days=60)
    )
    
    # Compost (€/tonne)
    compost: RevenueChannelInputs = Field(
        default_factory=lambda: RevenueChannelInputs(price=5.0, payment_delay_days=30)
    )


# ============================================================================
# OPEX STRUCTURE
# ============================================================================

class OpexCategoryInputs(BaseModel):
    """Single OPEX category configuration."""
    fixed_annual: float = Field(0.0, ge=0, description="Fixed annual cost (€)")
    variable_per_tonne: float = Field(0.0, ge=0, description="Variable cost per tonne FORSU (€/t)")
    variable_per_mwh: float = Field(0.0, ge=0, description="Variable cost per MWh biomethane (€/MWh)")
    percent_of_capex: float = Field(0.0, ge=0, le=1, description="Annual cost as % of total CAPEX")
    escalation_rate: float = Field(0.0, ge=-0.1, le=0.2, description="Annual escalation rate")
    payment_delay_days: int = Field(30, ge=0, le=365, description="Payment delay in days")
    ramp_up_profile: Optional[list[float]] = Field(
        None, description="Cost ramp-up per operating year (0-1), defaults to availability"
    )


class OpexInputs(BaseModel):
    """All OPEX category configurations."""
    feedstock_handling: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    utilities: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    chemicals: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    maintenance: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    personnel: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    insurance: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    overheads: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    digestate_handling: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    other: OpexCategoryInputs = Field(default_factory=OpexCategoryInputs)
    
    def all_categories(self) -> dict[str, OpexCategoryInputs]:
        """Return all OPEX categories as a dict."""
        return {
            "feedstock_handling": self.feedstock_handling,
            "utilities": self.utilities,
            "chemicals": self.chemicals,
            "maintenance": self.maintenance,
            "personnel": self.personnel,
            "insurance": self.insurance,
            "overheads": self.overheads,
            "digestate_handling": self.digestate_handling,
            "other": self.other,
        }


# ============================================================================
# CAPEX STRUCTURE
# ============================================================================

class CapexLineItem(BaseModel):
    """Single CAPEX line item."""
    amount: float = Field(..., ge=0, description="Total amount (€)")
    spend_profile: list[float] = Field(
        default=[1.0], 
        description="Spend profile across construction years (must sum to 1)"
    )
    capitalize: bool = Field(True, description="Whether to capitalize this item")
    useful_life_years: int = Field(20, ge=1, le=50, description="Useful life for depreciation")
    eligible_for_grant: bool = Field(True, description="Whether eligible for capital grant")
    eligible_for_tax_credit: bool = Field(False, description="Whether eligible for tax credit")
    
    @field_validator("spend_profile")
    @classmethod
    def validate_spend_profile(cls, v: list[float]) -> list[float]:
        total = sum(v)
        if abs(total - 1.0) > 0.001:
            raise ValueError(f"Spend profile must sum to 1.0, got {total}")
        return v


class CapexInputs(BaseModel):
    """All CAPEX line items."""
    epc: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    civils: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    upgrading_unit: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    grid_connection: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    engineering: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    permitting: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    contingency: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    startup_costs: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0, capitalize=False))
    other: CapexLineItem = Field(default_factory=lambda: CapexLineItem(amount=0))
    
    # Optional toggles
    capitalize_idc: bool = Field(False, description="Capitalize interest during construction")
    
    def all_items(self) -> dict[str, CapexLineItem]:
        """Return all CAPEX items as a dict."""
        return {
            "epc": self.epc,
            "civils": self.civils,
            "upgrading_unit": self.upgrading_unit,
            "grid_connection": self.grid_connection,
            "engineering": self.engineering,
            "permitting": self.permitting,
            "contingency": self.contingency,
            "startup_costs": self.startup_costs,
            "other": self.other,
        }
    
    def total_capex(self) -> float:
        """Total CAPEX amount."""
        return sum(item.amount for item in self.all_items().values())
    
    def eligible_for_grant(self) -> float:
        """Total CAPEX eligible for grant."""
        return sum(item.amount for item in self.all_items().values() if item.eligible_for_grant)


# ============================================================================
# FINANCING
# ============================================================================

class FinancingInputs(BaseModel):
    """Project financing configuration."""
    # Debt
    debt_amount: float = Field(0, ge=0, description="Initial debt amount (€)")
    debt_drawdown_profile: list[float] = Field(
        default=[1.0],
        description="Debt drawdown across construction years (must sum to 1)"
    )
    cost_of_debt: float | dict[int, float] = Field(0.05, description="Cost of debt (rd)")
    debt_repayment_years: int = Field(15, ge=0, le=30, description="Years to repay debt from COD")
    
    # Equity
    cash_at_base: float = Field(0, ge=0, description="Cash at valuation date")
    equity_book_at_base: float = Field(0, ge=0, description="Equity book value at valuation date")
    
    # Tax
    tax_rate: float = Field(0.24, ge=0, le=0.5, description="Corporate tax rate")
    
    # CAPM for Ke
    rf: float = Field(0.03, description="Risk-free rate")
    rm: float = Field(0.08, description="Market return")
    beta: float = Field(1.0, ge=0, description="Equity beta")
    ke_override: Optional[float] = Field(None, description="Direct Ke override")
    
    # WACC
    target_we: Optional[float] = Field(None, ge=0, le=1, description="Target equity weight")
    use_book_weights: bool = Field(False, description="Use book-value weights for WACC")
    
    def get_rd(self, year: int) -> float:
        """Get cost of debt for a specific year."""
        if isinstance(self.cost_of_debt, dict):
            return self.cost_of_debt.get(year, list(self.cost_of_debt.values())[-1])
        return self.cost_of_debt


# ============================================================================
# INCENTIVES
# ============================================================================

class CapitalGrantInputs(BaseModel):
    """Capital grant (contributo in conto impianti) configuration."""
    enabled: bool = Field(False, description="Whether capital grant is active")
    amount: Optional[float] = Field(None, ge=0, description="Fixed grant amount (€)")
    percent_of_eligible: Optional[float] = Field(None, ge=0, le=1, description="Grant as % of eligible CAPEX")
    
    # Recognition and cash timing
    recognition_trigger: GrantRecognitionTrigger = Field(
        GrantRecognitionTrigger.AT_COD, 
        description="When to recognize grant"
    )
    cash_receipt_schedule: list[float] = Field(
        default=[1.0],
        description="Cash receipt profile (year 1 = COD year, sum to 1)"
    )
    
    # Accounting policy
    accounting_policy: GrantAccountingPolicy = Field(
        GrantAccountingPolicy.DEFERRED_INCOME,
        description="OIC 15 accounting treatment"
    )
    
    # For deferred income release
    release_years: Optional[int] = Field(
        None, 
        description="Years to release deferred income (defaults to asset useful life)"
    )
    
    @model_validator(mode="after")
    def validate_grant_amount(self) -> "CapitalGrantInputs":
        if self.enabled and self.amount is None and self.percent_of_eligible is None:
            raise ValueError("Must provide either amount or percent_of_eligible when grant is enabled")
        return self


class TaxCreditInputs(BaseModel):
    """Tax credit (credito d'imposta / ZES) configuration."""
    enabled: bool = Field(False, description="Whether tax credit is active")
    amount: Optional[float] = Field(None, ge=0, description="Fixed tax credit amount (€)")
    percent_of_eligible: Optional[float] = Field(
        None, ge=0, le=1, 
        description="Credit as % of eligible base (default: 14.6189% ZES)"
    )
    eligible_base: str = Field("capex", description="Base for percentage: 'capex' or 'opex'")
    
    # Usability
    usable_from_year: Optional[int] = Field(None, description="Year when credit becomes usable")
    carry_forward_years: int = Field(5, ge=0, le=20, description="Years to carry forward unused credit")
    annual_cap_percent: float = Field(1.0, ge=0, le=1, description="Max % of tax liability that can be offset per year")
    
    # Accounting policy
    accounting_policy: TaxCreditPolicy = Field(
        TaxCreditPolicy.REDUCE_TAX_EXPENSE,
        description="OIC 25 accounting treatment"
    )
    
    @model_validator(mode="after")
    def validate_credit_amount(self) -> "TaxCreditInputs":
        if self.enabled:
            # If neither provided, use default ZES rate
            if self.amount is None and self.percent_of_eligible is None:
                self.percent_of_eligible = DEFAULT_ZES_CREDIT_RATE
            # If both provided, validate consistency (within 1% tolerance)
            elif self.amount is not None and self.percent_of_eligible is not None:
                # Note: Can't validate here without capex - will validate in builder
                pass
        return self


class IncentivesInputs(BaseModel):
    """All incentives configuration (legacy - kept for backwards compatibility)."""
    capital_grant: CapitalGrantInputs = Field(default_factory=CapitalGrantInputs)
    tax_credit: TaxCreditInputs = Field(default_factory=TaxCreditInputs)


# ============================================================================
# INCENTIVES POLICY (PNRR + ZES WATERFALL ALLOCATION)
# ============================================================================

class IncentivesProfile(str, Enum):
    """Incentives profile determining ESL intensity."""
    MEDIA_IMPRESA_SICILIA = "media_impresa_sicilia"
    GRANDE_IMPRESA = "grande_impresa"
    CUSTOM = "custom"


class ZesEligibility(str, Enum):
    """ZES eligibility status for CAPEX line items."""
    ELIGIBLE = "eligible"
    NOT_ELIGIBLE = "not_eligible"
    PARTIAL = "partial"


class PnrrPolicyInputs(BaseModel):
    """PNRR grant policy configuration for capacity-based eligible spend."""
    grant_rate: float = Field(0.40, ge=0, le=1, description="Grant rate (40%)")
    annual_hours: int = Field(8000, gt=0, description="Operating hours per year")
    biomethane_smc_per_year: float = Field(..., gt=0, description="Biomethane Smc/year")
    cs_max_base_eur_per_smcph: float = Field(50000, gt=0, description="Max specific cost €/Smc/h base")
    inflation_adj_factor: float = Field(1.137, gt=0, description="NIC inflation adjustment")
    tech_costs_cap_pct: float = Field(0.12, ge=0, le=1, description="Tech costs cap as % of eligible spend")


class ZesPolicyInputs(BaseModel):
    """ZES tax credit policy configuration with riparto."""
    theoretical_rate: float = Field(0.50, ge=0, le=1, description="ZES nominal rate (50% for media impresa)")
    riparto_coeff: float = Field(0.6038, ge=0, le=1, description="Riparto coefficient (stress case)")
    allow_overlap_with_pnrr: bool = Field(True, description="Allow ZES base to overlap with PNRR-covered CAPEX")
    enforce_no_double_financing: bool = Field(True, description="Enforce ESL cap")
    base_allocation_strategy: str = Field("waterfall", description="Allocation strategy")
    tech_costs_eligibility_share: float = Field(0.50, ge=0, le=1, description="Share of tech costs ZES-eligible")


class CapexBreakdownItem(BaseModel):
    """Single CAPEX line item with eligibility flags for incentives allocation."""
    name: str = Field(..., description="Line item name")
    amount: float = Field(..., ge=0, description="Amount in €")
    oic_class: str = Field("", description="OIC accounting class (e.g., B.II.1)")
    pnrr_eligible: bool = Field(True, description="PNRR eligible")
    zes_eligible: str = Field("eligible", description="ZES eligibility: 'eligible', 'not_eligible', or 'partial'")
    notes: str = Field("", description="Notes")
    
    @field_validator("zes_eligible")
    @classmethod
    def validate_zes_eligibility(cls, v: str) -> str:
        valid = ["eligible", "not_eligible", "partial", "true", "false"]
        if v.lower() not in valid:
            raise ValueError(f"zes_eligible must be one of {valid}")
        return v.lower()


class IncentivesPolicyInputs(BaseModel):
    """
    Incentives policy for PNRR + ZES waterfall allocation.
    
    This is the new model that implements:
    - PNRR grant from capacity-based eligible spend cap
    - ZES tax credit as residual ESL gap
    - Waterfall allocation of ZES base
    - Riparto coefficient for nominal vs cash benefit
    """
    enabled: bool = Field(False, description="Use new incentives policy calculation")
    profile: IncentivesProfile = Field(
        IncentivesProfile.MEDIA_IMPRESA_SICILIA, 
        description="Incentives profile"
    )
    max_esl_intensity: float = Field(0.50, ge=0, le=1, description="Max ESL aid intensity")
    
    # PNRR configuration
    pnrr: Optional[PnrrPolicyInputs] = Field(None, description="PNRR policy params")
    
    # ZES configuration
    zes: Optional[ZesPolicyInputs] = Field(None, description="ZES policy params")
    
    # CAPEX breakdown with per-line eligibility
    capex_breakdown: list[CapexBreakdownItem] = Field(
        default_factory=list,
        description="CAPEX breakdown with eligibility flags"
    )
    
    @model_validator(mode="after")
    def validate_policy(self) -> "IncentivesPolicyInputs":
        if self.enabled:
            if self.pnrr is None:
                raise ValueError("PNRR config required when incentives_policy is enabled")
            if not self.capex_breakdown:
                raise ValueError("capex_breakdown required when incentives_policy is enabled")
        return self


# ============================================================================
# TERMINAL VALUE
# ============================================================================

class TerminalValueInputs(BaseModel):
    """Terminal value configuration."""
    method: str = Field("perpetuity", description="'perpetuity' or 'exit_multiple'")
    perpetuity_growth: float = Field(0.0, ge=-0.05, le=0.05, description="Perpetuity growth rate")
    exit_multiple: Optional[float] = Field(None, gt=0, description="Exit EBITDA multiple")


# ============================================================================
# MAIN CASE MODEL
# ============================================================================

class BiometanoCase(BaseModel):
    """Complete Biometano project finance case."""
    # Required sections
    horizon: HorizonInputs
    production: ProductionInputs
    
    # Optional sections with defaults
    revenues: RevenuesInputs = Field(default_factory=RevenuesInputs)
    opex: OpexInputs = Field(default_factory=OpexInputs)
    capex: CapexInputs = Field(default_factory=CapexInputs)
    financing: FinancingInputs = Field(default_factory=FinancingInputs)
    incentives: IncentivesInputs = Field(default_factory=IncentivesInputs)
    incentives_policy: IncentivesPolicyInputs = Field(default_factory=IncentivesPolicyInputs)
    terminal_value: TerminalValueInputs = Field(default_factory=TerminalValueInputs)
    
    # Model config
    model_config = {"extra": "forbid"}
