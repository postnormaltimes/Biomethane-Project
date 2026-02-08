"""
Incentives Allocation Module

Implements PNRR grant and ZES tax credit calculations with:
- Capacity-based PNRR eligible spend cap
- ESL gap calculation for ZES nominal credit
- Waterfall allocation of ZES base (over-cap first, then overlap)
- Riparto coefficient for nominal vs cash ZES benefit

Reference formulas and expected anchors (32M total CAPEX, 4M Smc/y):
- EligibleSpend_PNRR = 28,425,000
- Grant_PNRR = 11,370,000
- Gap_ZES_Nominal = 4,630,000
- ZES_Base_Required = 9,260,000
- ZES_Cash_Benefit ≈ 2,795,594 (riparto 60.38%)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional
import math


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


@dataclass
class PnrrParams:
    """PNRR grant configuration parameters."""
    grant_rate: float = 0.40  # 40% grant rate
    annual_hours: int = 8000  # Operating hours per year
    biomethane_smc_per_year: float = 4_000_000  # Smc/year
    cs_max_base_eur_per_smcph: float = 50_000  # €/Smc/h base
    inflation_adj_factor: float = 1.137  # NIC adjustment
    tech_costs_cap_pct: float = 0.12  # 12% cap on technical costs
    
    @property
    def smc_per_hour(self) -> float:
        """Compute Smc/h capacity."""
        return self.biomethane_smc_per_year / self.annual_hours
    
    @property
    def cs_max_adjusted(self) -> float:
        """Compute inflation-adjusted max specific cost."""
        return self.cs_max_base_eur_per_smcph * self.inflation_adj_factor
    
    @property
    def eligible_spend_cap(self) -> float:
        """Compute PNRR eligible spend cap."""
        return self.smc_per_hour * self.cs_max_adjusted
    
    @property
    def tech_costs_limit(self) -> float:
        """Compute technical costs limit."""
        return self.eligible_spend_cap * self.tech_costs_cap_pct


@dataclass
class ZesParams:
    """ZES tax credit configuration parameters."""
    theoretical_rate: float = 0.50  # 50% for media impresa in Sicily
    riparto_coeff: float = 0.6038  # Stress case riparto coefficient
    allow_overlap_with_pnrr: bool = True
    enforce_no_double_financing: bool = True
    base_allocation_strategy: str = "waterfall"
    tech_costs_eligibility_share: float = 0.50  # For "partial" items


@dataclass
class CapexLineItem:
    """Single CAPEX line item with eligibility flags."""
    name: str
    amount: float
    oic_class: str = ""
    pnrr_eligible: bool = True
    zes_eligible: ZesEligibility | bool = True
    notes: str = ""
    
    def __post_init__(self):
        # Normalize zes_eligible to ZesEligibility enum
        if isinstance(self.zes_eligible, bool):
            self.zes_eligible = ZesEligibility.ELIGIBLE if self.zes_eligible else ZesEligibility.NOT_ELIGIBLE
        elif isinstance(self.zes_eligible, str):
            if self.zes_eligible.lower() == "partial":
                self.zes_eligible = ZesEligibility.PARTIAL
            elif self.zes_eligible.lower() in ("true", "yes", "eligible"):
                self.zes_eligible = ZesEligibility.ELIGIBLE
            else:
                self.zes_eligible = ZesEligibility.NOT_ELIGIBLE
    
    @property
    def is_zes_eligible(self) -> bool:
        """Full ZES eligibility."""
        return self.zes_eligible == ZesEligibility.ELIGIBLE
    
    @property
    def is_zes_partial(self) -> bool:
        """Partial ZES eligibility."""
        return self.zes_eligible == ZesEligibility.PARTIAL


@dataclass
class ZesBaseAllocation:
    """Allocation of ZES base to a single CAPEX line."""
    line_name: str
    line_amount: float
    zes_eligible: ZesEligibility
    allocated_from_overcap: float = 0.0
    allocated_from_overlap: float = 0.0
    
    @property
    def total_allocated(self) -> float:
        return self.allocated_from_overcap + self.allocated_from_overlap


@dataclass
class IncentiveAllocationResult:
    """Complete result of incentive allocation calculations."""
    # PNRR calculations
    smc_per_hour: float
    cs_max_base: float
    cs_max_adjusted: float
    eligible_spend_pnrr: float
    tech_costs_limit: float
    tech_costs_actual: float
    tech_costs_eligible: float
    tech_costs_warning: bool
    grant_rate: float
    grant_pnrr: float
    
    # ESL and ZES gap
    total_capex: float
    max_esl_intensity: float
    max_aid_amount: float
    gap_zes_nominal: float
    zes_rate: float
    zes_base_required: float
    
    # Waterfall allocation
    over_cap_total: float
    over_cap_zes_eligible: float
    zes_base_from_overcap: float
    zes_base_from_overlap: float
    
    # Riparto
    riparto_coeff: float
    zes_nominal_authorized: float
    zes_cash_benefit: float
    zes_benefit_lost: float
    
    # Compliance
    nominal_aid_intensity: float
    connection_excluded: bool
    compliance_pass: bool
    
    # Field with default (must come last)
    allocation_details: list[ZesBaseAllocation] = field(default_factory=list)
    
    @property
    def total_nominal_aid(self) -> float:
        """Total nominal aid (PNRR + ZES nominal)."""
        return self.grant_pnrr + self.zes_nominal_authorized
    
    @property
    def total_cash_benefit(self) -> float:
        """Total cash benefit (PNRR grant + ZES cash)."""
        return self.grant_pnrr + self.zes_cash_benefit


def compute_pnrr_eligible_spend(pnrr: PnrrParams) -> float:
    """
    Compute PNRR eligible spend cap based on capacity.
    
    Formula: Smc/h × CS_max_adjusted
    
    Expected: 500 × 56,850 = 28,425,000
    """
    return pnrr.eligible_spend_cap


def compute_pnrr_grant(pnrr: PnrrParams, eligible_spend: Optional[float] = None) -> float:
    """
    Compute PNRR grant amount.
    
    Formula: EligibleSpend × grant_rate
    
    Expected: 28,425,000 × 0.40 = 11,370,000
    """
    if eligible_spend is None:
        eligible_spend = compute_pnrr_eligible_spend(pnrr)
    return eligible_spend * pnrr.grant_rate


def compute_esl_gap(
    total_capex: float,
    max_esl_intensity: float,
    pnrr_grant: float,
) -> float:
    """
    Compute ZES nominal gap (residual needed to reach ESL cap).
    
    Formula: (TotalCAPEX × ESL_intensity) - PNRR_grant
    
    Expected: (32,000,000 × 0.50) - 11,370,000 = 4,630,000
    """
    max_aid = total_capex * max_esl_intensity
    return max(0, max_aid - pnrr_grant)


def compute_zes_base_required(gap_zes_nominal: float, zes_rate: float) -> float:
    """
    Compute ZES base required to generate the nominal gap.
    
    Formula: Gap_ZES_Nominal / ZES_rate
    
    Expected: 4,630,000 / 0.50 = 9,260,000
    """
    if zes_rate <= 0:
        return 0.0
    return gap_zes_nominal / zes_rate


def allocate_zes_base_waterfall(
    capex_lines: list[CapexLineItem],
    eligible_spend_pnrr: float,
    zes_base_required: float,
    zes_params: ZesParams,
) -> tuple[float, float, list[ZesBaseAllocation]]:
    """
    Allocate ZES base using waterfall strategy:
    1. First use over-cap CAPEX (costs above PNRR eligible spend)
    2. Then use overlap with PNRR-covered if allowed
    
    Returns:
        (zes_from_overcap, zes_from_overlap, allocation_details)
    
    Expected:
        over_cap = 32,000,000 - 28,425,000 = 3,575,000
        zes_from_overcap = 3,575,000
        zes_from_overlap = 9,260,000 - 3,575,000 = 5,685,000
    """
    total_capex = sum(line.amount for line in capex_lines)
    over_cap_total = max(0, total_capex - eligible_spend_pnrr)
    
    allocations = []
    remaining_base = zes_base_required
    zes_from_overcap = 0.0
    zes_from_overlap = 0.0
    
    # Calculate ZES-eligible amount for each line
    def get_zes_eligible_amount(line: CapexLineItem) -> float:
        if line.zes_eligible == ZesEligibility.ELIGIBLE:
            return line.amount
        elif line.zes_eligible == ZesEligibility.PARTIAL:
            return line.amount * zes_params.tech_costs_eligibility_share
        return 0.0
    
    # Sort lines: prioritize non-PNRR-eligible items (pure over-cap) first
    # Then PNRR-eligible items for overlap
    sorted_lines = sorted(
        capex_lines,
        key=lambda x: (x.pnrr_eligible, -get_zes_eligible_amount(x))
    )
    
    # Phase 1: Allocate from over-cap portion (not PNRR-eligible items)
    # and the portion of PNRR-eligible items that exceeds the cap
    
    # Calculate how much of total CAPEX is "over cap"
    pnrr_eligible_total = sum(line.amount for line in capex_lines if line.pnrr_eligible)
    
    # The over-cap is: total - eligible_spend_pnrr
    # This comes from either non-PNRR items or the excess of PNRR items
    
    for line in sorted_lines:
        zes_eligible_amount = get_zes_eligible_amount(line)
        if zes_eligible_amount <= 0:
            allocations.append(ZesBaseAllocation(
                line_name=line.name,
                line_amount=line.amount,
                zes_eligible=line.zes_eligible if isinstance(line.zes_eligible, ZesEligibility) else ZesEligibility.NOT_ELIGIBLE,
            ))
            continue
        
        alloc = ZesBaseAllocation(
            line_name=line.name,
            line_amount=line.amount,
            zes_eligible=line.zes_eligible if isinstance(line.zes_eligible, ZesEligibility) else ZesEligibility.ELIGIBLE,
        )
        
        if remaining_base <= 0:
            allocations.append(alloc)
            continue
        
        # Determine how much of this line is "over-cap"
        if not line.pnrr_eligible:
            # Non-PNRR lines are fully over-cap
            over_cap_portion = zes_eligible_amount
        else:
            # For PNRR-eligible lines, calculate their share of over-cap
            # Assume proportional distribution of over-cap across PNRR-eligible items
            if pnrr_eligible_total > eligible_spend_pnrr:
                line_over_cap_ratio = (pnrr_eligible_total - eligible_spend_pnrr) / pnrr_eligible_total
                over_cap_portion = zes_eligible_amount * line_over_cap_ratio
            else:
                over_cap_portion = 0
        
        # Allocate from over-cap
        if over_cap_portion > 0 and remaining_base > 0:
            from_overcap = min(over_cap_portion, remaining_base)
            alloc.allocated_from_overcap = from_overcap
            zes_from_overcap += from_overcap
            remaining_base -= from_overcap
        
        # Allocate from overlap (if allowed and still needed)
        if remaining_base > 0 and zes_params.allow_overlap_with_pnrr:
            overlap_portion = zes_eligible_amount - (alloc.allocated_from_overcap if alloc.allocated_from_overcap else 0)
            if overlap_portion > 0:
                from_overlap = min(overlap_portion, remaining_base)
                alloc.allocated_from_overlap = from_overlap
                zes_from_overlap += from_overlap
                remaining_base -= from_overlap
        
        allocations.append(alloc)
    
    return zes_from_overcap, zes_from_overlap, allocations


def compute_zes_cash_benefit(zes_nominal: float, riparto_coeff: float) -> float:
    """
    Compute effective ZES cash benefit after riparto.
    
    Formula: ZES_Nominal × riparto_coeff
    
    Expected: 4,630,000 × 0.6038 ≈ 2,795,594
    """
    return zes_nominal * riparto_coeff


def compute_full_incentive_allocation(
    capex_lines: list[CapexLineItem],
    pnrr: PnrrParams,
    zes: ZesParams,
    max_esl_intensity: float = 0.50,
    profile: IncentivesProfile = IncentivesProfile.MEDIA_IMPRESA_SICILIA,
) -> IncentiveAllocationResult:
    """
    Compute complete incentive allocation including PNRR grant, ZES gap,
    waterfall allocation, and riparto.
    
    Args:
        capex_lines: List of CAPEX line items with eligibility flags
        pnrr: PNRR configuration parameters
        zes: ZES configuration parameters
        max_esl_intensity: Maximum ESL aid intensity (0.50 for media impresa Sicily)
        profile: Incentives profile
    
    Returns:
        Complete IncentiveAllocationResult
    """
    # Total CAPEX
    total_capex = sum(line.amount for line in capex_lines)
    
    # PNRR calculations
    smc_h = pnrr.smc_per_hour
    cs_max_adj = pnrr.cs_max_adjusted
    eligible_spend = compute_pnrr_eligible_spend(pnrr)
    
    # Tech costs check
    tech_lines = [line for line in capex_lines if "tecnic" in line.name.lower() or "spese" in line.name.lower()]
    tech_costs_actual = sum(line.amount for line in tech_lines)
    tech_costs_limit = pnrr.tech_costs_limit
    tech_costs_eligible = min(tech_costs_actual, tech_costs_limit)
    tech_costs_warning = tech_costs_actual > tech_costs_limit
    
    # PNRR grant
    grant_pnrr = compute_pnrr_grant(pnrr, eligible_spend)
    
    # ESL and ZES gap
    max_aid = total_capex * max_esl_intensity
    gap_zes_nominal = compute_esl_gap(total_capex, max_esl_intensity, grant_pnrr)
    zes_base_required = compute_zes_base_required(gap_zes_nominal, zes.theoretical_rate)
    
    # Waterfall allocation
    over_cap_total = max(0, total_capex - eligible_spend)
    over_cap_zes_eligible = sum(
        line.amount for line in capex_lines 
        if not line.pnrr_eligible and line.is_zes_eligible
    )
    
    zes_from_overcap, zes_from_overlap, allocation_details = allocate_zes_base_waterfall(
        capex_lines, eligible_spend, zes_base_required, zes
    )
    
    # Riparto
    zes_nominal = gap_zes_nominal
    zes_cash = compute_zes_cash_benefit(zes_nominal, zes.riparto_coeff)
    zes_lost = zes_nominal - zes_cash
    
    # Compliance
    total_nominal_aid = grant_pnrr + zes_nominal
    nominal_intensity = total_nominal_aid / total_capex if total_capex > 0 else 0
    
    # Check connection excluded from ZES
    connection_lines = [line for line in capex_lines if "connessione" in line.name.lower() or "connection" in line.name.lower() or "rete" in line.name.lower()]
    connection_excluded = all(
        line.zes_eligible == ZesEligibility.NOT_ELIGIBLE or not line.is_zes_eligible
        for line in connection_lines
    ) if connection_lines else True
    
    compliance_pass = (
        abs(nominal_intensity - max_esl_intensity) < 0.001 and  # Within tolerance
        connection_excluded
    )
    
    return IncentiveAllocationResult(
        # PNRR
        smc_per_hour=smc_h,
        cs_max_base=pnrr.cs_max_base_eur_per_smcph,
        cs_max_adjusted=cs_max_adj,
        eligible_spend_pnrr=eligible_spend,
        tech_costs_limit=tech_costs_limit,
        tech_costs_actual=tech_costs_actual,
        tech_costs_eligible=tech_costs_eligible,
        tech_costs_warning=tech_costs_warning,
        grant_rate=pnrr.grant_rate,
        grant_pnrr=grant_pnrr,
        # ESL/ZES
        total_capex=total_capex,
        max_esl_intensity=max_esl_intensity,
        max_aid_amount=max_aid,
        gap_zes_nominal=gap_zes_nominal,
        zes_rate=zes.theoretical_rate,
        zes_base_required=zes_base_required,
        # Waterfall
        over_cap_total=over_cap_total,
        over_cap_zes_eligible=over_cap_zes_eligible,
        zes_base_from_overcap=zes_from_overcap,
        zes_base_from_overlap=zes_from_overlap,
        allocation_details=allocation_details,
        # Riparto
        riparto_coeff=zes.riparto_coeff,
        zes_nominal_authorized=zes_nominal,
        zes_cash_benefit=zes_cash,
        zes_benefit_lost=zes_lost,
        # Compliance
        nominal_aid_intensity=nominal_intensity,
        connection_excluded=connection_excluded,
        compliance_pass=compliance_pass,
    )


# Convenience function for default 32M case
def compute_default_case_allocation() -> IncentiveAllocationResult:
    """
    Compute allocation for the default 32M CAPEX case.
    
    Returns result with expected anchors:
    - EligibleSpend_PNRR = 28,425,000
    - Grant_PNRR = 11,370,000
    - Gap_ZES_Nominal = 4,630,000
    - ZES_Cash_Benefit ≈ 2,795,594
    """
    capex_lines = [
        CapexLineItem("Opere Civili", 9_500_000, "B.II.1", True, True),
        CapexLineItem("Impianto Processo", 13_500_000, "B.II.2", True, True),
        CapexLineItem("Upgrading", 5_500_000, "B.II.2", True, True),
        CapexLineItem("Connessione Rete", 1_500_000, "B.I.7", True, False),  # ZES ineligible
        CapexLineItem("Spese Tecniche", 2_000_000, "Capitalizzate", True, "partial"),
    ]
    
    pnrr = PnrrParams()
    zes = ZesParams()
    
    return compute_full_incentive_allocation(capex_lines, pnrr, zes)
