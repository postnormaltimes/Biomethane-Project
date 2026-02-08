"""
Biometano Accounting Module

OIC-compliant accounting treatment for incentives:
- OIC 15: Grant receivables and capital grants
- OIC 25: Tax credits and deferred tax assets

Implements configurable policies for:
- A1: Grant reduces asset cost
- A2: Grant as deferred income
- B1: Tax credit reduces tax expense
- B2: Tax credit as receivable/DTA
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from dcf_projects.biometano.schema import (
    BiometanoCase,
    GrantAccountingPolicy,
    TaxCreditPolicy,
    GrantRecognitionTrigger,
)


@dataclass
class FixedAssetSchedule:
    """Fixed asset roll-forward schedule."""
    year: int
    opening_gross: float = 0.0
    additions: float = 0.0            # CAPEX additions
    grant_reduction: float = 0.0       # Policy A1: grant reduces asset
    closing_gross: float = 0.0
    opening_accum_depr: float = 0.0
    depreciation: float = 0.0
    closing_accum_depr: float = 0.0
    net_book_value: float = 0.0


@dataclass
class DeferredIncomeSchedule:
    """Deferred income (grant) schedule for Policy A2."""
    year: int
    opening_balance: float = 0.0
    grant_received: float = 0.0
    release_to_pl: float = 0.0
    closing_balance: float = 0.0


@dataclass
class TaxCreditSchedule:
    """Tax credit utilization schedule."""
    year: int
    opening_balance: float = 0.0
    credit_available: float = 0.0     # New credit becoming available
    utilization: float = 0.0          # Used against taxes
    expiry: float = 0.0               # Expired unused credits
    closing_balance: float = 0.0


@dataclass
class GrantReceivableSchedule:
    """Grant receivable from PA (OIC 15)."""
    year: int
    opening_balance: float = 0.0
    grant_recognized: float = 0.0     # Recognition per trigger
    cash_received: float = 0.0
    closing_balance: float = 0.0


@dataclass
class AccountingOutputs:
    """Complete accounting outputs for statements."""
    fixed_assets: list[FixedAssetSchedule] = field(default_factory=list)
    deferred_income: list[DeferredIncomeSchedule] = field(default_factory=list)
    tax_credits: list[TaxCreditSchedule] = field(default_factory=list)
    grant_receivables: list[GrantReceivableSchedule] = field(default_factory=list)
    
    # Summary values
    total_grant_amount: float = 0.0
    total_tax_credit: float = 0.0
    
    def get_fixed_asset(self, year: int) -> Optional[FixedAssetSchedule]:
        for fa in self.fixed_assets:
            if fa.year == year:
                return fa
        return None
    
    def get_deferred_income(self, year: int) -> Optional[DeferredIncomeSchedule]:
        for di in self.deferred_income:
            if di.year == year:
                return di
        return None
    
    def get_tax_credit(self, year: int) -> Optional[TaxCreditSchedule]:
        for tc in self.tax_credits:
            if tc.year == year:
                return tc
        return None
    
    def get_grant_receivable(self, year: int) -> Optional[GrantReceivableSchedule]:
        for gr in self.grant_receivables:
            if gr.year == year:
                return gr
        return None


class AccountingCalculator:
    """
    Calculates accounting schedules for incentives.
    
    Implements OIC 15 (receivables) and OIC 25 (tax credits) logic
    with configurable policy toggles.
    """
    
    def __init__(self, case: BiometanoCase):
        self.case = case
        self.horizon = case.horizon
        self.capex = case.capex
        self.incentives = case.incentives
        
    def compute(
        self,
        capex_by_year: dict[int, float],
        taxes_before_credit: dict[int, float],
    ) -> AccountingOutputs:
        """
        Compute all accounting schedules.
        
        Args:
            capex_by_year: CAPEX spend by year from builder
            taxes_before_credit: Tax liability before credits by year
            
        Returns:
            AccountingOutputs with all schedules
        """
        outputs = AccountingOutputs()
        
        # Compute grant amount
        grant = self.incentives.capital_grant
        if grant.enabled:
            if grant.amount is not None:
                outputs.total_grant_amount = grant.amount
            else:
                eligible_base = self.capex.eligible_for_grant()
                outputs.total_grant_amount = eligible_base * (grant.percent_of_eligible or 0)
        
        # Compute tax credit amount
        tax_credit = self.incentives.tax_credit
        if tax_credit.enabled:
            if tax_credit.amount is not None:
                outputs.total_tax_credit = tax_credit.amount
            else:
                if tax_credit.eligible_base == "capex":
                    eligible_base = self.capex.total_capex()
                else:
                    eligible_base = 0  # Would need OPEX totals
                outputs.total_tax_credit = eligible_base * (tax_credit.percent_of_eligible or 0)
        
        # Build schedules
        outputs.fixed_assets = self._build_fixed_asset_schedule(
            capex_by_year, outputs.total_grant_amount
        )
        outputs.deferred_income = self._build_deferred_income_schedule(
            outputs.total_grant_amount
        )
        outputs.grant_receivables = self._build_grant_receivable_schedule(
            outputs.total_grant_amount
        )
        outputs.tax_credits = self._build_tax_credit_schedule(
            outputs.total_tax_credit, taxes_before_credit
        )
        
        return outputs
    
    def _build_fixed_asset_schedule(
        self,
        capex_by_year: dict[int, float],
        grant_amount: float,
    ) -> list[FixedAssetSchedule]:
        """Build fixed asset roll-forward with grant effect."""
        schedules = []
        all_years = [self.horizon.base_year] + self.horizon.all_forecast_years
        
        grant = self.incentives.capital_grant
        grant_reduction_applied = False
        
        # Compute weighted average useful life for depreciation
        total_capex = self.capex.total_capex()
        if total_capex > 0:
            weighted_life = sum(
                item.amount * item.useful_life_years 
                for item in self.capex.all_items().values()
            ) / total_capex
        else:
            weighted_life = 20.0
        
        prev_gross = 0.0
        prev_accum_depr = 0.0
        
        for year in all_years:
            schedule = FixedAssetSchedule(year=year)
            schedule.opening_gross = prev_gross
            schedule.opening_accum_depr = prev_accum_depr
            
            # Additions from CAPEX
            schedule.additions = capex_by_year.get(year, 0.0)
            
            # Grant reduction (Policy A1) - applied at recognition year
            if (grant.enabled and 
                grant.accounting_policy == GrantAccountingPolicy.REDUCE_ASSET and
                not grant_reduction_applied):
                
                recognition_year = self._get_grant_recognition_year()
                if year == recognition_year:
                    schedule.grant_reduction = grant_amount
                    grant_reduction_applied = True
            
            schedule.closing_gross = (
                schedule.opening_gross + 
                schedule.additions - 
                schedule.grant_reduction
            )
            
            # Depreciation (straight-line on net depreciable base)
            if schedule.closing_gross > 0 and year >= self.horizon.cod_year:
                # Depreciate over remaining useful life
                depreciable_base = schedule.closing_gross
                remaining_life = max(1, weighted_life)
                schedule.depreciation = depreciable_base / remaining_life
            
            schedule.closing_accum_depr = (
                schedule.opening_accum_depr + schedule.depreciation
            )
            schedule.net_book_value = schedule.closing_gross - schedule.closing_accum_depr
            
            schedules.append(schedule)
            prev_gross = schedule.closing_gross
            prev_accum_depr = schedule.closing_accum_depr
        
        return schedules
    
    def _build_deferred_income_schedule(
        self,
        grant_amount: float,
    ) -> list[DeferredIncomeSchedule]:
        """Build deferred income schedule for Policy A2."""
        schedules = []
        grant = self.incentives.capital_grant
        
        if not grant.enabled or grant.accounting_policy != GrantAccountingPolicy.DEFERRED_INCOME:
            # Return empty schedules with zero values
            for year in [self.horizon.base_year] + self.horizon.all_forecast_years:
                schedules.append(DeferredIncomeSchedule(year=year))
            return schedules
        
        all_years = [self.horizon.base_year] + self.horizon.all_forecast_years
        recognition_year = self._get_grant_recognition_year()
        
        # Determine release period
        release_years = grant.release_years
        if release_years is None:
            # Default to weighted average asset life
            total_capex = self.capex.total_capex()
            if total_capex > 0:
                release_years = int(sum(
                    item.amount * item.useful_life_years 
                    for item in self.capex.all_items().values()
                ) / total_capex)
            else:
                release_years = 20
        
        annual_release = grant_amount / max(1, release_years)
        
        prev_balance = 0.0
        grant_received = 0.0
        
        for i, year in enumerate(all_years):
            schedule = DeferredIncomeSchedule(year=year)
            schedule.opening_balance = prev_balance
            
            # Grant received (at recognition)
            if year == recognition_year:
                schedule.grant_received = grant_amount
            
            # Release to P&L (starting from COD)
            if year >= self.horizon.cod_year and prev_balance + schedule.grant_received > 0:
                schedule.release_to_pl = min(
                    annual_release, 
                    schedule.opening_balance + schedule.grant_received
                )
            
            schedule.closing_balance = (
                schedule.opening_balance + 
                schedule.grant_received - 
                schedule.release_to_pl
            )
            
            schedules.append(schedule)
            prev_balance = schedule.closing_balance
        
        return schedules
    
    def _build_grant_receivable_schedule(
        self,
        grant_amount: float,
    ) -> list[GrantReceivableSchedule]:
        """Build grant receivable from PA schedule (OIC 15)."""
        schedules = []
        grant = self.incentives.capital_grant
        
        if not grant.enabled:
            for year in [self.horizon.base_year] + self.horizon.all_forecast_years:
                schedules.append(GrantReceivableSchedule(year=year))
            return schedules
        
        all_years = [self.horizon.base_year] + self.horizon.all_forecast_years
        recognition_year = self._get_grant_recognition_year()
        
        # Build cash receipt schedule
        cash_profile = grant.cash_receipt_schedule
        cod_year = self.horizon.cod_year
        
        cash_by_year = {}
        for i, pct in enumerate(cash_profile):
            receipt_year = cod_year + i
            if receipt_year in all_years or receipt_year <= all_years[-1]:
                cash_by_year[receipt_year] = grant_amount * pct
        
        prev_balance = 0.0
        
        for year in all_years:
            schedule = GrantReceivableSchedule(year=year)
            schedule.opening_balance = prev_balance
            
            # Recognition
            if year == recognition_year:
                schedule.grant_recognized = grant_amount
            
            # Cash receipt
            schedule.cash_received = cash_by_year.get(year, 0.0)
            
            schedule.closing_balance = (
                schedule.opening_balance + 
                schedule.grant_recognized - 
                schedule.cash_received
            )
            
            schedules.append(schedule)
            prev_balance = schedule.closing_balance
        
        return schedules
    
    def _build_tax_credit_schedule(
        self,
        credit_amount: float,
        taxes_before_credit: dict[int, float],
    ) -> list[TaxCreditSchedule]:
        """Build tax credit utilization schedule (OIC 25)."""
        schedules = []
        tax_credit = self.incentives.tax_credit
        
        all_years = [self.horizon.base_year] + self.horizon.all_forecast_years
        
        if not tax_credit.enabled:
            for year in all_years:
                schedules.append(TaxCreditSchedule(year=year))
            return schedules
        
        # Determine when credit becomes available
        usable_from = tax_credit.usable_from_year or self.horizon.cod_year
        carry_forward = tax_credit.carry_forward_years
        cap_pct = tax_credit.annual_cap_percent
        
        prev_balance = 0.0
        credit_added = False
        credit_year_added = None
        
        for year in all_years:
            schedule = TaxCreditSchedule(year=year)
            schedule.opening_balance = prev_balance
            
            # Credit becomes available
            if year == usable_from and not credit_added:
                schedule.credit_available = credit_amount
                credit_added = True
                credit_year_added = year
            
            # Utilization (capped by tax liability and annual cap)
            if year >= usable_from:
                available = schedule.opening_balance + schedule.credit_available
                tax_liability = max(0, taxes_before_credit.get(year, 0.0))
                max_usage = tax_liability * cap_pct
                schedule.utilization = min(available, max_usage)
            
            # Expiry (after carry-forward period)
            if credit_year_added and year >= credit_year_added + carry_forward:
                # Expire any remaining balance
                remaining = (
                    schedule.opening_balance + 
                    schedule.credit_available - 
                    schedule.utilization
                )
                schedule.expiry = remaining
            
            schedule.closing_balance = (
                schedule.opening_balance + 
                schedule.credit_available - 
                schedule.utilization - 
                schedule.expiry
            )
            
            schedules.append(schedule)
            prev_balance = schedule.closing_balance
        
        return schedules
    
    def _get_grant_recognition_year(self) -> int:
        """Determine year when grant is recognized."""
        grant = self.incentives.capital_grant
        trigger = grant.recognition_trigger
        
        if trigger == GrantRecognitionTrigger.AT_COD:
            return self.horizon.cod_year
        elif trigger == GrantRecognitionTrigger.AT_DECREE:
            # Assume decree is granted at end of construction
            return self.horizon.cod_year - 1 if self.horizon.construction_years > 0 else self.horizon.cod_year
        else:  # AT_CASH_RECEIPT
            # First cash receipt year
            return self.horizon.cod_year


def compute_accounting(
    case: BiometanoCase,
    capex_by_year: dict[int, float],
    taxes_before_credit: dict[int, float],
) -> AccountingOutputs:
    """
    Convenience function to compute accounting schedules.
    
    Args:
        case: Biometano case inputs
        capex_by_year: CAPEX spend by year
        taxes_before_credit: Tax liability before credits by year
        
    Returns:
        AccountingOutputs with all schedules
    """
    calculator = AccountingCalculator(case)
    return calculator.compute(capex_by_year, taxes_before_credit)
