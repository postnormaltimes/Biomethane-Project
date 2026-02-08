"""
Biometano Builder Module

Converts BiometanoCase inputs into engine-ready projections and schedules.
Handles construction vs operating phases, ramp-up, escalation, and payment delays.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from dcf_projects.biometano.schema import (
    BiometanoCase,
    GrantAccountingPolicy,
)
from dcf_projects.biometano.incentives_allocation import (
    IncentiveAllocationResult,
    PnrrParams,
    ZesParams,
    CapexLineItem,
    compute_full_incentive_allocation,
)
from dcf_projects.biometano.accounting import (
    AccountingOutputs,
    compute_accounting,
)


@dataclass
class YearlyRevenue:
    """Revenue breakdown for a single year."""
    year: int
    gate_fee: float = 0.0
    tariff: float = 0.0
    co2: float = 0.0
    go: float = 0.0
    compost: float = 0.0
    
    @property
    def total(self) -> float:
        return self.gate_fee + self.tariff + self.co2 + self.go + self.compost


@dataclass
class YearlyOpex:
    """OPEX breakdown for a single year."""
    year: int
    feedstock_handling: float = 0.0
    utilities: float = 0.0
    chemicals: float = 0.0
    maintenance: float = 0.0
    personnel: float = 0.0
    insurance: float = 0.0
    overheads: float = 0.0
    digestate_handling: float = 0.0
    other: float = 0.0
    
    @property
    def total(self) -> float:
        return (
            self.feedstock_handling + self.utilities + self.chemicals +
            self.maintenance + self.personnel + self.insurance +
            self.overheads + self.digestate_handling + self.other
        )


@dataclass
class YearlyProduction:
    """Production volumes for a single year."""
    year: int
    forsu_tonnes: float = 0.0
    biomethane_mwh: float = 0.0
    co2_tonnes: float = 0.0
    compost_tonnes: float = 0.0
    availability: float = 0.0


@dataclass
class YearlyCapex:
    """CAPEX breakdown for a single year."""
    year: int
    epc: float = 0.0
    civils: float = 0.0
    upgrading_unit: float = 0.0
    grid_connection: float = 0.0
    engineering: float = 0.0
    permitting: float = 0.0
    contingency: float = 0.0
    startup_costs: float = 0.0
    other: float = 0.0
    
    @property
    def total(self) -> float:
        return (
            self.epc + self.civils + self.upgrading_unit + self.grid_connection +
            self.engineering + self.permitting + self.contingency +
            self.startup_costs + self.other
        )
    
    @property
    def capitalizable(self) -> float:
        """Total capitalizable CAPEX (excludes startup costs if not capitalized)."""
        return self.total  # Simplified; could exclude startup_costs


@dataclass
class YearlyFinancing:
    """Financing details for a single year."""
    year: int
    debt_balance_opening: float = 0.0
    debt_drawdown: float = 0.0
    debt_repayment: float = 0.0
    debt_balance_closing: float = 0.0
    interest_expense: float = 0.0
    cost_of_debt: float = 0.0


@dataclass
class BiometanoProjections:
    """Complete projections from BiometanoCase."""
    base_year: int
    cod_year: int
    construction_years: list[int]
    operating_years: list[int]
    all_forecast_years: list[int]
    
    production: list[YearlyProduction] = field(default_factory=list)
    revenues: list[YearlyRevenue] = field(default_factory=list)
    opex: list[YearlyOpex] = field(default_factory=list)
    capex: list[YearlyCapex] = field(default_factory=list)
    financing: list[YearlyFinancing] = field(default_factory=list)
    accounting: Optional[AccountingOutputs] = None
    incentive_allocation: Optional[IncentiveAllocationResult] = None
    
    # Derived schedules
    ebitda: dict[int, float] = field(default_factory=dict)
    depreciation: dict[int, float] = field(default_factory=dict)
    grant_income_release: dict[int, float] = field(default_factory=dict)
    ebit: dict[int, float] = field(default_factory=dict)
    interest: dict[int, float] = field(default_factory=dict)
    ebt: dict[int, float] = field(default_factory=dict)
    taxes_before_credit: dict[int, float] = field(default_factory=dict)
    tax_credit_utilization: dict[int, float] = field(default_factory=dict)
    taxes_paid: dict[int, float] = field(default_factory=dict)
    net_income: dict[int, float] = field(default_factory=dict)
    
    # Cash flow components
    nwc: dict[int, float] = field(default_factory=dict)
    delta_nwc: dict[int, float] = field(default_factory=dict)
    fcff: dict[int, float] = field(default_factory=dict)
    fcfe: dict[int, float] = field(default_factory=dict)
    
    # Receivables/Payables for NWC
    ar_trade: dict[int, float] = field(default_factory=dict)
    ar_grant: dict[int, float] = field(default_factory=dict)
    ar_tax_credit: dict[int, float] = field(default_factory=dict)
    ap_trade: dict[int, float] = field(default_factory=dict)
    
    def get_revenue(self, year: int) -> Optional[YearlyRevenue]:
        for r in self.revenues:
            if r.year == year:
                return r
        return None
    
    def get_opex(self, year: int) -> Optional[YearlyOpex]:
        for o in self.opex:
            if o.year == year:
                return o
        return None
    
    def get_capex(self, year: int) -> Optional[YearlyCapex]:
        for c in self.capex:
            if c.year == year:
                return c
        return None
    
    def get_financing(self, year: int) -> Optional[YearlyFinancing]:
        for f in self.financing:
            if f.year == year:
                return f
        return None


class BiometanoBuilder:
    """
    Builds projections from BiometanoCase inputs.
    
    Orchestrates all calculation steps:
    1. Timeline expansion
    2. Production volumes with availability ramp-up
    3. Revenue by channel with escalation
    4. OPEX by category with ramp-up and escalation
    5. CAPEX by item with spend profile
    6. Financing schedule
    7. Accounting (grant/tax credit)
    8. Income statement items (EBITDA, D&A, EBIT, taxes, NI)
    9. Cash flow items (NWC, FCFF, FCFE)
    """
    
    def __init__(self, case: BiometanoCase):
        self.case = case
        self.horizon = case.horizon
        
    def build(self) -> BiometanoProjections:
        """Build complete projections from case inputs."""
        projections = BiometanoProjections(
            base_year=self.horizon.base_year,
            cod_year=self.horizon.cod_year,
            construction_years=self.horizon.construction_years_list,
            operating_years=self.horizon.operating_years_list,
            all_forecast_years=self.horizon.all_forecast_years,
        )
        
        # Step 1: Production volumes
        projections.production = self._build_production()
        
        # Step 2: CAPEX schedule
        projections.capex = self._build_capex()
        capex_by_year = {c.year: c.total for c in projections.capex}
        
        # Step 3: Financing schedule
        projections.financing = self._build_financing(capex_by_year)
        
        # Step 4: Revenue schedule
        projections.revenues = self._build_revenues(projections.production)
        
        # Step 5: OPEX schedule
        total_capex = self.case.capex.total_capex()
        projections.opex = self._build_opex(projections.production, total_capex)
        
        # Step 6: EBITDA
        self._compute_ebitda(projections)
        
        # Step 7: D&A (preliminary, before accounting adjustments)
        self._compute_depreciation_preliminary(projections, capex_by_year)
        
        # Step 8: EBIT, Interest, EBT, Taxes (preliminary for accounting)
        self._compute_ebit_preliminary(projections)
        
        # Step 9: Accounting schedules (needs preliminary taxes)
        projections.accounting = compute_accounting(
            self.case, capex_by_year, projections.taxes_before_credit
        )
        
        # Step 9b: Incentives allocation (PNRR + ZES waterfall)
        if self.case.incentives_policy.enabled:
            projections.incentive_allocation = self._compute_incentive_allocation()
        
        # Step 10: Final D&A with grant adjustments
        self._compute_depreciation_final(projections)
        
        # Step 11: Final income statement
        self._compute_income_statement_final(projections)
        
        # Step 12: NWC and cash flows
        self._compute_nwc(projections)
        self._compute_cash_flows(projections)
        
        return projections
    
    def _build_production(self) -> list[YearlyProduction]:
        """Build production schedule with availability ramp-up."""
        prod_inputs = self.case.production
        production = []
        
        for year in self.horizon.all_forecast_years:
            p = YearlyProduction(year=year)
            
            if year in self.horizon.operating_years_list:
                # Operating year - apply availability
                op_year_idx = year - self.horizon.cod_year
                availability = prod_inputs.get_availability(op_year_idx)
                p.availability = availability
                
                p.forsu_tonnes = prod_inputs.forsu_throughput_tpy * availability
                p.biomethane_mwh = prod_inputs.get_biomethane_mwh() * availability
                p.co2_tonnes = prod_inputs.co2_tpy * availability
                p.compost_tonnes = prod_inputs.compost_tpy * availability
            else:
                # Construction year - no production
                p.availability = 0.0
            
            production.append(p)
        
        return production
    
    def _build_capex(self) -> list[YearlyCapex]:
        """Build CAPEX schedule from spend profiles."""
        capex_inputs = self.case.capex
        capex = []
        
        construction_years = self.horizon.construction_years_list
        if not construction_years:
            # If no construction years, all CAPEX in base year
            construction_years = [self.horizon.base_year]
        
        for year in self.horizon.all_forecast_years:
            c = YearlyCapex(year=year)
            
            for name, item in capex_inputs.all_items().items():
                if item.amount > 0:
                    # Find spend for this year based on profile
                    if year in construction_years:
                        idx = construction_years.index(year)
                        if idx < len(item.spend_profile):
                            spend = item.amount * item.spend_profile[idx]
                            setattr(c, name, spend)
            
            capex.append(c)
        
        return capex
    
    def _build_financing(self, capex_by_year: dict[int, float]) -> list[YearlyFinancing]:
        """Build financing schedule with debt drawdown and repayment."""
        fin_inputs = self.case.financing
        financing = []
        
        construction_years = self.horizon.construction_years_list
        if not construction_years:
            construction_years = [self.horizon.base_year]
        
        # Debt drawdown during construction
        drawdown_by_year = {}
        total_debt = fin_inputs.debt_amount
        for i, year in enumerate(construction_years):
            if i < len(fin_inputs.debt_drawdown_profile):
                drawdown_by_year[year] = total_debt * fin_inputs.debt_drawdown_profile[i]
        
        # Debt repayment during operation
        repayment_years = fin_inputs.debt_repayment_years
        if repayment_years > 0 and total_debt > 0:
            annual_repayment = total_debt / repayment_years
        else:
            annual_repayment = 0
        
        prev_debt = 0.0
        
        for year in self.horizon.all_forecast_years:
            f = YearlyFinancing(year=year)
            f.debt_balance_opening = prev_debt
            f.cost_of_debt = fin_inputs.get_rd(year)
            
            # Drawdown
            f.debt_drawdown = drawdown_by_year.get(year, 0.0)
            
            # Repayment (only during operation)
            if year >= self.horizon.cod_year and prev_debt > 0:
                f.debt_repayment = min(annual_repayment, f.debt_balance_opening + f.debt_drawdown)
            
            f.debt_balance_closing = (
                f.debt_balance_opening + f.debt_drawdown - f.debt_repayment
            )
            
            # Interest on closing balance (end-of-period convention)
            f.interest_expense = f.debt_balance_closing * f.cost_of_debt
            
            financing.append(f)
            prev_debt = f.debt_balance_closing
        
        return financing
    
    def _build_revenues(self, production: list[YearlyProduction]) -> list[YearlyRevenue]:
        """Build revenue schedule with escalation."""
        rev_inputs = self.case.revenues
        revenues = []
        
        cod_year = self.horizon.cod_year
        
        for prod in production:
            year = prod.year
            r = YearlyRevenue(year=year)
            
            if year >= cod_year:
                years_from_cod = year - cod_year
                
                # Gate fee (€/tonne FORSU)
                if rev_inputs.gate_fee.enabled:
                    price = rev_inputs.gate_fee.price * (1 + rev_inputs.gate_fee.escalation_rate) ** years_from_cod
                    r.gate_fee = prod.forsu_tonnes * price
                
                # Tariff (€/MWh biomethane)
                if rev_inputs.tariff.enabled:
                    price = rev_inputs.tariff.price * (1 + rev_inputs.tariff.escalation_rate) ** years_from_cod
                    r.tariff = prod.biomethane_mwh * price
                
                # CO2 (€/tonne)
                if rev_inputs.co2.enabled:
                    price = rev_inputs.co2.price * (1 + rev_inputs.co2.escalation_rate) ** years_from_cod
                    r.co2 = prod.co2_tonnes * price
                
                # GO (€/MWh biomethane)
                if rev_inputs.go.enabled:
                    price = rev_inputs.go.price * (1 + rev_inputs.go.escalation_rate) ** years_from_cod
                    r.go = prod.biomethane_mwh * price
                
                # Compost (€/tonne)
                if rev_inputs.compost.enabled:
                    price = rev_inputs.compost.price * (1 + rev_inputs.compost.escalation_rate) ** years_from_cod
                    r.compost = prod.compost_tonnes * price
            
            revenues.append(r)
        
        return revenues
    
    def _build_opex(
        self, 
        production: list[YearlyProduction],
        total_capex: float,
    ) -> list[YearlyOpex]:
        """Build OPEX schedule with ramp-up and escalation."""
        opex_inputs = self.case.opex
        opex = []
        
        cod_year = self.horizon.cod_year
        prod_by_year = {p.year: p for p in production}
        
        for year in self.horizon.all_forecast_years:
            o = YearlyOpex(year=year)
            
            if year >= cod_year:
                years_from_cod = year - cod_year
                prod = prod_by_year.get(year)
                availability = prod.availability if prod else 0.95
                
                for cat_name, cat_inputs in opex_inputs.all_categories().items():
                    # Determine ramp-up factor
                    if cat_inputs.ramp_up_profile:
                        if years_from_cod < len(cat_inputs.ramp_up_profile):
                            ramp_factor = cat_inputs.ramp_up_profile[years_from_cod]
                        else:
                            ramp_factor = cat_inputs.ramp_up_profile[-1]
                    else:
                        ramp_factor = availability
                    
                    # Apply escalation
                    escalation = (1 + cat_inputs.escalation_rate) ** years_from_cod
                    
                    # Compute cost
                    cost = cat_inputs.fixed_annual * escalation * ramp_factor
                    
                    if prod:
                        cost += cat_inputs.variable_per_tonne * prod.forsu_tonnes * escalation
                        cost += cat_inputs.variable_per_mwh * prod.biomethane_mwh * escalation
                    
                    cost += total_capex * cat_inputs.percent_of_capex * escalation
                    
                    setattr(o, cat_name, cost)
            
            opex.append(o)
        
        return opex
    
    def _compute_ebitda(self, projections: BiometanoProjections) -> None:
        """Compute EBITDA = Revenue - OPEX."""
        for year in projections.all_forecast_years:
            rev = projections.get_revenue(year)
            opex = projections.get_opex(year)
            
            revenue = rev.total if rev else 0.0
            costs = opex.total if opex else 0.0
            
            projections.ebitda[year] = revenue - costs
    
    def _compute_depreciation_preliminary(
        self, 
        projections: BiometanoProjections,
        capex_by_year: dict[int, float],
    ) -> None:
        """Compute preliminary D&A for tax calculation."""
        # Simple straight-line depreciation
        total_capex = sum(capex_by_year.values())
        weighted_life = 20.0  # Default
        
        capex_inputs = self.case.capex
        if capex_inputs.total_capex() > 0:
            weighted_life = sum(
                item.amount * item.useful_life_years 
                for item in capex_inputs.all_items().values()
            ) / capex_inputs.total_capex()
        
        annual_depr = total_capex / max(1, weighted_life) if total_capex > 0 else 0
        
        for year in projections.all_forecast_years:
            if year >= projections.cod_year:
                projections.depreciation[year] = annual_depr
            else:
                projections.depreciation[year] = 0.0
    
    def _compute_ebit_preliminary(self, projections: BiometanoProjections) -> None:
        """Compute preliminary income statement for accounting calculations."""
        tax_rate = self.case.financing.tax_rate
        
        for year in projections.all_forecast_years:
            ebitda = projections.ebitda.get(year, 0.0)
            depreciation = projections.depreciation.get(year, 0.0)
            
            ebit = ebitda - depreciation
            projections.ebit[year] = ebit
            
            # Interest
            fin = projections.get_financing(year)
            interest = fin.interest_expense if fin else 0.0
            projections.interest[year] = interest
            
            # EBT
            ebt = ebit - interest
            projections.ebt[year] = ebt
            
            # Taxes before credits
            taxes = max(0, ebt * tax_rate)
            projections.taxes_before_credit[year] = taxes
    
    def _compute_depreciation_final(self, projections: BiometanoProjections) -> None:
        """Compute final D&A from accounting schedules (grant-adjusted)."""
        if projections.accounting:
            for fa in projections.accounting.fixed_assets:
                if fa.year in projections.all_forecast_years:
                    projections.depreciation[fa.year] = fa.depreciation
    
    def _compute_income_statement_final(self, projections: BiometanoProjections) -> None:
        """Compute final income statement with accounting adjustments."""
        tax_rate = self.case.financing.tax_rate
        grant = self.case.incentives.capital_grant
        
        for year in projections.all_forecast_years:
            ebitda = projections.ebitda.get(year, 0.0)
            depreciation = projections.depreciation.get(year, 0.0)
            
            # Grant income release (Policy A2)
            grant_release = 0.0
            if projections.accounting and grant.accounting_policy == GrantAccountingPolicy.DEFERRED_INCOME:
                di = projections.accounting.get_deferred_income(year)
                if di:
                    grant_release = di.release_to_pl
            projections.grant_income_release[year] = grant_release
            
            # EBIT
            ebit = ebitda - depreciation + grant_release
            projections.ebit[year] = ebit
            
            # Interest
            fin = projections.get_financing(year)
            interest = fin.interest_expense if fin else 0.0
            projections.interest[year] = interest
            
            # EBT
            ebt = ebit - interest
            projections.ebt[year] = ebt
            
            # Taxes (with credit utilization)
            taxes_before = max(0, ebt * tax_rate)
            projections.taxes_before_credit[year] = taxes_before
            
            credit_used = 0.0
            if projections.accounting:
                tc = projections.accounting.get_tax_credit(year)
                if tc:
                    credit_used = tc.utilization
            projections.tax_credit_utilization[year] = credit_used
            projections.taxes_paid[year] = taxes_before - credit_used
            
            # Net income
            projections.net_income[year] = ebt - projections.taxes_paid[year]
    
    def _compute_nwc(self, projections: BiometanoProjections) -> None:
        """Compute NWC from payment delays."""
        rev_inputs = self.case.revenues
        opex_inputs = self.case.opex
        
        for year in projections.all_forecast_years:
            rev = projections.get_revenue(year)
            opex = projections.get_opex(year)
            
            # AR from trade (based on payment delays)
            ar = 0.0
            if rev:
                # Gate fee AR
                delay_days = rev_inputs.gate_fee.payment_delay_days
                ar += rev.gate_fee * (delay_days / 365)
                
                # Tariff AR
                delay_days = rev_inputs.tariff.payment_delay_days
                ar += rev.tariff * (delay_days / 365)
                
                # CO2 AR
                delay_days = rev_inputs.co2.payment_delay_days
                ar += rev.co2 * (delay_days / 365)
                
                # GO AR
                delay_days = rev_inputs.go.payment_delay_days
                ar += rev.go * (delay_days / 365)
                
                # Compost AR
                delay_days = rev_inputs.compost.payment_delay_days
                ar += rev.compost * (delay_days / 365)
            
            projections.ar_trade[year] = ar
            
            # AP from trade (based on payment delays)
            ap = 0.0
            if opex:
                for cat_name, cat_inputs in opex_inputs.all_categories().items():
                    cat_cost = getattr(opex, cat_name, 0.0)
                    delay_days = cat_inputs.payment_delay_days
                    ap += cat_cost * (delay_days / 365)
            
            projections.ap_trade[year] = ap
            
            # Grant receivable
            ar_grant = 0.0
            if projections.accounting:
                gr = projections.accounting.get_grant_receivable(year)
                if gr:
                    ar_grant = gr.closing_balance
            projections.ar_grant[year] = ar_grant
            
            # Tax credit receivable (if policy B2)
            ar_tax = 0.0
            if projections.accounting:
                tc = projections.accounting.get_tax_credit(year)
                if tc:
                    ar_tax = tc.closing_balance
            projections.ar_tax_credit[year] = ar_tax
            
            # NWC = AR - AP (excluding grant and tax credit receivables for operating NWC)
            projections.nwc[year] = ar - ap
        
        # Delta NWC
        all_years = [projections.base_year] + projections.all_forecast_years
        projections.nwc[projections.base_year] = 0.0  # Assume zero at base
        
        for i, year in enumerate(projections.all_forecast_years):
            prev_year = all_years[i]  # Previous is either base_year or forecast year
            prev_nwc = projections.nwc.get(prev_year, 0.0)
            curr_nwc = projections.nwc.get(year, 0.0)
            projections.delta_nwc[year] = curr_nwc - prev_nwc
    
    def _compute_cash_flows(self, projections: BiometanoProjections) -> None:
        """Compute FCFF and FCFE."""
        tax_rate = self.case.financing.tax_rate
        
        for year in projections.all_forecast_years:
            ebit = projections.ebit.get(year, 0.0)
            depreciation = projections.depreciation.get(year, 0.0)
            delta_nwc = projections.delta_nwc.get(year, 0.0)
            
            capex = projections.get_capex(year)
            capex_total = capex.total if capex else 0.0
            
            interest = projections.interest.get(year, 0.0)
            tax_credit = projections.tax_credit_utilization.get(year, 0.0)
            
            # NOPAT = EBIT * (1 - tax_rate)
            nopat = ebit * (1 - tax_rate)
            
            # FCFF = NOPAT + D&A - ΔNWC - CAPEX
            fcff = nopat + depreciation - delta_nwc - capex_total
            projections.fcff[year] = fcff
            
            # Net borrowing
            fin = projections.get_financing(year)
            net_borrowing = 0.0
            if fin:
                net_borrowing = fin.debt_drawdown - fin.debt_repayment
            
            # FCFE = FCFF - Interest*(1-t) + Net Borrowing
            fcfe = fcff - interest * (1 - tax_rate) + net_borrowing
            projections.fcfe[year] = fcfe
    
    def _compute_incentive_allocation(self) -> IncentiveAllocationResult:
        """Compute incentives allocation from policy config."""
        policy = self.case.incentives_policy
        
        # Build PnrrParams from schema
        pnrr_params = PnrrParams(
            grant_rate=policy.pnrr.grant_rate,
            annual_hours=policy.pnrr.annual_hours,
            biomethane_smc_per_year=policy.pnrr.biomethane_smc_per_year,
            cs_max_base_eur_per_smcph=policy.pnrr.cs_max_base_eur_per_smcph,
            inflation_adj_factor=policy.pnrr.inflation_adj_factor,
            tech_costs_cap_pct=policy.pnrr.tech_costs_cap_pct,
        )
        
        # Build ZesParams from schema (use defaults if not specified)
        zes_params = ZesParams()
        if policy.zes:
            zes_params = ZesParams(
                theoretical_rate=policy.zes.theoretical_rate,
                riparto_coeff=policy.zes.riparto_coeff,
                allow_overlap_with_pnrr=policy.zes.allow_overlap_with_pnrr,
                enforce_no_double_financing=policy.zes.enforce_no_double_financing,
                base_allocation_strategy=policy.zes.base_allocation_strategy,
                tech_costs_eligibility_share=policy.zes.tech_costs_eligibility_share,
            )
        
        # Build CapexLineItem list from schema
        capex_lines = []
        for item in policy.capex_breakdown:
            capex_lines.append(CapexLineItem(
                name=item.name,
                amount=item.amount,
                oic_class=item.oic_class,
                pnrr_eligible=item.pnrr_eligible,
                zes_eligible=item.zes_eligible,
                notes=item.notes,
            ))
        
        # Compute and return allocation result
        return compute_full_incentive_allocation(
            capex_lines=capex_lines,
            pnrr=pnrr_params,
            zes=zes_params,
            max_esl_intensity=policy.max_esl_intensity,
        )


def build_projections(case: BiometanoCase) -> BiometanoProjections:
    """
    Convenience function to build projections from case.
    
    Args:
        case: Biometano case inputs
        
    Returns:
        Complete BiometanoProjections
    """
    builder = BiometanoBuilder(case)
    return builder.build()
