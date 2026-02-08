"""
DCF Engine Main Orchestrator

Coordinates all calculation modules to produce complete DCF outputs.
"""
from __future__ import annotations

from dcf_engine.models import (
    DCFInputs,
    DCFOutputs,
    DiscountingMode,
    TerminalValueMethod,
    WeightingMode,
    YearlyProjection,
    YearlyNWC,
    YearlyCashFlow,
    YearlyDiscount,
    TerminalValueOutput,
    ValuationBridge,
    WACCDetails,
)
from dcf_engine.validation import validate_inputs
from dcf_engine.projections import (
    compute_revenue,
    compute_operating_costs,
    compute_ebitda,
    compute_ebit,
    compute_nwc,
    compute_delta_nwc,
)
from dcf_engine.taxes import (
    compute_tax_on_ebit,
    compute_nopat,
    compute_ebt,
    compute_taxes_on_ebt,
    compute_net_income,
)
from dcf_engine.cashflows import (
    compute_interest_expense,
    compute_net_borrowing,
    compute_fcff,
    compute_fcfe_from_fcff,
    compute_fcfe_from_net_income,
)
from dcf_engine.discount_rates import (
    compute_ke,
    compute_equity_book_rollforward,
    compute_wacc,
    get_wacc_weights,
)
from dcf_engine.discounting import (
    compute_discount_factors,
    compute_pv_series,
    compute_pv_single,
    sum_pv,
)
from dcf_engine.terminal_value import compute_terminal_value
from dcf_engine.valuation import (
    compute_enterprise_value,
    compute_net_debt,
    compute_equity_from_ev,
    compute_equity_direct,
    reconcile_valuations,
)


class DCFEngine:
    """
    Main DCF computation engine.
    
    Orchestrates all calculation modules to produce auditable outputs.
    """
    
    def __init__(self, inputs: DCFInputs):
        """
        Initialize engine with inputs.
        
        Args:
            inputs: Complete DCF model inputs
        """
        self.inputs = inputs
        self._validate()
    
    def _validate(self) -> None:
        """Validate inputs before computation."""
        validate_inputs(self.inputs)
    
    def run(self) -> DCFOutputs:
        """
        Execute full DCF computation.
        
        Returns:
            DCFOutputs with all projections, cash flows, and valuations
        """
        inputs = self.inputs
        base_year = inputs.timeline.base_year
        forecast_years = inputs.timeline.forecast_years
        final_year = forecast_years[-1]
        n_periods = len(forecast_years)
        
        # ====================================================================
        # STEP 1: Operating Projections
        # ====================================================================
        
        revenue = compute_revenue(inputs)
        operating_costs = compute_operating_costs(inputs, revenue)
        ebitda = compute_ebitda(inputs, revenue, operating_costs)
        ebit = compute_ebit(inputs, ebitda)
        da = inputs.operating.depreciation_amortization
        
        # NWC schedule
        nwc = compute_nwc(inputs, revenue)
        delta_nwc = compute_delta_nwc(inputs, nwc)
        
        # ====================================================================
        # STEP 2: Tax Calculations
        # ====================================================================
        
        # Mode A: Tax on EBIT / NOPAT
        tax_on_ebit = compute_tax_on_ebit(inputs, ebit)
        nopat = compute_nopat(inputs, ebit)
        
        # Interest expense for Mode B
        interest_expense = compute_interest_expense(inputs)
        
        # Mode B: EBT / Net Income
        ebt = compute_ebt(ebit, interest_expense)
        taxes_on_ebt = compute_taxes_on_ebt(inputs, ebt)
        net_income = compute_net_income(ebt, taxes_on_ebt)
        
        # ====================================================================
        # STEP 3: Cash Flow Construction
        # ====================================================================
        
        capex = inputs.investments.capex
        net_borrowing = compute_net_borrowing(inputs)
        
        # FCFF
        fcff = compute_fcff(nopat, da, delta_nwc, capex)
        
        # FCFE (both derivations)
        fcfe_from_fcff = compute_fcfe_from_fcff(
            inputs, fcff, interest_expense, net_borrowing
        )
        fcfe_from_ni = compute_fcfe_from_net_income(
            net_income, da, delta_nwc, capex, net_borrowing
        )
        
        # Use FCFE from FCFF as primary
        fcfe = fcfe_from_fcff
        
        # ====================================================================
        # STEP 4: Discount Rates
        # ====================================================================
        
        ke = compute_ke(inputs)
        
        # Equity book roll-forward (for book-value WACC weights)
        equity_book = None
        if inputs.wacc.weighting_mode == WeightingMode.BOOK_VALUE:
            equity_book = compute_equity_book_rollforward(inputs, net_income)
        
        wacc = compute_wacc(inputs, ke, equity_book)
        wacc_weights = get_wacc_weights(inputs, equity_book)
        
        # ====================================================================
        # STEP 5: Discounting
        # ====================================================================
        
        # Discount factors for WACC
        df_wacc = compute_discount_factors(
            wacc, forecast_years, base_year, inputs.discounting_mode
        )
        
        # Discount factors for Ke (constant rate)
        ke_rates = {year: ke for year in forecast_years}
        df_ke = compute_discount_factors(
            ke_rates, forecast_years, base_year, inputs.discounting_mode
        )
        
        # Present values
        pv_fcff = compute_pv_series(fcff, df_wacc)
        pv_fcfe = compute_pv_series(fcfe, df_ke)
        
        sum_pv_fcff = sum_pv(pv_fcff)
        sum_pv_fcfe = sum_pv(pv_fcfe)
        
        # ====================================================================
        # STEP 6: Terminal Value
        # ====================================================================
        
        tv_fcff, tv_fcfe = compute_terminal_value(
            inputs, fcff, fcfe, wacc, ke, ebitda, ebit, revenue
        )
        
        # PV of terminal value
        final_wacc = wacc[final_year]
        pv_tv_fcff = compute_pv_single(tv_fcff, final_wacc, n_periods)
        pv_tv_fcfe = compute_pv_single(tv_fcfe, ke, n_periods)
        
        # ====================================================================
        # STEP 7: Valuation Bridges
        # ====================================================================
        
        # Enterprise Value (FCFF approach)
        ev = compute_enterprise_value(sum_pv_fcff, pv_tv_fcff)
        
        # Net Debt at base year
        base_debt = inputs.debt.debt_balances[base_year]
        cash = inputs.net_debt.cash_and_equivalents
        net_debt_value = compute_net_debt(base_debt, cash)
        
        # Equity from EV
        equity_from_ev_value = compute_equity_from_ev(ev, net_debt_value)
        
        # Equity direct (FCFE approach)
        equity_direct_value = compute_equity_direct(sum_pv_fcfe, pv_tv_fcfe)
        
        # Reconciliation
        recon_diff, recon_notes = reconcile_valuations(
            equity_from_ev_value, equity_direct_value
        )
        
        # ====================================================================
        # STEP 8: Build Output Models
        # ====================================================================
        
        # Projections
        projections = []
        for year in forecast_years:
            projections.append(YearlyProjection(
                year=year,
                revenue=revenue[year],
                operating_costs=operating_costs[year],
                ebitda=ebitda[year],
                depreciation_amortization=da[year],
                ebit=ebit[year],
                tax_on_ebit=tax_on_ebit[year],
                nopat=nopat[year],
                interest_expense=interest_expense[year],
                ebt=ebt[year],
                taxes_on_ebt=taxes_on_ebt[year],
                net_income=net_income[year],
            ))
        
        # NWC schedule (include base year)
        nwc_schedule = []
        all_years = [base_year] + forecast_years
        for i, year in enumerate(all_years):
            d_nwc = delta_nwc.get(year, 0.0)
            nwc_schedule.append(YearlyNWC(
                year=year,
                nwc=nwc[year],
                delta_nwc=d_nwc,
            ))
        
        # Cash flows
        cash_flows = []
        for year in forecast_years:
            tax_rate = inputs.tax.get_rate(year)
            cash_flows.append(YearlyCashFlow(
                year=year,
                nopat=nopat[year],
                depreciation_amortization=da[year],
                delta_nwc=delta_nwc[year],
                capex=capex[year],
                fcff=fcff[year],
                interest_expense=interest_expense[year],
                interest_tax_shield=interest_expense[year] * tax_rate,
                net_borrowing=net_borrowing[year],
                fcfe=fcfe[year],
            ))
        
        # WACC details
        wacc_details = []
        for year in forecast_years:
            w_e, w_d = wacc_weights[year]
            debt_balance = inputs.debt.debt_balances[year]
            eq_book = equity_book.get(year, 0.0) if equity_book else 0.0
            wacc_details.append(WACCDetails(
                year=year,
                ke=ke,
                rd=inputs.debt.get_rd(year),
                tax_rate=inputs.tax.get_rate(year),
                debt=debt_balance,
                equity_book=eq_book,
                weight_debt=w_d,
                weight_equity=w_e,
                wacc=wacc[year],
            ))
        
        # Discount schedule
        discount_schedule = []
        for i, year in enumerate(forecast_years):
            period = i + 1
            discount_schedule.append(YearlyDiscount(
                year=year,
                period=period,
                wacc=wacc[year],
                ke=ke,
                discount_factor_wacc=df_wacc[year],
                discount_factor_ke=df_ke[year],
                fcff=fcff[year],
                fcfe=fcfe[year],
                pv_fcff=pv_fcff[year],
                pv_fcfe=pv_fcfe[year],
            ))
        
        # Terminal value output
        tv_output = TerminalValueOutput(
            method=inputs.terminal_value.method,
            final_year=final_year,
            growth_rate=inputs.terminal_value.perpetuity_growth_rate,
            exit_multiple=inputs.terminal_value.exit_multiple,
            exit_metric=inputs.terminal_value.exit_metric.value if inputs.terminal_value.exit_metric else None,
            metric_value=None,  # Set for exit multiple
            terminal_value_fcff=tv_fcff,
            terminal_value_fcfe=tv_fcfe,
            discount_rate_wacc=final_wacc,
            discount_rate_ke=ke,
            pv_terminal_value_fcff=pv_tv_fcff,
            pv_terminal_value_fcfe=pv_tv_fcfe,
        )
        
        # Valuation bridge
        valuation_bridge = ValuationBridge(
            sum_pv_fcff=sum_pv_fcff,
            sum_pv_fcfe=sum_pv_fcfe,
            pv_terminal_value_fcff=pv_tv_fcff,
            pv_terminal_value_fcfe=pv_tv_fcfe,
            enterprise_value=ev,
            debt_at_base=base_debt,
            cash_at_base=cash,
            net_debt=net_debt_value,
            equity_value_from_ev=equity_from_ev_value,
            equity_value_direct=equity_direct_value,
            reconciliation_difference=recon_diff,
            reconciliation_notes=recon_notes,
        )
        
        return DCFOutputs(
            base_year=base_year,
            forecast_years=forecast_years,
            discounting_mode=inputs.discounting_mode,
            projections=projections,
            nwc_schedule=nwc_schedule,
            cash_flows=cash_flows,
            ke=ke,
            wacc_details=wacc_details,
            discount_schedule=discount_schedule,
            terminal_value=tv_output,
            valuation_bridge=valuation_bridge,
            equity_book_values=equity_book,
        )
