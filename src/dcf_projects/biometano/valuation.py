"""
Biometano Valuation Module

Computes DCF valuation (EV and Equity Value) from projections.
Bridges to the core DCF engine for discounting and terminal value.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from dcf_projects.biometano.schema import BiometanoCase
from dcf_projects.biometano.builder import BiometanoProjections


@dataclass
class ValuationOutputs:
    """Complete valuation outputs."""
    base_year: int
    cod_year: int
    final_year: int
    
    # Discount rates
    ke: float = 0.0
    wacc: dict[int, float] = field(default_factory=dict)
    
    # Cash flows
    fcff: dict[int, float] = field(default_factory=dict)
    fcfe: dict[int, float] = field(default_factory=dict)
    
    # Present values (year -> PV)
    pv_fcff: dict[int, float] = field(default_factory=dict)
    pv_fcfe: dict[int, float] = field(default_factory=dict)
    
    sum_pv_fcff: float = 0.0
    sum_pv_fcfe: float = 0.0
    
    # Terminal value
    terminal_value_fcff: float = 0.0
    terminal_value_fcfe: float = 0.0
    pv_terminal_value_fcff: float = 0.0
    pv_terminal_value_fcfe: float = 0.0
    
    # Enterprise value
    enterprise_value: float = 0.0
    
    # Net debt at base
    debt_at_base: float = 0.0
    cash_at_base: float = 0.0
    net_debt: float = 0.0
    
    # Equity value
    equity_value: float = 0.0  # From FCFF/WACC
    equity_value_direct: float = 0.0  # From FCFE/Ke
    
    # Reconciliation
    reconciliation_difference: float = 0.0
    
    # Commentary
    commentary: list[str] = field(default_factory=list)


class ValuationCalculator:
    """
    Computes DCF valuation from projections.
    
    Uses existing DCF engine conventions for discounting.
    """
    
    def __init__(self, case: BiometanoCase, projections: BiometanoProjections):
        self.case = case
        self.projections = projections
        
    def compute(self) -> ValuationOutputs:
        """Compute complete valuation."""
        proj = self.projections
        fin = self.case.financing
        
        # Operating years only for valuation
        operating_years = proj.operating_years
        if not operating_years:
            operating_years = proj.all_forecast_years
        
        base_year = proj.base_year
        final_year = operating_years[-1] if operating_years else base_year
        
        outputs = ValuationOutputs(
            base_year=base_year,
            cod_year=proj.cod_year,
            final_year=final_year,
        )
        
        # Step 1: Compute Ke (CAPM)
        ke = self._compute_ke(fin)
        outputs.ke = ke
        
        # Step 2: Compute WACC by year
        wacc = self._compute_wacc(ke, operating_years)
        outputs.wacc = wacc
        
        # Step 3: Get cash flows
        outputs.fcff = {y: proj.fcff.get(y, 0.0) for y in operating_years}
        outputs.fcfe = {y: proj.fcfe.get(y, 0.0) for y in operating_years}
        
        # Step 4: Compute present values
        outputs.pv_fcff = self._compute_pv(outputs.fcff, wacc, base_year, proj.cod_year)
        outputs.pv_fcfe = self._compute_pv_constant_rate(outputs.fcfe, ke, base_year, proj.cod_year)
        
        outputs.sum_pv_fcff = sum(outputs.pv_fcff.values())
        outputs.sum_pv_fcfe = sum(outputs.pv_fcfe.values())
        
        # Step 5: Terminal value
        tv_inputs = self.case.terminal_value
        fcff_final = outputs.fcff.get(final_year, 0.0)
        fcfe_final = outputs.fcfe.get(final_year, 0.0)
        wacc_final = wacc.get(final_year, ke)
        g = tv_inputs.perpetuity_growth
        
        if tv_inputs.method == "perpetuity":
            if wacc_final > g:
                outputs.terminal_value_fcff = fcff_final * (1 + g) / (wacc_final - g)
            else:
                outputs.terminal_value_fcff = 0.0
            
            if ke > g:
                outputs.terminal_value_fcfe = fcfe_final * (1 + g) / (ke - g)
            else:
                outputs.terminal_value_fcfe = 0.0
        else:
            # Exit multiple (on EBITDA)
            ebitda_final = proj.ebitda.get(final_year, 0.0)
            multiple = tv_inputs.exit_multiple or 8.0
            outputs.terminal_value_fcff = ebitda_final * multiple
            outputs.terminal_value_fcfe = outputs.terminal_value_fcff  # Simplified
        
        # PV of terminal value
        n_periods = len(operating_years)
        if n_periods > 0 and wacc_final > 0:
            outputs.pv_terminal_value_fcff = outputs.terminal_value_fcff / ((1 + wacc_final) ** n_periods)
            outputs.pv_terminal_value_fcfe = outputs.terminal_value_fcfe / ((1 + ke) ** n_periods)
        
        # Step 6: Enterprise value
        outputs.enterprise_value = outputs.sum_pv_fcff + outputs.pv_terminal_value_fcff
        
        # Step 7: Net debt
        outputs.debt_at_base = fin.debt_amount  # Initial debt
        outputs.cash_at_base = fin.cash_at_base
        outputs.net_debt = outputs.debt_at_base - outputs.cash_at_base
        
        # Step 8: Equity value
        outputs.equity_value = outputs.enterprise_value - outputs.net_debt
        outputs.equity_value_direct = outputs.sum_pv_fcfe + outputs.pv_terminal_value_fcfe
        
        # Step 9: Reconciliation
        outputs.reconciliation_difference = outputs.equity_value - outputs.equity_value_direct
        
        # Step 10: Commentary
        outputs.commentary = self._generate_commentary(outputs)
        
        return outputs
    
    def _compute_ke(self, fin) -> float:
        """Compute cost of equity using CAPM."""
        if fin.ke_override is not None:
            return fin.ke_override
        
        ke = fin.rf + fin.beta * (fin.rm - fin.rf)
        return ke
    
    def _compute_wacc(self, ke: float, years: list[int]) -> dict[int, float]:
        """Compute WACC by year."""
        fin = self.case.financing
        tax_rate = fin.tax_rate
        
        wacc = {}
        for year in years:
            # Get debt balance for year
            year_fin = self.projections.get_financing(year)
            debt = year_fin.debt_balance_closing if year_fin else 0.0
            rd = year_fin.cost_of_debt if year_fin else fin.get_rd(year)
            
            # Equity value (use book value)
            equity = fin.equity_book_at_base + sum(
                self.projections.net_income.get(y, 0.0) 
                for y in years if y <= year
            )
            
            total = debt + equity
            if total > 0:
                wd = debt / total
                we = equity / total
            else:
                wd = 0.0
                we = 1.0
            
            wacc[year] = we * ke + wd * rd * (1 - tax_rate)
        
        return wacc
    
    def _compute_pv(
        self,
        cash_flows: dict[int, float],
        rates: dict[int, float],
        base_year: int,
        cod_year: int,
    ) -> dict[int, float]:
        """Compute present values using year-specific rates."""
        pv = {}
        for year, cf in cash_flows.items():
            # Period from COD
            period = year - cod_year + 1
            rate = rates.get(year, 0.10)
            if period > 0 and rate > -1:
                pv[year] = cf / ((1 + rate) ** period)
            else:
                pv[year] = cf
        return pv
    
    def _compute_pv_constant_rate(
        self,
        cash_flows: dict[int, float],
        rate: float,
        base_year: int,
        cod_year: int,
    ) -> dict[int, float]:
        """Compute present values using constant rate."""
        pv = {}
        for year, cf in cash_flows.items():
            period = year - cod_year + 1
            if period > 0 and rate > -1:
                pv[year] = cf / ((1 + rate) ** period)
            else:
                pv[year] = cf
        return pv
    
    def _generate_commentary(self, outputs: ValuationOutputs) -> list[str]:
        """Generate data-driven commentary."""
        commentary = []
        
        # EV composition
        if outputs.enterprise_value > 0:
            pv_flows_pct = outputs.sum_pv_fcff / outputs.enterprise_value * 100
            pv_tv_pct = outputs.pv_terminal_value_fcff / outputs.enterprise_value * 100
            commentary.append(
                f"Enterprise Value of €{outputs.enterprise_value:,.0f} comprises "
                f"€{outputs.sum_pv_fcff:,.0f} ({pv_flows_pct:.1f}%) from explicit forecast "
                f"and €{outputs.pv_terminal_value_fcff:,.0f} ({pv_tv_pct:.1f}%) from terminal value."
            )
        
        # TV share warning
        if outputs.enterprise_value > 0:
            tv_share = outputs.pv_terminal_value_fcff / outputs.enterprise_value
            if tv_share > 0.7:
                commentary.append(
                    f"⚠️ Terminal value represents {tv_share:.0%} of EV. "
                    "Consider extending explicit forecast period."
                )
        
        # Net debt bridge
        commentary.append(
            f"Net Debt of €{outputs.net_debt:,.0f} (Debt €{outputs.debt_at_base:,.0f} "
            f"less Cash €{outputs.cash_at_base:,.0f}) bridges EV to "
            f"Equity Value of €{outputs.equity_value:,.0f}."
        )
        
        # Reconciliation
        if abs(outputs.reconciliation_difference) > 1000:
            recon_pct = abs(outputs.reconciliation_difference / outputs.equity_value * 100) if outputs.equity_value != 0 else 0
            commentary.append(
                f"FCFF/WACC and FCFE/Ke methods differ by €{outputs.reconciliation_difference:,.0f} "
                f"({recon_pct:.1f}%) due to financing structure assumptions."
            )
        
        # Ke and WACC
        commentary.append(
            f"Cost of equity (Ke) = {outputs.ke:.2%}. "
            f"WACC ranges from {min(outputs.wacc.values()):.2%} to {max(outputs.wacc.values()):.2%}."
        )
        
        return commentary


def compute_valuation(
    case: BiometanoCase,
    projections: BiometanoProjections,
) -> ValuationOutputs:
    """
    Convenience function to compute valuation.
    
    Args:
        case: Biometano case inputs
        projections: Built projections
        
    Returns:
        Complete valuation outputs
    """
    calculator = ValuationCalculator(case, projections)
    return calculator.compute()
