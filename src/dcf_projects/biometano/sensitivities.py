"""
Biometano Sensitivity Analysis Module

Provides expanded sensitivity analysis framework including:
- One-at-a-time shock runner with multiple shock levels
- Scenario comparison (Base/Upside/Downside with bundled assumptions)
- Tornado chart data generation (EV-based by default)
- Two-way sensitivity grids
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Callable, Optional
import copy

from dcf_projects.biometano.schema import BiometanoCase


class SensitivityParameter(str, Enum):
    """Sensitivity parameters that can be shocked."""
    GATE_FEE = "gate_fee"
    TARIFF = "tariff"
    GO_PRICE = "go_price"
    CO2_PRICE = "co2_price"
    COMPOST_PRICE = "compost_price"
    THROUGHPUT = "throughput"
    IMPURITY_RATE = "impurity_rate"  # Sovvalli
    AVAILABILITY = "availability"
    ELECTRICITY_PRICE = "electricity_price"
    CAPEX = "capex"
    OPEX_ESCALATION = "opex_escalation"
    WACC = "wacc"
    PNRR_GRANT = "pnrr_grant"
    ZES_TIMING = "zes_timing"


@dataclass
class SensitivityShock:
    """Definition of a sensitivity shock."""
    parameter: SensitivityParameter
    name: str
    low_pct: float  # e.g., -0.20 for -20%
    high_pct: float  # e.g., 0.20 for +20%
    units: str = "%"
    extra_shocks: list[float] = field(default_factory=list)  # Additional shock levels


@dataclass
class SensitivityResult:
    """Result of a single sensitivity shock."""
    parameter: SensitivityParameter
    shock_name: str
    shock_value: float  # The % or absolute shock applied
    base_value: float  # Base case EV
    shocked_value: float  # Shocked case EV
    base_equity: float
    shocked_equity: float
    delta: float  # Change from base EV
    delta_pct: float  # % change from base EV


@dataclass
class TornadoData:
    """Data for tornado chart - EV-based."""
    parameter: str
    low_label: str
    high_label: str
    # EV values (primary)
    low_ev: float
    high_ev: float
    base_ev: float
    low_delta_ev: float
    high_delta_ev: float
    spread_ev: float
    # Equity values (secondary)
    low_value: float  # legacy compatibility
    high_value: float
    base_value: float
    low_delta: float
    high_delta: float
    spread: float


@dataclass
class ScenarioResult:
    """Result for a named scenario with detailed metrics."""
    name: str
    description: str
    adjustments: dict[str, float]  # parameter -> multiplier
    # EV metrics (primary)
    ev: float
    pv_fcff: float
    pv_tv: float
    delta_ev: float
    delta_ev_pct: float
    # Equity metrics (secondary)
    equity_value: float
    delta_from_base: float
    delta_pct: float


@dataclass
class TwoWayGrid:
    """Two-way sensitivity grid."""
    row_param: str
    col_param: str
    row_values: list[float]
    col_values: list[float]
    ev_grid: list[list[float]]  # [row_idx][col_idx] = EV


@dataclass
class SensitivityAnalysisOutputs:
    """Complete sensitivity analysis outputs."""
    base_ev: float
    base_equity_value: float
    
    # One-at-a-time results
    individual_results: list[SensitivityResult] = field(default_factory=list)
    
    # Tornado data (sorted by spread)
    tornado_data: list[TornadoData] = field(default_factory=list)
    
    # Scenarios
    scenarios: list[ScenarioResult] = field(default_factory=list)
    
    # Two-way grids
    two_way_grids: list[TwoWayGrid] = field(default_factory=list)


# Expanded sensitivity shocks per requirements
DEFAULT_SHOCKS: list[SensitivityShock] = [
    # Revenue parameters
    SensitivityShock(SensitivityParameter.GATE_FEE, "Gate Fee", -0.20, 0.20, extra_shocks=[-0.10, 0.10]),
    SensitivityShock(SensitivityParameter.TARIFF, "Tariff", -0.20, 0.20, extra_shocks=[-0.10, 0.10]),
    SensitivityShock(SensitivityParameter.GO_PRICE, "GO Price", -0.50, 0.50, extra_shocks=[-0.25, 0.25]),
    SensitivityShock(SensitivityParameter.CO2_PRICE, "CO₂ Price", -0.40, 0.40, extra_shocks=[-0.20, 0.20]),
    SensitivityShock(SensitivityParameter.COMPOST_PRICE, "Compost", -0.50, 0.50),
    
    # Operational parameters
    SensitivityShock(SensitivityParameter.THROUGHPUT, "Throughput", -0.10, 0.10),
    SensitivityShock(SensitivityParameter.IMPURITY_RATE, "Sovvalli %", 0.02, 0.05, units="pp"),  # 18% -> 20%, 23%
    SensitivityShock(SensitivityParameter.AVAILABILITY, "Availability", -0.10, 0.05),
    SensitivityShock(SensitivityParameter.ELECTRICITY_PRICE, "Electricity", -0.20, 0.40, extra_shocks=[0.20]),
    
    # Investment parameters
    SensitivityShock(SensitivityParameter.CAPEX, "CAPEX Overrun", -0.05, 0.20, extra_shocks=[0.05, 0.10]),
    SensitivityShock(SensitivityParameter.OPEX_ESCALATION, "OPEX Escalation", -0.02, 0.02, units="pp"),
    
    # Discount rate
    SensitivityShock(SensitivityParameter.WACC, "WACC", -0.01, 0.02, units="pp", extra_shocks=[0.01]),
    
    # Incentives
    SensitivityShock(SensitivityParameter.PNRR_GRANT, "PNRR Grant", -0.50, -1.0),  # 50% reduction, full removal
    SensitivityShock(SensitivityParameter.ZES_TIMING, "ZES Timing", 1.0, 2.0, units="years"),  # Delay 1-2 years
]


class SensitivityRunner:
    """
    Runs sensitivity analysis on a Biometano case.
    
    Supports:
    - One-at-a-time shocks (EV-based)
    - Combined scenarios (Base/Upside/Downside)
    - Two-way sensitivity grids
    - Tornado chart generation
    """
    
    def __init__(
        self, 
        case: BiometanoCase,
        value_function: Optional[Callable[[BiometanoCase], tuple[float, float, float, float]]] = None,
    ):
        """
        Initialize runner.
        
        Args:
            case: Base case inputs
            value_function: Function that takes case and returns (equity_value, ev, pv_fcff, pv_tv).
                           If None, uses default valuation.
        """
        self.base_case = case
        self.value_function = value_function or self._default_value_function
        
    def run_full_analysis(
        self,
        shocks: Optional[list[SensitivityShock]] = None,
        include_scenarios: bool = True,
        include_two_way: bool = True,
    ) -> SensitivityAnalysisOutputs:
        """
        Run complete sensitivity analysis.
        
        Args:
            shocks: Sensitivity shocks to run (defaults to DEFAULT_SHOCKS)
            include_scenarios: Whether to include Base/Upside/Downside scenarios
            include_two_way: Whether to include two-way sensitivity grids
            
        Returns:
            Complete sensitivity outputs
        """
        if shocks is None:
            shocks = DEFAULT_SHOCKS
        
        # Get base case values
        base_equity, base_ev, base_pv_fcff, base_pv_tv = self.value_function(self.base_case)
        
        outputs = SensitivityAnalysisOutputs(
            base_ev=base_ev,
            base_equity_value=base_equity,
        )
        
        # Run individual shocks
        for shock in shocks:
            low_result = self._run_shock(shock, shock.low_pct, base_ev, base_equity)
            high_result = self._run_shock(shock, shock.high_pct, base_ev, base_equity)
            
            outputs.individual_results.append(low_result)
            outputs.individual_results.append(high_result)
            
            # Build tornado data (EV-based)
            tornado = TornadoData(
                parameter=shock.name,
                low_label=self._format_shock_label(shock.low_pct, shock.units),
                high_label=self._format_shock_label(shock.high_pct, shock.units),
                # EV values
                low_ev=low_result.shocked_value,
                high_ev=high_result.shocked_value,
                base_ev=base_ev,
                low_delta_ev=low_result.delta,
                high_delta_ev=high_result.delta,
                spread_ev=abs(high_result.delta - low_result.delta),
                # Equity values (legacy)
                low_value=low_result.shocked_equity,
                high_value=high_result.shocked_equity,
                base_value=base_equity,
                low_delta=low_result.shocked_equity - base_equity,
                high_delta=high_result.shocked_equity - base_equity,
                spread=abs(high_result.shocked_equity - low_result.shocked_equity),
            )
            outputs.tornado_data.append(tornado)
        
        # Sort tornado by EV spread (largest first)
        outputs.tornado_data.sort(key=lambda x: x.spread_ev, reverse=True)
        
        # Run scenarios
        if include_scenarios:
            outputs.scenarios = self._run_scenarios(base_ev, base_equity, base_pv_fcff, base_pv_tv)
        
        # Run two-way grids
        if include_two_way:
            outputs.two_way_grids = self._run_two_way_grids(base_ev)
        
        return outputs
    
    def _format_shock_label(self, shock: float, units: str) -> str:
        """Format shock value for display."""
        if units == "pp":
            return f"{shock*100:+.0f}pp"
        elif units == "years":
            return f"+{shock:.0f}y"
        else:
            return f"{shock:+.0%}"
    
    def _run_shock(
        self,
        shock: SensitivityShock,
        shock_pct: float,
        base_ev: float,
        base_equity: float,
    ) -> SensitivityResult:
        """Run a single shock and return result."""
        shocked_case = self._apply_shock(self.base_case, shock.parameter, shock_pct)
        shocked_equity, shocked_ev, _, _ = self.value_function(shocked_case)
        
        delta = shocked_ev - base_ev
        delta_pct = delta / base_ev if base_ev != 0 else 0.0
        
        return SensitivityResult(
            parameter=shock.parameter,
            shock_name=shock.name,
            shock_value=shock_pct,
            base_value=base_ev,
            shocked_value=shocked_ev,
            base_equity=base_equity,
            shocked_equity=shocked_equity,
            delta=delta,
            delta_pct=delta_pct,
        )
    
    def _apply_shock(
        self,
        case: BiometanoCase,
        parameter: SensitivityParameter,
        shock_pct: float,
    ) -> BiometanoCase:
        """Apply a shock to create a modified case."""
        # Deep copy the case
        shocked = case.model_copy(deep=True)
        
        if parameter == SensitivityParameter.GATE_FEE:
            shocked.revenues.gate_fee.price *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.TARIFF:
            shocked.revenues.tariff.price *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.GO_PRICE:
            shocked.revenues.go.price *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.CO2_PRICE:
            shocked.revenues.co2.price *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.COMPOST_PRICE:
            shocked.revenues.compost.price *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.THROUGHPUT:
            shocked.production.forsu_throughput_tpy *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.IMPURITY_RATE:
            # Additive shock to impurity rate
            current = getattr(shocked.production, 'impurity_rate', 0.20)
            shocked.production.impurity_rate = min(1.0, max(0.0, current + shock_pct))
            
        elif parameter == SensitivityParameter.AVAILABILITY:
            # Shock all availability values
            shocked.production.availability_profile = [
                min(1.0, max(0.0, a * (1 + shock_pct)))
                for a in shocked.production.availability_profile
            ]
            
        elif parameter == SensitivityParameter.ELECTRICITY_PRICE:
            # Shock utility costs
            if hasattr(shocked.opex.utilities, 'fixed_annual'):
                shocked.opex.utilities.fixed_annual *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.CAPEX:
            # Shock all CAPEX items
            for name, item in shocked.capex.all_items().items():
                new_amount = item.amount * (1 + shock_pct)
                setattr(getattr(shocked.capex, name), "amount", new_amount)
            
        elif parameter == SensitivityParameter.OPEX_ESCALATION:
            # Shock escalation rate (additive, not multiplicative)
            for name, cat in shocked.opex.all_categories().items():
                new_escalation = cat.escalation_rate + shock_pct
                setattr(getattr(shocked.opex, name), "escalation_rate", new_escalation)
            
        elif parameter == SensitivityParameter.WACC:
            # Shock risk-free rate (affects both WACC and Ke)
            shocked.financing.rf += shock_pct
            
        elif parameter == SensitivityParameter.PNRR_GRANT:
            if shock_pct == -1.0:
                # Full removal
                shocked.incentives.capital_grant.enabled = False
            elif shocked.incentives.capital_grant.percent_of_eligible is not None:
                shocked.incentives.capital_grant.percent_of_eligible *= (1 + shock_pct)
            
        elif parameter == SensitivityParameter.ZES_TIMING:
            # Delay in years (absolute)
            current = shocked.incentives.tax_credit.usable_from_year or shocked.horizon.cod_year
            shocked.incentives.tax_credit.usable_from_year = current + int(shock_pct)
        
        return shocked
    
    def _run_scenarios(
        self, 
        base_ev: float, 
        base_equity: float,
        base_pv_fcff: float,
        base_pv_tv: float,
    ) -> list[ScenarioResult]:
        """Run predefined scenarios: Base, Upside, Downside."""
        scenarios = []
        
        # Base case
        scenarios.append(ScenarioResult(
            name="Base",
            description="Base case assumptions",
            adjustments={},
            ev=base_ev,
            pv_fcff=base_pv_fcff,
            pv_tv=base_pv_tv,
            delta_ev=0.0,
            delta_ev_pct=0.0,
            equity_value=base_equity,
            delta_from_base=0.0,
            delta_pct=0.0,
        ))
        
        # Upside case - favorable bundled assumptions
        upside_adjustments = {
            "gate_fee": 0.10,       # +10% gate fee
            "go_price": 0.25,       # +25% GO price
            "impurity_rate": -0.02, # Lower impurity (18% -> 16%)
            "availability": 0.05,   # +5% availability
            "capex": -0.05,         # -5% CAPEX
            "opex_escalation": -0.01,  # -1pp OPEX escalation
        }
        upside_case = self._apply_scenario(self.base_case, upside_adjustments)
        upside_equity, upside_ev, upside_pv_fcff, upside_pv_tv = self.value_function(upside_case)
        scenarios.append(ScenarioResult(
            name="Upside",
            description="+10% gate fee, +25% GO, -2pp impurity, +5% avail, -5% CAPEX",
            adjustments=upside_adjustments,
            ev=upside_ev,
            pv_fcff=upside_pv_fcff,
            pv_tv=upside_pv_tv,
            delta_ev=upside_ev - base_ev,
            delta_ev_pct=(upside_ev - base_ev) / base_ev if base_ev != 0 else 0.0,
            equity_value=upside_equity,
            delta_from_base=upside_equity - base_equity,
            delta_pct=(upside_equity - base_equity) / base_equity if base_equity != 0 else 0.0,
        ))
        
        # Downside case - adverse bundled assumptions
        downside_adjustments = {
            "gate_fee": -0.15,      # -15% gate fee
            "impurity_rate": 0.04,  # Higher impurity (18% -> 22%)
            "electricity_price": 0.30,  # +30% electricity
            "capex": 0.15,          # +15% CAPEX overrun
            "pnrr_delay": 1.0,      # 1 year grant delay (if modeled)
            "availability": -0.05,  # -5% availability
        }
        downside_case = self._apply_scenario(self.base_case, downside_adjustments)
        downside_equity, downside_ev, downside_pv_fcff, downside_pv_tv = self.value_function(downside_case)
        scenarios.append(ScenarioResult(
            name="Downside",
            description="-15% gate fee, +4pp impurity, +30% elec, +15% CAPEX, -5% avail",
            adjustments=downside_adjustments,
            ev=downside_ev,
            pv_fcff=downside_pv_fcff,
            pv_tv=downside_pv_tv,
            delta_ev=downside_ev - base_ev,
            delta_ev_pct=(downside_ev - base_ev) / base_ev if base_ev != 0 else 0.0,
            equity_value=downside_equity,
            delta_from_base=downside_equity - base_equity,
            delta_pct=(downside_equity - base_equity) / base_equity if base_equity != 0 else 0.0,
        ))
        
        return scenarios
    
    def _apply_scenario(
        self,
        case: BiometanoCase,
        adjustments: dict[str, float],
    ) -> BiometanoCase:
        """Apply multiple adjustments for a scenario."""
        shocked = case.model_copy(deep=True)
        
        for param, shock in adjustments.items():
            if param == "gate_fee":
                shocked.revenues.gate_fee.price *= (1 + shock)
            elif param == "tariff":
                shocked.revenues.tariff.price *= (1 + shock)
            elif param == "go_price":
                shocked.revenues.go.price *= (1 + shock)
            elif param == "impurity_rate":
                current = getattr(shocked.production, 'impurity_rate', 0.20)
                shocked.production.impurity_rate = min(1.0, max(0.0, current + shock))
            elif param == "electricity_price":
                if hasattr(shocked.opex.utilities, 'fixed_annual'):
                    shocked.opex.utilities.fixed_annual *= (1 + shock)
            elif param == "availability":
                shocked.production.availability_profile = [
                    min(1.0, max(0.0, a * (1 + shock)))
                    for a in shocked.production.availability_profile
                ]
            elif param == "capex":
                for name, item in shocked.capex.all_items().items():
                    new_amount = item.amount * (1 + shock)
                    setattr(getattr(shocked.capex, name), "amount", new_amount)
            elif param == "opex_escalation":
                for name, cat in shocked.opex.all_categories().items():
                    new_esc = cat.escalation_rate + shock
                    setattr(getattr(shocked.opex, name), "escalation_rate", new_esc)
        
        return shocked
    
    def _run_two_way_grids(self, base_ev: float) -> list[TwoWayGrid]:
        """Run two-way sensitivity grids."""
        grids = []
        
        # Gate Fee vs WACC
        gate_fee_shocks = [-0.20, -0.10, 0.0, 0.10, 0.20]
        wacc_shocks = [-0.02, -0.01, 0.0, 0.01, 0.02]
        
        ev_grid = []
        for wacc_shock in wacc_shocks:
            row = []
            for gf_shock in gate_fee_shocks:
                shocked = self.base_case.model_copy(deep=True)
                shocked.revenues.gate_fee.price *= (1 + gf_shock)
                shocked.financing.rf += wacc_shock
                _, ev, _, _ = self.value_function(shocked)
                row.append(ev)
            ev_grid.append(row)
        
        grids.append(TwoWayGrid(
            row_param="WACC",
            col_param="Gate Fee",
            row_values=wacc_shocks,
            col_values=gate_fee_shocks,
            ev_grid=ev_grid,
        ))
        
        return grids
    
    def _default_value_function(self, case: BiometanoCase) -> tuple[float, float, float, float]:
        """
        Default valuation function using built projections.
        
        Returns (equity_value, enterprise_value, pv_fcff, pv_tv).
        """
        from dcf_projects.biometano.builder import build_projections
        from dcf_projects.biometano.valuation import compute_valuation
        
        projections = build_projections(case)
        valuation = compute_valuation(case, projections)
        
        return (
            valuation.equity_value, 
            valuation.enterprise_value,
            valuation.sum_pv_fcff,
            valuation.pv_terminal_value_fcff,
        )


def run_sensitivity_analysis(
    case: BiometanoCase,
    shocks: Optional[list[SensitivityShock]] = None,
    value_function: Optional[Callable[[BiometanoCase], tuple[float, float, float, float]]] = None,
) -> SensitivityAnalysisOutputs:
    """
    Convenience function to run sensitivity analysis.
    
    Args:
        case: Base case inputs
        shocks: Sensitivity shocks (defaults to DEFAULT_SHOCKS)
        value_function: Custom valuation function
        
    Returns:
        Complete sensitivity outputs
    """
    runner = SensitivityRunner(case, value_function)
    return runner.run_full_analysis(shocks)
