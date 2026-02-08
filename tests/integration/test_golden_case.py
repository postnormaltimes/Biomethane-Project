"""
Integration Test - Golden Case

Full end-to-end test using the embedded golden case inputs and expected outputs.
All values must match within specified tolerances.
"""
import pytest
from dcf_engine.engine import DCFEngine
from dcf_engine.models import (
    DCFInputs,
    TimelineInputs,
    RevenueInputs,
    OperatingInputs,
    NWCInputs,
    InvestmentInputs,
    TaxInputs,
    CAPMInputs,
    DebtInputs,
    WACCInputs,
    EquityBookInputs,
    TerminalValueInputs,
    NetDebtInputs,
    DiscountingMode,
    TerminalValueMethod,
    WeightingMode,
)


# ============================================================================
# Tolerances
# ============================================================================

RATE_TOLERANCE = 1e-4      # For rates/percentages
CURRENCY_TOLERANCE = 1e-2  # For currency values


# ============================================================================
# Golden Case Inputs
# ============================================================================

def create_golden_inputs() -> DCFInputs:
    """
    Create the golden case inputs as specified in the requirements.
    
    Horizon: t0=2022, forecast 2023-2025 (N=3)
    Discounting mode: year_specific_flat
    """
    return DCFInputs(
        discounting_mode=DiscountingMode.YEAR_SPECIFIC_FLAT,
        timeline=TimelineInputs(
            base_year=2022,
            forecast_years=[2023, 2024, 2025],
        ),
        revenue=RevenueInputs(
            base_revenue=12500.0,
            growth_rates={
                2023: 0.15,
                2024: 0.10,
                2025: 0.10,
            },
        ),
        operating=OperatingInputs(
            cost_ratios={
                2023: 0.85,
                2024: 0.83,
                2025: 0.80,
            },
            depreciation_amortization={
                2022: 500.0,
                2023: 550.0,
                2024: 650.0,
                2025: 700.0,
            },
        ),
        nwc=NWCInputs(
            nwc_percent={
                2022: 0.16,
                2023: 0.16,
                2024: 0.13,
                2025: 0.10,
            },
        ),
        investments=InvestmentInputs(
            capex={
                2023: 800.0,
                2024: 900.0,
                2025: 1000.0,
            },
        ),
        tax=TaxInputs(
            tax_rate=0.30,
        ),
        capm=CAPMInputs(
            rf=0.04,
            rm=0.10,
            beta=1.30,
        ),
        debt=DebtInputs(
            debt_balances={
                2022: 1500.0,
                2023: 2050.0,
                2024: 2055.63,
                2025: 2039.38,
            },
            rd={
                2022: 0.05,
                2023: 0.06,
                2024: 0.065,
                2025: 0.065,
            },
        ),
        wacc=WACCInputs(
            weighting_mode=WeightingMode.BOOK_VALUE,
            equity_book_inputs=EquityBookInputs(
                base_equity_book=10000.0,
                # Dividends and NewEquity default to 0
            ),
        ),
        terminal_value=TerminalValueInputs(
            method=TerminalValueMethod.PERPETUITY,
            g=0.0,  # Zero growth perpetuity
        ),
        net_debt=NetDebtInputs(
            cash_and_equivalents=1492.10,
        ),
    )


# ============================================================================
# Golden Case Expected Outputs
# ============================================================================

EXPECTED_REVENUE = {
    2023: 14375.0,
    2024: 15812.5,
    2025: 17393.75,
}

EXPECTED_EBITDA = {
    2023: 2156.25,
    2024: 2688.125,
    2025: 3478.75,
}

EXPECTED_EBIT = {
    2023: 1606.25,
    2024: 2038.125,
    2025: 2778.75,
}

EXPECTED_NWC = {
    2022: 2000.0,
    2023: 2300.0,
    2024: 2055.625,
    2025: 1739.375,
}

EXPECTED_DELTA_NWC = {
    2023: 300.0,
    2024: -244.375,
    2025: -316.25,
}

EXPECTED_TAX_ON_EBIT = {
    2023: 481.875,
    2024: 611.4375,
    2025: 833.625,
}

EXPECTED_FCFF = {
    2023: 574.375,
    2024: 1421.0625,
    2025: 1961.375,
}

EXPECTED_KE = 0.118

EXPECTED_WACC = {
    2023: 0.1060962159,
    2024: 0.1076698869,
    2025: 0.1089085817,
}

EXPECTED_PV_FCFF = {
    2023: 519.28122685,
    2024: 1158.22378917,
    2025: 1438.37922612,
}

EXPECTED_SUM_PV_FCFF = 3115.88424214

EXPECTED_TV = 18009.37051479

EXPECTED_PV_TV = 13207.21658219

EXPECTED_EV = 16323.10082433

EXPECTED_NET_DEBT = 7.90

EXPECTED_EQUITY = 16315.20082433


# ============================================================================
# Integration Tests
# ============================================================================

class TestGoldenCase:
    """Full integration test against the golden case."""
    
    @pytest.fixture
    def outputs(self):
        """Run the engine and return outputs."""
        inputs = create_golden_inputs()
        engine = DCFEngine(inputs)
        return engine.run()
    
    # ========================================================================
    # Revenue Tests
    # ========================================================================
    
    def test_revenue_2023(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2023)
        assert proj.revenue == pytest.approx(EXPECTED_REVENUE[2023], abs=CURRENCY_TOLERANCE)
    
    def test_revenue_2024(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2024)
        assert proj.revenue == pytest.approx(EXPECTED_REVENUE[2024], abs=CURRENCY_TOLERANCE)
    
    def test_revenue_2025(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2025)
        assert proj.revenue == pytest.approx(EXPECTED_REVENUE[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # EBITDA Tests
    # ========================================================================
    
    def test_ebitda_2023(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2023)
        assert proj.ebitda == pytest.approx(EXPECTED_EBITDA[2023], abs=CURRENCY_TOLERANCE)
    
    def test_ebitda_2024(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2024)
        assert proj.ebitda == pytest.approx(EXPECTED_EBITDA[2024], abs=CURRENCY_TOLERANCE)
    
    def test_ebitda_2025(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2025)
        assert proj.ebitda == pytest.approx(EXPECTED_EBITDA[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # EBIT Tests
    # ========================================================================
    
    def test_ebit_2023(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2023)
        assert proj.ebit == pytest.approx(EXPECTED_EBIT[2023], abs=CURRENCY_TOLERANCE)
    
    def test_ebit_2024(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2024)
        assert proj.ebit == pytest.approx(EXPECTED_EBIT[2024], abs=CURRENCY_TOLERANCE)
    
    def test_ebit_2025(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2025)
        assert proj.ebit == pytest.approx(EXPECTED_EBIT[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # NWC Tests
    # ========================================================================
    
    def test_nwc_2022(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2022)
        assert nwc.nwc == pytest.approx(EXPECTED_NWC[2022], abs=CURRENCY_TOLERANCE)
    
    def test_nwc_2023(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2023)
        assert nwc.nwc == pytest.approx(EXPECTED_NWC[2023], abs=CURRENCY_TOLERANCE)
    
    def test_nwc_2024(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2024)
        assert nwc.nwc == pytest.approx(EXPECTED_NWC[2024], abs=CURRENCY_TOLERANCE)
    
    def test_nwc_2025(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2025)
        assert nwc.nwc == pytest.approx(EXPECTED_NWC[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # ΔNWC Tests
    # ========================================================================
    
    def test_delta_nwc_2023(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2023)
        assert nwc.delta_nwc == pytest.approx(EXPECTED_DELTA_NWC[2023], abs=CURRENCY_TOLERANCE)
    
    def test_delta_nwc_2024(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2024)
        assert nwc.delta_nwc == pytest.approx(EXPECTED_DELTA_NWC[2024], abs=CURRENCY_TOLERANCE)
    
    def test_delta_nwc_2025(self, outputs):
        nwc = next(n for n in outputs.nwc_schedule if n.year == 2025)
        assert nwc.delta_nwc == pytest.approx(EXPECTED_DELTA_NWC[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # Tax on EBIT Tests
    # ========================================================================
    
    def test_tax_on_ebit_2023(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2023)
        assert proj.tax_on_ebit == pytest.approx(EXPECTED_TAX_ON_EBIT[2023], abs=CURRENCY_TOLERANCE)
    
    def test_tax_on_ebit_2024(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2024)
        assert proj.tax_on_ebit == pytest.approx(EXPECTED_TAX_ON_EBIT[2024], abs=CURRENCY_TOLERANCE)
    
    def test_tax_on_ebit_2025(self, outputs):
        proj = next(p for p in outputs.projections if p.year == 2025)
        assert proj.tax_on_ebit == pytest.approx(EXPECTED_TAX_ON_EBIT[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # FCFF Tests
    # ========================================================================
    
    def test_fcff_2023(self, outputs):
        cf = next(c for c in outputs.cash_flows if c.year == 2023)
        assert cf.fcff == pytest.approx(EXPECTED_FCFF[2023], abs=CURRENCY_TOLERANCE)
    
    def test_fcff_2024(self, outputs):
        cf = next(c for c in outputs.cash_flows if c.year == 2024)
        assert cf.fcff == pytest.approx(EXPECTED_FCFF[2024], abs=CURRENCY_TOLERANCE)
    
    def test_fcff_2025(self, outputs):
        cf = next(c for c in outputs.cash_flows if c.year == 2025)
        assert cf.fcff == pytest.approx(EXPECTED_FCFF[2025], abs=CURRENCY_TOLERANCE)
    
    # ========================================================================
    # Ke Test
    # ========================================================================
    
    def test_ke(self, outputs):
        assert outputs.ke == pytest.approx(EXPECTED_KE, abs=RATE_TOLERANCE)
    
    # ========================================================================
    # WACC Tests
    # ========================================================================
    
    def test_wacc_2023(self, outputs):
        wacc = next(w for w in outputs.wacc_details if w.year == 2023)
        assert wacc.wacc == pytest.approx(EXPECTED_WACC[2023], abs=RATE_TOLERANCE)
    
    def test_wacc_2024(self, outputs):
        wacc = next(w for w in outputs.wacc_details if w.year == 2024)
        assert wacc.wacc == pytest.approx(EXPECTED_WACC[2024], abs=RATE_TOLERANCE)
    
    def test_wacc_2025(self, outputs):
        wacc = next(w for w in outputs.wacc_details if w.year == 2025)
        assert wacc.wacc == pytest.approx(EXPECTED_WACC[2025], abs=RATE_TOLERANCE)
    
    # ========================================================================
    # PV(FCFF) Tests
    # ========================================================================
    
    def test_pv_fcff_2023(self, outputs):
        disc = next(d for d in outputs.discount_schedule if d.year == 2023)
        assert disc.pv_fcff == pytest.approx(EXPECTED_PV_FCFF[2023], abs=CURRENCY_TOLERANCE)
    
    def test_pv_fcff_2024(self, outputs):
        disc = next(d for d in outputs.discount_schedule if d.year == 2024)
        assert disc.pv_fcff == pytest.approx(EXPECTED_PV_FCFF[2024], abs=CURRENCY_TOLERANCE)
    
    def test_pv_fcff_2025(self, outputs):
        disc = next(d for d in outputs.discount_schedule if d.year == 2025)
        assert disc.pv_fcff == pytest.approx(EXPECTED_PV_FCFF[2025], abs=CURRENCY_TOLERANCE)
    
    def test_sum_pv_fcff(self, outputs):
        assert outputs.valuation_bridge.sum_pv_fcff == pytest.approx(
            EXPECTED_SUM_PV_FCFF, abs=CURRENCY_TOLERANCE
        )
    
    # ========================================================================
    # Terminal Value Tests
    # ========================================================================
    
    def test_terminal_value(self, outputs):
        assert outputs.terminal_value.terminal_value_fcff == pytest.approx(
            EXPECTED_TV, abs=CURRENCY_TOLERANCE
        )
    
    def test_pv_terminal_value(self, outputs):
        assert outputs.terminal_value.pv_terminal_value_fcff == pytest.approx(
            EXPECTED_PV_TV, abs=CURRENCY_TOLERANCE
        )
    
    # ========================================================================
    # Enterprise Value and Equity Tests
    # ========================================================================
    
    def test_enterprise_value(self, outputs):
        assert outputs.valuation_bridge.enterprise_value == pytest.approx(
            EXPECTED_EV, abs=CURRENCY_TOLERANCE
        )
    
    def test_net_debt(self, outputs):
        assert outputs.valuation_bridge.net_debt == pytest.approx(
            EXPECTED_NET_DEBT, abs=CURRENCY_TOLERANCE
        )
    
    def test_equity_value_from_ev(self, outputs):
        assert outputs.valuation_bridge.equity_value_from_ev == pytest.approx(
            EXPECTED_EQUITY, abs=CURRENCY_TOLERANCE
        )
    
    # ========================================================================
    # Alternative FCFF Calculation Check
    # ========================================================================
    
    def test_fcff_alternative_formula(self, outputs):
        """
        Verify FCFF matches alternative formula:
        FCFF = EBITDA - TaxOnEBIT - ΔNWC - Capex
        """
        for cf in outputs.cash_flows:
            proj = next(p for p in outputs.projections if p.year == cf.year)
            nwc = next(n for n in outputs.nwc_schedule if n.year == cf.year)
            
            # Alternative formula
            alt_fcff = (
                proj.ebitda
                - proj.tax_on_ebit
                - nwc.delta_nwc
                - cf.capex
            )
            
            assert cf.fcff == pytest.approx(alt_fcff, abs=CURRENCY_TOLERANCE)


# ============================================================================
# Additional Validation Tests
# ============================================================================

class TestGoldenCaseValidation:
    """Additional validation tests for the golden case."""
    
    @pytest.fixture
    def outputs(self):
        """Run the engine and return outputs."""
        inputs = create_golden_inputs()
        engine = DCFEngine(inputs)
        return engine.run()
    
    def test_discounting_mode_is_year_specific_flat(self, outputs):
        """Verify the correct discounting mode is used."""
        assert outputs.discounting_mode == DiscountingMode.YEAR_SPECIFIC_FLAT
    
    def test_forecast_years(self, outputs):
        """Verify forecast years match expected."""
        assert outputs.forecast_years == [2023, 2024, 2025]
    
    def test_base_year(self, outputs):
        """Verify base year matches expected."""
        assert outputs.base_year == 2022
    
    def test_tv_method_is_perpetuity(self, outputs):
        """Verify terminal value method."""
        assert outputs.terminal_value.method == TerminalValueMethod.PERPETUITY
    
    def test_tv_growth_is_zero(self, outputs):
        """Verify zero growth rate is used."""
        assert outputs.terminal_value.growth_rate == 0.0


# ============================================================================
# Smoke Test
# ============================================================================

def test_engine_runs_without_error():
    """Basic smoke test that engine completes without exceptions."""
    inputs = create_golden_inputs()
    engine = DCFEngine(inputs)
    outputs = engine.run()
    
    assert outputs is not None
    assert len(outputs.projections) == 3
    assert len(outputs.cash_flows) == 3
    assert len(outputs.discount_schedule) == 3
    assert outputs.valuation_bridge is not None
    assert outputs.terminal_value is not None
