"""
Unit Tests for DCF Engine Formulas

Tests for individual calculation functions.
"""
import pytest
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
from dcf_engine.discount_rates import compute_ke
from dcf_engine.discounting import (
    compute_discount_factors,
    compute_pv_series,
    compute_pv_single,
)
from dcf_engine.terminal_value import (
    compute_terminal_value_perpetuity,
    compute_terminal_value_exit_multiple,
    TerminalValueError,
)
from dcf_engine.valuation import (
    compute_enterprise_value,
    compute_net_debt,
    compute_equity_from_ev,
)
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
# Test fixtures
# ============================================================================

@pytest.fixture
def golden_inputs() -> DCFInputs:
    """Create inputs matching the golden case."""
    return DCFInputs(
        timeline=TimelineInputs(
            base_year=2022,
            forecast_years=[2023, 2024, 2025],
        ),
        revenue=RevenueInputs(
            base_revenue=12500.0,
            growth_rates={2023: 0.15, 2024: 0.10, 2025: 0.10},
        ),
        operating=OperatingInputs(
            cost_ratios={2023: 0.85, 2024: 0.83, 2025: 0.80},
            depreciation_amortization={2022: 500, 2023: 550, 2024: 650, 2025: 700},
        ),
        nwc=NWCInputs(
            nwc_percent={2022: 0.16, 2023: 0.16, 2024: 0.13, 2025: 0.10},
        ),
        investments=InvestmentInputs(
            capex={2023: 800, 2024: 900, 2025: 1000},
        ),
        tax=TaxInputs(tax_rate=0.30),
        capm=CAPMInputs(rf=0.04, rm=0.10, beta=1.30),
        debt=DebtInputs(
            debt_balances={2022: 1500, 2023: 2050, 2024: 2055.63, 2025: 2039.38},
            rd={2022: 0.05, 2023: 0.06, 2024: 0.065, 2025: 0.065},
        ),
        wacc=WACCInputs(
            weighting_mode=WeightingMode.BOOK_VALUE,
            equity_book_inputs=EquityBookInputs(base_equity_book=10000.0),
        ),
        terminal_value=TerminalValueInputs(
            method=TerminalValueMethod.PERPETUITY,
            g=0.0,
        ),
        net_debt=NetDebtInputs(cash_and_equivalents=1492.10),
        discounting_mode=DiscountingMode.YEAR_SPECIFIC_FLAT,
    )


# ============================================================================
# Revenue tests
# ============================================================================

class TestRevenue:
    def test_revenue_from_growth_rates(self, golden_inputs):
        """Test revenue projection from base + growth rates."""
        revenue = compute_revenue(golden_inputs)
        
        assert revenue[2022] == 12500.0
        assert revenue[2023] == pytest.approx(14375.0, abs=1e-2)
        assert revenue[2024] == pytest.approx(15812.5, abs=1e-2)
        assert revenue[2025] == pytest.approx(17393.75, abs=1e-2)


# ============================================================================
# EBITDA/EBIT tests
# ============================================================================

class TestOperating:
    def test_ebitda_from_cost_ratios(self, golden_inputs):
        """Test EBITDA calculation from cost ratios."""
        revenue = compute_revenue(golden_inputs)
        operating_costs = compute_operating_costs(golden_inputs, revenue)
        ebitda = compute_ebitda(golden_inputs, revenue, operating_costs)
        
        assert ebitda[2023] == pytest.approx(2156.25, abs=1e-2)
        assert ebitda[2024] == pytest.approx(2688.125, abs=1e-2)
        assert ebitda[2025] == pytest.approx(3478.75, abs=1e-2)
    
    def test_ebit_calculation(self, golden_inputs):
        """Test EBIT = EBITDA - D&A."""
        revenue = compute_revenue(golden_inputs)
        operating_costs = compute_operating_costs(golden_inputs, revenue)
        ebitda = compute_ebitda(golden_inputs, revenue, operating_costs)
        ebit = compute_ebit(golden_inputs, ebitda)
        
        assert ebit[2023] == pytest.approx(1606.25, abs=1e-2)
        assert ebit[2024] == pytest.approx(2038.125, abs=1e-2)
        assert ebit[2025] == pytest.approx(2778.75, abs=1e-2)


# ============================================================================
# NWC tests
# ============================================================================

class TestNWC:
    def test_nwc_from_percent(self, golden_inputs):
        """Test NWC calculation from % of revenue."""
        revenue = compute_revenue(golden_inputs)
        nwc = compute_nwc(golden_inputs, revenue)
        
        assert nwc[2022] == pytest.approx(2000.0, abs=1e-2)
        assert nwc[2023] == pytest.approx(2300.0, abs=1e-2)
        assert nwc[2024] == pytest.approx(2055.625, abs=1e-2)
        assert nwc[2025] == pytest.approx(1739.375, abs=1e-2)
    
    def test_delta_nwc(self, golden_inputs):
        """Test ΔNWC calculation."""
        revenue = compute_revenue(golden_inputs)
        nwc = compute_nwc(golden_inputs, revenue)
        delta_nwc = compute_delta_nwc(golden_inputs, nwc)
        
        assert delta_nwc[2023] == pytest.approx(300.0, abs=1e-2)
        assert delta_nwc[2024] == pytest.approx(-244.375, abs=1e-2)
        assert delta_nwc[2025] == pytest.approx(-316.25, abs=1e-2)
    
    def test_delta_nwc_sign_convention(self, golden_inputs):
        """Test that positive ΔNWC means cash consumed."""
        revenue = compute_revenue(golden_inputs)
        nwc = compute_nwc(golden_inputs, revenue)
        delta_nwc = compute_delta_nwc(golden_inputs, nwc)
        
        # 2023: NWC increased -> cash consumed -> positive
        assert delta_nwc[2023] > 0
        # 2024, 2025: NWC decreased -> cash released -> negative
        assert delta_nwc[2024] < 0
        assert delta_nwc[2025] < 0


# ============================================================================
# Tax tests
# ============================================================================

class TestTaxes:
    def test_tax_on_ebit(self, golden_inputs):
        """Test tax on EBIT (Mode A)."""
        revenue = compute_revenue(golden_inputs)
        operating_costs = compute_operating_costs(golden_inputs, revenue)
        ebitda = compute_ebitda(golden_inputs, revenue, operating_costs)
        ebit = compute_ebit(golden_inputs, ebitda)
        tax_on_ebit = compute_tax_on_ebit(golden_inputs, ebit)
        
        assert tax_on_ebit[2023] == pytest.approx(481.875, abs=1e-2)
        assert tax_on_ebit[2024] == pytest.approx(611.4375, abs=1e-2)
        assert tax_on_ebit[2025] == pytest.approx(833.625, abs=1e-2)
    
    def test_nopat(self, golden_inputs):
        """Test NOPAT = EBIT * (1 - TaxRate)."""
        revenue = compute_revenue(golden_inputs)
        operating_costs = compute_operating_costs(golden_inputs, revenue)
        ebitda = compute_ebitda(golden_inputs, revenue, operating_costs)
        ebit = compute_ebit(golden_inputs, ebitda)
        nopat = compute_nopat(golden_inputs, ebit)
        
        # NOPAT = EBIT * 0.70
        assert nopat[2023] == pytest.approx(1606.25 * 0.70, abs=1e-2)
        assert nopat[2024] == pytest.approx(2038.125 * 0.70, abs=1e-2)
        assert nopat[2025] == pytest.approx(2778.75 * 0.70, abs=1e-2)


# ============================================================================
# Cash flow tests
# ============================================================================

class TestCashFlows:
    def test_fcff_construction(self, golden_inputs):
        """Test FCFF = NOPAT + D&A - ΔNWC - Capex."""
        revenue = compute_revenue(golden_inputs)
        operating_costs = compute_operating_costs(golden_inputs, revenue)
        ebitda = compute_ebitda(golden_inputs, revenue, operating_costs)
        ebit = compute_ebit(golden_inputs, ebitda)
        nopat = compute_nopat(golden_inputs, ebit)
        nwc = compute_nwc(golden_inputs, revenue)
        delta_nwc = compute_delta_nwc(golden_inputs, nwc)
        da = golden_inputs.operating.depreciation_amortization
        capex = golden_inputs.investments.capex
        
        fcff = compute_fcff(nopat, da, delta_nwc, capex)
        
        assert fcff[2023] == pytest.approx(574.375, abs=1e-2)
        assert fcff[2024] == pytest.approx(1421.0625, abs=1e-2)
        assert fcff[2025] == pytest.approx(1961.375, abs=1e-2)
    
    def test_interest_expense_end_of_period(self, golden_inputs):
        """Test interest expense uses end-of-period convention."""
        interest = compute_interest_expense(golden_inputs)
        
        # Interest_t = Debt_t * rd_t
        assert interest[2023] == pytest.approx(2050 * 0.06, abs=1e-4)
        assert interest[2024] == pytest.approx(2055.63 * 0.065, abs=1e-4)
        assert interest[2025] == pytest.approx(2039.38 * 0.065, abs=1e-4)
    
    def test_net_borrowing(self, golden_inputs):
        """Test net borrowing = Debt_t - Debt_(t-1)."""
        net_borrowing = compute_net_borrowing(golden_inputs)
        
        assert net_borrowing[2023] == pytest.approx(2050 - 1500, abs=1e-2)
        assert net_borrowing[2024] == pytest.approx(2055.63 - 2050, abs=1e-2)
        assert net_borrowing[2025] == pytest.approx(2039.38 - 2055.63, abs=1e-2)


# ============================================================================
# Discount rate tests
# ============================================================================

class TestDiscountRates:
    def test_capm_ke(self, golden_inputs):
        """Test CAPM: Ke = rf + beta * (rm - rf)."""
        ke = compute_ke(golden_inputs)
        
        # Ke = 0.04 + 1.30 * (0.10 - 0.04) = 0.04 + 0.078 = 0.118
        assert ke == pytest.approx(0.118, abs=1e-4)


# ============================================================================
# Discounting tests
# ============================================================================

class TestDiscounting:
    def test_year_specific_flat_discounting(self):
        """Test year-specific-flat discounting mode."""
        rates = {2023: 0.10, 2024: 0.11, 2025: 0.12}
        years = [2023, 2024, 2025]
        
        df = compute_discount_factors(
            rates, years, 2022, DiscountingMode.YEAR_SPECIFIC_FLAT
        )
        
        # DF_i = 1 / (1 + r_i)^i
        assert df[2023] == pytest.approx(1 / (1.10 ** 1), abs=1e-6)
        assert df[2024] == pytest.approx(1 / (1.11 ** 2), abs=1e-6)
        assert df[2025] == pytest.approx(1 / (1.12 ** 3), abs=1e-6)
    
    def test_constant_discounting(self):
        """Test constant rate discounting mode."""
        rate = 0.10
        years = [2023, 2024, 2025]
        
        df = compute_discount_factors(
            rate, years, 2022, DiscountingMode.CONSTANT
        )
        
        # DF_i = 1 / (1 + r)^i
        assert df[2023] == pytest.approx(1 / (1.10 ** 1), abs=1e-6)
        assert df[2024] == pytest.approx(1 / (1.10 ** 2), abs=1e-6)
        assert df[2025] == pytest.approx(1 / (1.10 ** 3), abs=1e-6)
    
    def test_pv_single(self):
        """Test single value present value calculation."""
        pv = compute_pv_single(1000, 0.10, 3)
        assert pv == pytest.approx(1000 / (1.10 ** 3), abs=1e-6)


# ============================================================================
# Terminal value tests
# ============================================================================

class TestTerminalValue:
    def test_perpetuity_with_growth(self):
        """Test Gordon growth model."""
        cf = 100
        r = 0.10
        g = 0.02
        
        tv = compute_terminal_value_perpetuity(cf, r, g)
        
        # TV = 100 * (1 + 0.02) / (0.10 - 0.02) = 102 / 0.08 = 1275
        assert tv == pytest.approx(1275, abs=1e-2)
    
    def test_perpetuity_zero_growth(self):
        """Test perpetuity with zero growth."""
        cf = 100
        r = 0.10
        g = 0.0
        
        tv = compute_terminal_value_perpetuity(cf, r, g)
        
        # TV = 100 / 0.10 = 1000
        assert tv == pytest.approx(1000, abs=1e-2)
    
    def test_perpetuity_growth_exceeds_rate(self):
        """Test that growth >= rate raises error."""
        with pytest.raises(TerminalValueError):
            compute_terminal_value_perpetuity(100, 0.10, 0.10)
        
        with pytest.raises(TerminalValueError):
            compute_terminal_value_perpetuity(100, 0.10, 0.15)
    
    def test_exit_multiple(self):
        """Test exit multiple terminal value."""
        tv = compute_terminal_value_exit_multiple(500, 8.0)
        assert tv == pytest.approx(4000, abs=1e-2)


# ============================================================================
# Valuation tests
# ============================================================================

class TestValuation:
    def test_enterprise_value(self):
        """Test EV = Sum PV + PV(TV)."""
        ev = compute_enterprise_value(1000, 5000)
        assert ev == pytest.approx(6000, abs=1e-2)
    
    def test_net_debt(self):
        """Test NetDebt = Debt - Cash."""
        nd = compute_net_debt(1500, 1492.10)
        assert nd == pytest.approx(7.90, abs=1e-2)
    
    def test_equity_from_ev(self):
        """Test Equity = EV - NetDebt."""
        equity = compute_equity_from_ev(10000, 500)
        assert equity == pytest.approx(9500, abs=1e-2)
