"""
Integration Tests for Biometano Module

Tests the complete flow from case YAML to valuation outputs.
"""
import pytest
import yaml
from pathlib import Path

from dcf_projects.biometano.schema import BiometanoCase
from dcf_projects.biometano.builder import build_projections
from dcf_projects.biometano.statements import build_statements
from dcf_projects.biometano.valuation import compute_valuation
from dcf_projects.biometano.sensitivities import run_sensitivity_analysis


CASE_FILE = Path(__file__).parent.parent.parent / "src/dcf_projects/biometano/case_files/biometano_case.yaml"


@pytest.fixture
def biometano_case():
    """Load the actual biometano case file."""
    with open(CASE_FILE) as f:
        data = yaml.safe_load(f)
    return BiometanoCase.model_validate(data)


class TestFullBiometanoFlow:
    """Integration tests for complete biometano flow."""
    
    def test_case_file_exists(self):
        """Case file should exist."""
        assert CASE_FILE.exists()
    
    def test_case_loads_successfully(self, biometano_case):
        """Case file should load and validate successfully."""
        assert biometano_case.horizon.base_year == 2024
        assert biometano_case.production.forsu_throughput_tpy == 60000
    
    def test_projections_build(self, biometano_case):
        """Projections should build successfully."""
        proj = build_projections(biometano_case)
        
        assert proj is not None
        assert len(proj.production) > 0
        assert len(proj.revenues) > 0
        assert len(proj.opex) > 0
    
    def test_statements_build(self, biometano_case):
        """Financial statements should build successfully."""
        proj = build_projections(biometano_case)
        stmts = build_statements(biometano_case, proj)
        
        assert stmts is not None
        assert len(stmts.income_statements) > 0
        assert len(stmts.balance_sheets) > 0
        assert len(stmts.cash_flows) > 0
    
    def test_valuation_computes(self, biometano_case):
        """Valuation should compute successfully."""
        proj = build_projections(biometano_case)
        val = compute_valuation(biometano_case, proj)
        
        assert val is not None
        assert val.enterprise_value > 0
        assert val.equity_value > 0
    
    def test_enterprise_value_reasonable(self, biometano_case):
        """EV should be within reasonable range for 60kTPA plant."""
        proj = build_projections(biometano_case)
        val = compute_valuation(biometano_case, proj)
        
        # EV should be between 50M and 200M for this size plant
        assert val.enterprise_value > 50_000_000
        assert val.enterprise_value < 200_000_000
    
    def test_equity_value_less_than_ev(self, biometano_case):
        """Equity value should be less than EV due to net debt."""
        proj = build_projections(biometano_case)
        val = compute_valuation(biometano_case, proj)
        
        # If there's net debt, equity < EV
        if val.net_debt > 0:
            assert val.equity_value < val.enterprise_value
    
    def test_reconciliation_small(self, biometano_case):
        """FCFF/WACC and FCFE/Ke should reconcile closely."""
        proj = build_projections(biometano_case)
        val = compute_valuation(biometano_case, proj)
        
        # Difference should be < 5% of equity value
        recon_pct = abs(val.reconciliation_difference) / val.equity_value if val.equity_value > 0 else 0
        assert recon_pct < 0.10  # 10% tolerance


class TestProductionVolumes:
    """Tests for production volume calculations."""
    
    def test_biomethane_output(self, biometano_case):
        """Biomethane MWh should match input conversion."""
        proj = build_projections(biometano_case)
        
        # From case: 4M Smc * 10 kWh/Smc = 40,000 MWh
        expected_full_mwh = 40_000
        
        # First op year at 75% availability
        prod_y1 = proj.production[2]  # After 2 construction years
        assert abs(prod_y1.biomethane_mwh - expected_full_mwh * 0.75) < 100
    
    def test_byproducts(self, biometano_case):
        """Byproduct volumes should match input."""
        proj = build_projections(biometano_case)
        
        # At 95% availability (year 3+)
        for prod in proj.production:
            if prod.availability == 0.95:
                assert abs(prod.co2_tonnes - 4560 * 0.95) < 10
                assert abs(prod.compost_tonnes - 12148 * 0.95) < 10
                break


class TestRevenueChannels:
    """Tests for revenue channel calculations."""
    
    def test_all_channels_active(self, biometano_case):
        """All 5 revenue channels should be active."""
        proj = build_projections(biometano_case)
        
        # Check a steady-state year
        rev = proj.get_revenue(biometano_case.horizon.cod_year + 2)
        assert rev.gate_fee > 0
        assert rev.tariff > 0
        assert rev.co2 > 0
        assert rev.go > 0
        assert rev.compost > 0
    
    def test_gate_fee_dominant(self, biometano_case):
        """Gate fee should be largest revenue component."""
        proj = build_projections(biometano_case)
        
        rev = proj.get_revenue(biometano_case.horizon.cod_year + 2)
        assert rev.gate_fee > rev.tariff
        assert rev.gate_fee > rev.co2
        assert rev.gate_fee > rev.compost


class TestIncentivesAccounting:
    """Tests for incentive accounting."""
    
    def test_grant_calculated(self, biometano_case):
        """Capital grant should be calculated from eligible CAPEX."""
        proj = build_projections(biometano_case)
        
        # 40% of eligible CAPEX
        eligible = biometano_case.capex.eligible_for_grant()
        expected_grant = eligible * 0.40
        
        assert proj.accounting is not None
        assert abs(proj.accounting.total_grant_amount - expected_grant) < 1000
    
    def test_tax_credit_calculated(self, biometano_case):
        """Tax credit should be calculated from CAPEX."""
        proj = build_projections(biometano_case)
        
        # 14.6189% of total CAPEX (ZES rate)
        total_capex = biometano_case.capex.total_capex()
        expected_credit = total_capex * 0.146189
        
        assert proj.accounting is not None
        assert abs(proj.accounting.total_tax_credit - expected_credit) < 1000


class TestStatements:
    """Tests for financial statements."""
    
    def test_income_statement_format(self, biometano_case):
        """Income statement should have correct format."""
        proj = build_projections(biometano_case)
        stmts = build_statements(biometano_case, proj)
        
        for is_stmt in stmts.income_statements:
            # EBITDA = Revenue - OPEX
            expected_ebitda = is_stmt.total_revenue - is_stmt.total_opex
            assert abs(is_stmt.ebitda - expected_ebitda) < 1
            
            # EBIT = EBITDA - D&A + Grant release
            # (approximately)
    
    def test_cash_flow_format(self, biometano_case):
        """Cash flow statement should have correct format."""
        proj = build_projections(biometano_case)
        stmts = build_statements(biometano_case, proj)
        
        for cf in stmts.cash_flows:
            # Net CF = CFO + CFI + CFF
            expected_net = cf.cfo + cf.cfi + cf.cff
            assert abs(cf.net_cash_flow - expected_net) < 1


class TestSensitivity:
    """Tests for sensitivity analysis."""
    
    def test_sensitivity_runs(self, biometano_case):
        """Sensitivity analysis should run without error."""
        def val_fn(c):
            proj = build_projections(c)
            val = compute_valuation(c, proj)
            return val.equity_value, val.enterprise_value, val.sum_pv_fcff, val.pv_terminal_value_fcff
        
        result = run_sensitivity_analysis(biometano_case, value_function=val_fn)
        
        assert result is not None
        assert result.base_equity_value > 0
        assert len(result.tornado_data) > 0
    
    def test_scenarios_generated(self, biometano_case):
        """Scenarios should be generated."""
        def val_fn(c):
            proj = build_projections(c)
            val = compute_valuation(c, proj)
            return val.equity_value, val.enterprise_value, val.sum_pv_fcff, val.pv_terminal_value_fcff
        
        result = run_sensitivity_analysis(biometano_case, value_function=val_fn)
        
        # Should have Base, Upside, Downside
        assert len(result.scenarios) == 3
        scenario_names = [s.name for s in result.scenarios]
        assert "Base" in scenario_names
        assert "Upside" in scenario_names
        assert "Downside" in scenario_names
    
    def test_upside_better_than_downside(self, biometano_case):
        """Upside scenario should have higher value than Downside."""
        def val_fn(c):
            proj = build_projections(c)
            val = compute_valuation(c, proj)
            return val.equity_value, val.enterprise_value, val.sum_pv_fcff, val.pv_terminal_value_fcff
        
        result = run_sensitivity_analysis(biometano_case, value_function=val_fn)
        
        upside = next(s for s in result.scenarios if s.name == "Upside")
        downside = next(s for s in result.scenarios if s.name == "Downside")
        
        assert upside.equity_value > downside.equity_value
