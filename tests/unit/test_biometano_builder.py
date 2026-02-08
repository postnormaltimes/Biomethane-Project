"""
Unit Tests for Biometano Builder
"""
import pytest

from dcf_projects.biometano.schema import BiometanoCase
from dcf_projects.biometano.builder import (
    BiometanoBuilder,
    BiometanoProjections,
    build_projections,
)


@pytest.fixture
def sample_case_data():
    """Sample case data for testing."""
    return {
        "horizon": {
            "base_year": 2024,
            "years_forecast": 5,
            "construction_years": 1,
        },
        "production": {
            "forsu_throughput_tpy": 60000,
            "biomethane_mwh_y": 40000,
            "availability_profile": [0.75, 0.90, 0.95],
            "compost_tpy": 12000,
            "co2_tpy": 4500,
        },
        "revenues": {
            "gate_fee": {"price": 190.0, "payment_delay_days": 45},
            "tariff": {"price": 70.0, "payment_delay_days": 90},
            "co2": {"price": 120.0},
            "go": {"price": 0.30},
            "compost": {"price": 5.0},
        },
        "opex": {
            "feedstock_handling": {"fixed_annual": 500000, "variable_per_tonne": 5.0},
            "utilities": {"fixed_annual": 300000},
            "maintenance": {"percent_of_capex": 0.02},
            "personnel": {"fixed_annual": 800000},
        },
        "capex": {
            "epc": {"amount": 20000000, "spend_profile": [1.0], "useful_life_years": 20},
            "civils": {"amount": 5000000},
        },
        "financing": {
            "debt_amount": 15000000,
            "debt_drawdown_profile": [1.0],
            "cost_of_debt": 0.05,
            "debt_repayment_years": 10,
            "cash_at_base": 2000000,
            "equity_book_at_base": 10000000,
            "tax_rate": 0.24,
            "rf": 0.03,
            "rm": 0.08,
            "beta": 1.2,
        },
        "incentives": {},
        "terminal_value": {"method": "perpetuity", "perpetuity_growth": 0.0},
    }


@pytest.fixture
def sample_case(sample_case_data):
    return BiometanoCase.model_validate(sample_case_data)


class TestBiometanoBuilder:
    """Tests for BiometanoBuilder."""
    
    def test_build_returns_projections(self, sample_case):
        builder = BiometanoBuilder(sample_case)
        proj = builder.build()
        
        assert isinstance(proj, BiometanoProjections)
        assert proj.base_year == 2024
        assert proj.cod_year == 2026
    
    def test_production_ramp_up(self, sample_case):
        proj = build_projections(sample_case)
        
        # Operating years (years_forecast = 5 operating years)
        op_years = proj.operating_years
        assert len(op_years) == 5  # 5 operating years from COD

        
        # First operating year
        prod_y1 = next(p for p in proj.production if p.year == 2026)
        assert prod_y1.availability == 0.75
        assert prod_y1.forsu_tonnes == 60000 * 0.75
        
        # Second year
        prod_y2 = next(p for p in proj.production if p.year == 2027)
        assert prod_y2.availability == 0.90
    
    def test_revenue_calculation(self, sample_case):
        proj = build_projections(sample_case)
        
        # First operating year
        rev = proj.get_revenue(2026)
        assert rev is not None
        
        # Gate fee = 60000 * 0.75 * 190 = 8,550,000
        expected_gate_fee = 60000 * 0.75 * 190
        assert abs(rev.gate_fee - expected_gate_fee) < 1
        
        # Tariff = 40000 * 0.75 * 70 = 2,100,000
        expected_tariff = 40000 * 0.75 * 70
        assert abs(rev.tariff - expected_tariff) < 1
    
    def test_capex_in_construction_year(self, sample_case):
        proj = build_projections(sample_case)
        
        capex_2025 = proj.get_capex(2025)
        assert capex_2025 is not None
        assert capex_2025.total == 25000000  # 20M EPC + 5M civils
    
    def test_ebitda_calculation(self, sample_case):
        proj = build_projections(sample_case)
        
        # EBITDA in first operating year
        ebitda_2026 = proj.ebitda.get(2026)
        assert ebitda_2026 is not None
        
        rev = proj.get_revenue(2026)
        opex = proj.get_opex(2026)
        expected_ebitda = rev.total - opex.total
        
        assert abs(ebitda_2026 - expected_ebitda) < 1
    
    def test_fcff_calculation(self, sample_case):
        proj = build_projections(sample_case)
        
        # FCFF should exist for operating years
        for year in proj.operating_years:
            assert year in proj.fcff
            # FCFF = NOPAT + D&A - ΔNWC - CAPEX
    
    def test_fcfe_calculation(self, sample_case):
        proj = build_projections(sample_case)
        
        # FCFE should exist for operating years
        for year in proj.operating_years:
            assert year in proj.fcfe


class TestBuildProjectionsConvenience:
    """Tests for build_projections convenience function."""
    
    def test_convenience_function(self, sample_case):
        proj = build_projections(sample_case)
        
        assert isinstance(proj, BiometanoProjections)
        assert len(proj.revenues) > 0
        assert len(proj.opex) > 0


class TestWithIncentives:
    """Tests with incentives enabled."""
    
    @pytest.fixture
    def case_with_grant(self, sample_case_data):
        sample_case_data["incentives"] = {
            "capital_grant": {
                "enabled": True,
                "percent_of_eligible": 0.40,
                "accounting_policy": "A2",
                "cash_receipt_schedule": [0.5, 0.5],
            },
        }
        return BiometanoCase.model_validate(sample_case_data)
    
    def test_grant_in_accounting(self, case_with_grant):
        proj = build_projections(case_with_grant)
        
        assert proj.accounting is not None
        assert proj.accounting.total_grant_amount > 0
        # Grant = 40% of 25M = 10M
        expected_grant = 25000000 * 0.40
        assert abs(proj.accounting.total_grant_amount - expected_grant) < 1
    
    def test_deferred_income_schedule(self, case_with_grant):
        proj = build_projections(case_with_grant)
        
        # Should have deferred income entries
        assert len(proj.accounting.deferred_income) > 0
        
        # First operating year should have release
        cod_year = proj.cod_year
        di = proj.accounting.get_deferred_income(cod_year)
        assert di is not None
