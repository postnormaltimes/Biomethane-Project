"""
Unit Tests for Biometano Schema
"""
import pytest
from pydantic import ValidationError

from dcf_projects.biometano.schema import (
    BiometanoCase,
    HorizonInputs,
    ProductionInputs,
    RevenueChannelInputs,
    RevenuesInputs,
    OpexCategoryInputs,
    OpexInputs,
    CapexLineItem,
    CapexInputs,
    FinancingInputs,
    CapitalGrantInputs,
    TaxCreditInputs,
    IncentivesInputs,
    TerminalValueInputs,
    GrantAccountingPolicy,
    TaxCreditPolicy,
    GrantRecognitionTrigger,
)


class TestHorizon:
    """Tests for Horizon model."""
    
    def test_horizon_basic(self):
        h = HorizonInputs(base_year=2024, years_forecast=15, construction_years=2)
        assert h.base_year == 2024
        assert h.cod_year == 2027
        assert len(h.construction_years_list) == 2
        assert len(h.operating_years_list) == 15
    
    def test_horizon_no_construction(self):
        h = HorizonInputs(base_year=2024, years_forecast=10, construction_years=0)
        assert h.cod_year == 2025
        assert len(h.construction_years_list) == 0
    
    def test_horizon_all_forecast_years(self):
        h = HorizonInputs(base_year=2024, years_forecast=5, construction_years=1)
        # years_forecast applies to operating years; construction years are separate
        # Total forecast = 1 construction + 5 operating = 6 years
        assert len(h.all_forecast_years) == 6  # [2025, 2026, 2027, 2028, 2029, 2030]


class TestProduction:
    """Tests for Production model."""
    
    def test_production_with_smc(self):
        p = ProductionInputs(
            forsu_throughput_tpy=60000,
            biomethane_smc_y=4000000,
            kwh_per_smc=10,
        )
        assert p.get_biomethane_mwh() == 40000
    
    def test_production_with_mwh(self):
        p = ProductionInputs(
            forsu_throughput_tpy=60000,
            biomethane_mwh_y=35000,
        )
        assert p.get_biomethane_mwh() == 35000
    
    def test_availability_profile(self):
        p = ProductionInputs(
            forsu_throughput_tpy=60000,
            biomethane_mwh_y=40000,
            availability_profile=[0.75, 0.90, 0.95],
        )
        assert p.get_availability(0) == 0.75
        assert p.get_availability(1) == 0.90
        assert p.get_availability(10) == 0.95  # Last value


class TestRevenueChannel:
    """Tests for RevenueChannel model."""
    
    def test_channel_enabled_by_default_with_price(self):
        rc = RevenueChannelInputs(price=100.0)
        assert rc.enabled is True
    
    def test_channel_default_escalation(self):
        rc = RevenueChannelInputs(price=100.0)
        assert rc.escalation_rate == 0.0


class TestCapexItem:
    """Tests for CapexItem model."""
    
    def test_capex_item_defaults(self):
        ci = CapexLineItem(amount=1000000)
        assert ci.spend_profile == [1.0]
        assert ci.useful_life_years == 20
        assert ci.capitalize is True
    
    def test_capex_item_spend_profile_validation(self):
        ci = CapexLineItem(
            amount=1000000,
            spend_profile=[0.4, 0.6],
        )
        assert sum(ci.spend_profile) == 1.0


class TestCapexInputs:
    """Tests for CapexInputs model."""
    
    def test_total_capex(self):
        capex = CapexInputs(
            epc=CapexLineItem(amount=10000000),
            civils=CapexLineItem(amount=5000000),
            upgrading_unit=CapexLineItem(amount=3000000),
        )
        assert capex.total_capex() == 18000000
    
    def test_eligible_for_grant(self):
        capex = CapexInputs(
            epc=CapexLineItem(amount=10000000, eligible_for_grant=True),
            civils=CapexLineItem(amount=5000000, eligible_for_grant=False),
        )
        assert capex.eligible_for_grant() == 10000000


class TestIncentives:
    """Tests for incentive models."""
    
    def test_grant_accounting_policy_enum(self):
        assert GrantAccountingPolicy.REDUCE_ASSET.value == "A1"
        assert GrantAccountingPolicy.DEFERRED_INCOME.value == "A2"
    
    def test_tax_credit_policy_enum(self):
        assert TaxCreditPolicy.REDUCE_TAX_EXPENSE.value == "B1"
        assert TaxCreditPolicy.TAX_RECEIVABLE.value == "B2"
    
    def test_capital_grant_defaults(self):
        grant = CapitalGrantInputs(enabled=True, percent_of_eligible=0.40)
        assert grant.accounting_policy == GrantAccountingPolicy.DEFERRED_INCOME


class TestBiometanoCase:
    """Tests for complete case model."""
    
    @pytest.fixture
    def minimal_case_data(self):
        return {
            "horizon": {
                "base_year": 2024,
                "years_forecast": 10,
                "construction_years": 2,
            },
            "production": {
                "forsu_throughput_tpy": 60000,
                "biomethane_mwh_y": 40000,
            },
            "revenues": {
                "gate_fee": {"price": 190},
                "tariff": {"price": 70},
            },
            "opex": {},
            "capex": {
                "epc": {"amount": 20000000},
            },
            "financing": {
                "tax_rate": 0.24,
                "rf": 0.03,
                "rm": 0.08,
                "beta": 1.2,
            },
            "incentives": {},
            "terminal_value": {
                "method": "perpetuity",
                "perpetuity_growth": 0.0,
            },
        }
    
    def test_minimal_case_loads(self, minimal_case_data):
        case = BiometanoCase.model_validate(minimal_case_data)
        assert case.horizon.base_year == 2024
        assert case.production.forsu_throughput_tpy == 60000
        assert case.capex.total_capex() == 20000000
    
    def test_case_with_incentives(self, minimal_case_data):
        minimal_case_data["incentives"] = {
            "capital_grant": {
                "enabled": True,
                "percent_of_eligible": 0.40,
                "accounting_policy": "A2",
            },
            "tax_credit": {
                "enabled": True,
                "percent_of_eligible": 0.10,
                "accounting_policy": "B1",
            },
        }
        case = BiometanoCase.model_validate(minimal_case_data)
        assert case.incentives.capital_grant.enabled is True
        assert case.incentives.capital_grant.percent_of_eligible == 0.40
        assert case.incentives.tax_credit.enabled is True
