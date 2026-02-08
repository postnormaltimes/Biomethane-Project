"""
Biometano Project Finance Module

Specialized module for biometano plant project finance modeling with:
- Construction and operating phase modeling
- Multi-channel revenue (gate fee, tariff, CO2, GO, compost)
- OIC-compliant incentive accounting (grants, tax credits)
- DCF valuation integration
"""
from dcf_projects.biometano.schema import BiometanoCase
from dcf_projects.biometano.builder import build_projections, BiometanoProjections
from dcf_projects.biometano.accounting import compute_accounting, AccountingOutputs
from dcf_projects.biometano.statements import build_statements, FinancialStatements
from dcf_projects.biometano.valuation import compute_valuation, ValuationOutputs
from dcf_projects.biometano.sensitivities import run_sensitivity_analysis, SensitivityAnalysisOutputs

__all__ = [
    "BiometanoCase",
    "build_projections",
    "BiometanoProjections",
    "compute_accounting",
    "AccountingOutputs",
    "build_statements",
    "FinancialStatements",
    "compute_valuation",
    "ValuationOutputs",
    "run_sensitivity_analysis",
    "SensitivityAnalysisOutputs",
]
