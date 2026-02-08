"""
Inputs Recap Module

Renders technical, economic, and derived quantity summaries for the Inputs Recap section.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from dcf_projects.biometano.schema import (
    BiometanoCase,
    DEFAULT_ZES_CREDIT_RATE,
    DEFAULT_PNRR_GRANT_RATE,
)


@dataclass
class InputsRecapData:
    """Structured data for inputs recap rendering."""
    # Technical
    throughput_tpy: float
    impurity_rate: float
    net_throughput_tpy: float
    biogas_yield_smc_t: float
    ch4_percent: float
    upgrading_efficiency: float
    kwh_per_smc: float
    biomethane_mwh_y: float
    compost_tpy: float
    co2_tpy: float
    
    # Economic - Revenue channels
    gate_fee_price: float
    gate_fee_delay: int
    tariff_price: float
    tariff_delay: int
    go_price: float
    go_delay: int
    co2_price: float
    co2_delay: int
    compost_price: float
    compost_delay: int
    
    # Economic - Costs and rates
    tax_rate: float
    pnrr_grant_percent: float
    zes_credit_percent: float
    
    # Derived quantities
    gate_fee_revenue_steady: float
    tariff_revenue_steady: float
    go_revenue_steady: float
    co2_revenue_steady: float
    compost_revenue_steady: float
    total_revenue_steady: float
    total_capex: float
    eligible_capex: float
    pnrr_grant_amount: float
    zes_credit_amount: float


def extract_inputs_recap_data(case: BiometanoCase) -> InputsRecapData:
    """Extract all inputs recap data from case."""
    prod = case.production
    rev = case.revenues
    fin = case.financing
    capex = case.capex
    incent = case.incentives
    
    # Technical parameters
    throughput = prod.forsu_throughput_tpy
    impurity_rate = getattr(prod, 'impurity_rate', 0.20)  # Default 20% if not specified
    net_throughput = throughput * (1 - impurity_rate)
    
    # Biogas/biomethane parameters
    kwh_per_smc = prod.kwh_per_smc
    biomethane_mwh = prod.get_biomethane_mwh()
    biogas_yield = getattr(prod, 'biogas_yield_smc_t', 0)  # Smc/t net
    ch4_percent = getattr(prod, 'ch4_percent', 0.55)  # Default 55%
    upgrading_eff = getattr(prod, 'upgrading_efficiency', 0.98)  # Default 98%
    
    # Byproducts
    compost_tpy = prod.compost_tpy
    co2_tpy = prod.co2_tpy
    
    # Revenue channels
    steady_availability = prod.availability_profile[-1] if prod.availability_profile else 0.95
    
    gate_fee_rev = throughput * steady_availability * rev.gate_fee.price if rev.gate_fee.enabled else 0
    tariff_rev = biomethane_mwh * steady_availability * rev.tariff.price if rev.tariff.enabled else 0
    go_rev = biomethane_mwh * steady_availability * rev.go.price if rev.go.enabled else 0
    co2_rev = co2_tpy * steady_availability * rev.co2.price if rev.co2.enabled else 0
    compost_rev = compost_tpy * steady_availability * rev.compost.price if rev.compost.enabled else 0
    
    total_rev = gate_fee_rev + tariff_rev + go_rev + co2_rev + compost_rev
    
    # CAPEX and incentives
    total_capex = capex.total_capex()
    eligible_capex = capex.eligible_for_grant()
    
    pnrr_percent = incent.capital_grant.percent_of_eligible or DEFAULT_PNRR_GRANT_RATE
    zes_percent = incent.tax_credit.percent_of_eligible or DEFAULT_ZES_CREDIT_RATE
    
    pnrr_amount = eligible_capex * pnrr_percent if incent.capital_grant.enabled else 0
    # ZES is based on TOTAL CAPEX, not eligible
    zes_amount = total_capex * zes_percent if incent.tax_credit.enabled else 0
    
    return InputsRecapData(
        throughput_tpy=throughput,
        impurity_rate=impurity_rate,
        net_throughput_tpy=net_throughput,
        biogas_yield_smc_t=biogas_yield,
        ch4_percent=ch4_percent,
        upgrading_efficiency=upgrading_eff,
        kwh_per_smc=kwh_per_smc,
        biomethane_mwh_y=biomethane_mwh,
        compost_tpy=compost_tpy,
        co2_tpy=co2_tpy,
        gate_fee_price=rev.gate_fee.price,
        gate_fee_delay=rev.gate_fee.payment_delay_days,
        tariff_price=rev.tariff.price,
        tariff_delay=rev.tariff.payment_delay_days,
        go_price=rev.go.price,
        go_delay=rev.go.payment_delay_days,
        co2_price=rev.co2.price,
        co2_delay=rev.co2.payment_delay_days,
        compost_price=rev.compost.price,
        compost_delay=rev.compost.payment_delay_days,
        tax_rate=fin.tax_rate,
        pnrr_grant_percent=pnrr_percent,
        zes_credit_percent=zes_percent,
        gate_fee_revenue_steady=gate_fee_rev,
        tariff_revenue_steady=tariff_rev,
        go_revenue_steady=go_rev,
        co2_revenue_steady=co2_rev,
        compost_revenue_steady=compost_rev,
        total_revenue_steady=total_rev,
        total_capex=total_capex,
        eligible_capex=eligible_capex,
        pnrr_grant_amount=pnrr_amount,
        zes_credit_amount=zes_amount,
    )


def display_inputs_recap(case: BiometanoCase, console: Optional[Console] = None) -> None:
    """Display the Inputs Recap section with all tables."""
    if console is None:
        console = Console()
    
    data = extract_inputs_recap_data(case)
    
    console.print("\n[bold cyan]📋 INPUTS RECAP[/bold cyan]\n")
    
    # Table 1: Technical & Process Assumptions
    tech_table = Table(title="Technical & Process Assumptions", show_header=True)
    tech_table.add_column("Parameter", style="dim")
    tech_table.add_column("Value", justify="right")
    tech_table.add_column("Unit", style="dim")
    
    tech_table.add_row("FORSU Throughput", f"{data.throughput_tpy:,.0f}", "t/y")
    tech_table.add_row("Impurity Rate (Sovvalli)", f"{data.impurity_rate:.1%}", "")
    tech_table.add_row("Net Throughput", f"{data.net_throughput_tpy:,.0f}", "t/y")
    tech_table.add_row("─" * 20, "─" * 10, "")
    tech_table.add_row("Biomethane Output", f"{data.biomethane_mwh_y:,.0f}", "MWh/y")
    tech_table.add_row("kWh/Smc Conversion (PCS)", f"{data.kwh_per_smc:.1f}", "kWh/Smc")
    if data.biogas_yield_smc_t > 0:
        tech_table.add_row("Biogas Yield", f"{data.biogas_yield_smc_t:,.0f}", "Smc/t net")
    if data.ch4_percent > 0:
        tech_table.add_row("CH₄ Content", f"{data.ch4_percent:.1%}", "")
    if data.upgrading_efficiency > 0:
        tech_table.add_row("Upgrading Efficiency", f"{data.upgrading_efficiency:.1%}", "")
    tech_table.add_row("─" * 20, "─" * 10, "")
    tech_table.add_row("CO₂ Output", f"{data.co2_tpy:,.0f}", "t/y")
    tech_table.add_row("Compost Output", f"{data.compost_tpy:,.0f}", "t/y")
    
    console.print(tech_table)
    console.print()
    
    # Table 2: Economic & Contractual Assumptions
    econ_table = Table(title="Economic & Contractual Assumptions", show_header=True)
    econ_table.add_column("Revenue Channel", style="dim")
    econ_table.add_column("Unit Price", justify="right")
    econ_table.add_column("Payment Delay", justify="right")
    
    econ_table.add_row("Gate Fee", f"€{data.gate_fee_price:,.2f}/t", f"{data.gate_fee_delay} days")
    econ_table.add_row("Incentive Tariff (GSE)", f"€{data.tariff_price:,.2f}/MWh", f"{data.tariff_delay} days")
    econ_table.add_row("Garanzie d'Origine (GO)", f"€{data.go_price:,.2f}/MWh", f"{data.go_delay} days")
    econ_table.add_row("CO₂ Sales", f"€{data.co2_price:,.2f}/t", f"{data.co2_delay} days")
    econ_table.add_row("Compost Sales", f"€{data.compost_price:,.2f}/t", f"{data.compost_delay} days")
    
    console.print(econ_table)
    console.print()
