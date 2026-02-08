"""
Biometano CLI Display Functions

Rich table formatting for Biometano-specific outputs.
Section ordering:
1. Inputs Recap (from inputs_recap.py)
2. Incentives Waterfall (PNRR + ZES)
3. Production Summary
4. Revenue by Channel
5. OPEX Breakdown
6. Financial Statement (Yearly)
7. Balance Sheet (Yearly)
8. FCFF Schedule
9. Valuation Summary (compact)
10. Sensitivity Analysis
11. Scenario Comparison
"""
from __future__ import annotations

from typing import Optional

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box

from dcf_projects.biometano.builder import BiometanoProjections
from dcf_projects.biometano.statements import FinancialStatements
from dcf_projects.biometano.valuation import ValuationOutputs
from dcf_projects.biometano.sensitivities import SensitivityAnalysisOutputs
from dcf_projects.biometano.schema import REVENUE_CHANNEL_ORDER
from dcf_projects.biometano.incentives_allocation import IncentiveAllocationResult


console = Console()


def display_production_summary(projections: BiometanoProjections) -> None:
    """Display production summary table."""
    table = Table(
        title="🌿 Production Summary",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold cyan",
    )
    
    table.add_column("Year", justify="center")
    table.add_column("Availability", justify="right")
    table.add_column("FORSU (t)", justify="right")
    table.add_column("Biomethane (MWh)", justify="right")
    table.add_column("CO₂ (t)", justify="right")
    table.add_column("Compost (t)", justify="right")
    
    for prod in projections.production:
        if prod.year >= projections.cod_year:
            table.add_row(
                str(prod.year),
                f"{prod.availability:.0%}",
                f"{prod.forsu_tonnes:,.0f}",
                f"{prod.biomethane_mwh:,.0f}",
                f"{prod.co2_tonnes:,.0f}",
                f"{prod.compost_tonnes:,.0f}",
            )
    
    console.print(table)
    console.print()


def display_revenue_breakdown(projections: BiometanoProjections) -> None:
    """Display revenue by channel table with correct column order.
    
    Order: Gate Fee, Tariff, GO, CO₂, Compost, Total
    """
    table = Table(
        title="💰 Revenue by Channel",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold green",
    )
    
    # Correct column order per requirements
    table.add_column("Year", justify="center")
    table.add_column("Gate Fee", justify="right")
    table.add_column("Tariff", justify="right")
    table.add_column("GO", justify="right")
    table.add_column("CO₂", justify="right")
    table.add_column("Compost", justify="right")
    table.add_column("Total", justify="right", style="bold")
    
    for rev in projections.revenues:
        if rev.year >= projections.cod_year:
            table.add_row(
                str(rev.year),
                f"{rev.gate_fee:,.0f}",
                f"{rev.tariff:,.0f}",
                f"{rev.go:,.0f}",
                f"{rev.co2:,.0f}",
                f"{rev.compost:,.0f}",
                f"{rev.total:,.0f}",
            )
    
    console.print(table)
    console.print()


def display_opex_breakdown(projections: BiometanoProjections) -> None:
    """Display OPEX by category table."""
    table = Table(
        title="📊 OPEX by Category",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold red",
    )
    
    table.add_column("Year", justify="center")
    table.add_column("Feedstock", justify="right")
    table.add_column("Utilities", justify="right")
    table.add_column("Maintenance", justify="right")
    table.add_column("Personnel", justify="right")
    table.add_column("Other", justify="right")
    table.add_column("Total", justify="right", style="bold")
    
    for opex in projections.opex:
        if opex.year >= projections.cod_year:
            other = opex.chemicals + opex.insurance + opex.overheads + opex.digestate_handling + opex.other
            table.add_row(
                str(opex.year),
                f"{opex.feedstock_handling:,.0f}",
                f"{opex.utilities:,.0f}",
                f"{opex.maintenance:,.0f}",
                f"{opex.personnel:,.0f}",
                f"{other:,.0f}",
                f"{opex.total:,.0f}",
            )
    
    console.print(table)
    console.print()


def display_income_statement(
    projections: BiometanoProjections,
    statements: FinancialStatements,
    tax_rate: float = 0.24,
) -> None:
    """Display Income Statement (Yearly).
    
    Shows P&L items from Revenue to Net Income.
    """
    table = Table(
        title="📋 Income Statement (Yearly)",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold yellow",
    )
    
    table.add_column("Year", justify="center")
    table.add_column("Revenue", justify="right")
    table.add_column("OPEX", justify="right")
    table.add_column("EBITDA", justify="right")
    table.add_column("D&A", justify="right")
    table.add_column("EBIT", justify="right")
    table.add_column("Interest", justify="right")
    table.add_column("EBT", justify="right")
    table.add_column("Taxes", justify="right")
    table.add_column("Net Income", justify="right", style="bold")
    
    for stmt in statements.income_statements:
        if stmt.year >= projections.cod_year:
            table.add_row(
                str(stmt.year),
                f"{stmt.total_revenue:,.0f}",
                f"({stmt.total_opex:,.0f})",
                f"{stmt.ebitda:,.0f}",
                f"({stmt.depreciation:,.0f})",
                f"{stmt.ebit:,.0f}",
                f"({stmt.interest_expense:,.0f})",
                f"{stmt.ebt:,.0f}",
                f"({stmt.taxes_paid:,.0f})",
                f"{stmt.net_income:,.0f}",
            )
    
    console.print(table)
    console.print()


def display_fcff_bridge(
    projections: BiometanoProjections,
    statements: FinancialStatements,
    tax_rate: float = 0.24,
) -> None:
    """Display FCFF Bridge (Yearly) - EBIT to Free Cash Flow.
    
    Shows the financial bridge from EBIT to FCFF per Maccarrone DCF methodology:
    FCFF = EBIT × (1-t) + D&A ± ΔNWC - CAPEX
    """
    table = Table(
        title="💰 Financial Statement / FCFF Bridge (Yearly)",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold cyan",
    )
    
    table.add_column("Year", justify="center")
    table.add_column("EBIT", justify="right")
    table.add_column("Tax @24%", justify="right")
    table.add_column("NOPAT", justify="right")
    table.add_column("+ D&A", justify="right")
    table.add_column("± ΔNWC", justify="right")
    table.add_column("- CAPEX", justify="right")
    table.add_column("FCFF", justify="right", style="bold green")
    
    for stmt in statements.income_statements:
        year = stmt.year
        if year >= projections.cod_year:
            ebit = projections.ebit.get(year, 0)
            tax_on_ebit = ebit * tax_rate if ebit > 0 else 0
            nopat = ebit - tax_on_ebit
            da = projections.depreciation.get(year, 0)
            delta_nwc = projections.delta_nwc.get(year, 0)
            capex = projections.get_capex(year)
            capex_val = capex.total if capex else 0
            fcff = projections.fcff.get(year, 0)
            
            table.add_row(
                str(year),
                f"{ebit:,.0f}",
                f"({tax_on_ebit:,.0f})",
                f"{nopat:,.0f}",
                f"{da:,.0f}",
                f"({delta_nwc:,.0f})" if delta_nwc > 0 else f"{-delta_nwc:,.0f}",
                f"({capex_val:,.0f})",
                f"{fcff:,.0f}",
            )
    
    console.print(table)
    console.print()


def display_balance_sheet_recap(
    statements: FinancialStatements,
    projections: BiometanoProjections,
) -> None:
    """Display Balance Sheet (Yearly recap).
    
    Shows key balance sheet items with balance check.
    """
    table = Table(
        title="📊 Balance Sheet (Yearly)",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold blue",
    )
    
    table.add_column("Year", justify="center")
    table.add_column("Fixed Assets", justify="right")
    table.add_column("Curr Assets", justify="right")
    table.add_column("ZES Credit", justify="right")
    table.add_column("Total Assets", justify="right", style="bold")
    table.add_column("Debt", justify="right")
    table.add_column("Def. Income", justify="right")
    table.add_column("Equity", justify="right")
    table.add_column("Check", justify="center")
    
    for bs in statements.balance_sheets:
        # Use actual attribute names from BalanceSheetLine
        balanced = abs(bs.balance_check) < 1
        check_mark = "✓" if balanced else "⚠"
        check_style = "green" if balanced else "red"
        
        table.add_row(
            str(bs.year),
            f"{bs.fixed_assets_net:,.0f}",
            f"{bs.total_current_assets:,.0f}",
            f"{bs.tax_credit_receivable:,.0f}",
            f"{bs.total_assets:,.0f}",
            f"{bs.debt:,.0f}",
            f"{bs.deferred_income:,.0f}",
            f"{bs.total_equity:,.0f}",
            f"[{check_style}]{check_mark}[/{check_style}]",
        )
    
    console.print(table)
    console.print()


def display_fcff_schedule(projections: BiometanoProjections, tax_rate: float = 0.24) -> None:
    """Display FCFF breakdown."""
    table = Table(
        title="🔄 Free Cash Flow to Firm (FCFF)",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold magenta",
    )
    
    table.add_column("Year", justify="center")
    table.add_column("EBIT", justify="right")
    table.add_column("Tax on EBIT", justify="right")
    table.add_column("NOPAT", justify="right")
    table.add_column("+D&A", justify="right")
    table.add_column("-ΔNWC", justify="right")
    table.add_column("-CAPEX", justify="right")
    table.add_column("FCFF", justify="right", style="bold")
    
    for year in projections.operating_years:
        ebit = projections.ebit.get(year, 0.0)
        tax = ebit * tax_rate if ebit > 0 else 0.0
        nopat = ebit - tax
        da = projections.depreciation.get(year, 0.0)
        delta_nwc = projections.delta_nwc.get(year, 0.0)
        capex = projections.get_capex(year)
        capex_total = capex.total if capex else 0.0
        fcff = projections.fcff.get(year, 0.0)
        
        table.add_row(
            str(year),
            f"{ebit:,.0f}",
            f"({tax:,.0f})",
            f"{nopat:,.0f}",
            f"{da:,.0f}",
            f"({delta_nwc:,.0f})",
            f"({capex_total:,.0f})",
            f"{fcff:,.0f}",
        )
    
    console.print(table)
    console.print()


def display_valuation_summary(valuation: ValuationOutputs, methodology: str = "enterprise") -> None:
    """Display compact Valuation Summary - replaces Valuation Bridge.
    
    No bridge waterfall, no commentary - just core valuation metrics.
    """
    table = Table(
        title="🎯 Valuation Summary (FCFF/WACC)",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold cyan",
    )
    
    table.add_column("Metric", justify="left")
    table.add_column("Value", justify="right")
    
    table.add_row("Sum PV(FCFF)", f"€{valuation.sum_pv_fcff:,.0f}")
    table.add_row("Terminal Value", f"€{valuation.terminal_value_fcff:,.0f}")
    table.add_row("PV(Terminal Value)", f"€{valuation.pv_terminal_value_fcff:,.0f}")
    table.add_row("", "")
    table.add_row("[bold]Enterprise Value (FCFF/WACC)[/bold]", f"[bold]€{valuation.enterprise_value:,.0f}[/bold]")
    
    # TV share
    tv_share = valuation.pv_terminal_value_fcff / valuation.enterprise_value * 100 if valuation.enterprise_value > 0 else 0
    table.add_row("TV Share of EV", f"{tv_share:.1f}%")
    
    # Only show equity if explicitly requested
    if methodology in ("equity", "both"):
        table.add_row("", "")
        table.add_row("Net Debt", f"€{valuation.net_debt:,.0f}")
        table.add_row("[bold]Equity Value[/bold]", f"[bold]€{valuation.equity_value:,.0f}[/bold]")
    
    console.print(table)
    console.print()


def display_sensitivity_tornado(sensitivity: SensitivityAnalysisOutputs, methodology: str = "enterprise") -> None:
    """Display sensitivity tornado table for Enterprise Value."""
    value_label = "Enterprise Value" if methodology == "enterprise" else "Equity Value"
    
    table = Table(
        title=f"🌪️ Sensitivity Analysis ({value_label})",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold yellow",
    )
    
    table.add_column("Parameter", justify="left")
    table.add_column("Low Shock", justify="center")
    table.add_column("Low EV", justify="right")
    table.add_column("High Shock", justify="center")
    table.add_column("High EV", justify="right")
    table.add_column("Spread", justify="right", style="bold")
    
    for t in sensitivity.tornado_data[:12]:  # Top 12
        # Use EV values for tornado display
        low_val = t.low_ev if methodology == "enterprise" else t.low_value
        high_val = t.high_ev if methodology == "enterprise" else t.high_value
        spread = abs(high_val - low_val)
        
        table.add_row(
            t.parameter,
            t.low_label,
            f"€{low_val:,.0f}",
            t.high_label,
            f"€{high_val:,.0f}",
            f"€{spread:,.0f}",
        )
    
    console.print(table)
    console.print()
    
    # Base value
    base_val = sensitivity.base_ev if methodology == "enterprise" else sensitivity.base_equity_value
    console.print(f"[dim]Base {value_label}: €{base_val:,.0f}[/dim]")
    console.print()


def display_scenario_comparison(sensitivity: SensitivityAnalysisOutputs, methodology: str = "enterprise") -> None:
    """Display scenario comparison table for EV."""
    value_label = "Enterprise Value" if methodology == "enterprise" else "Equity Value"
    
    table = Table(
        title=f"📊 Scenario Comparison ({value_label})",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold green",
    )
    
    table.add_column("Scenario", justify="left")
    table.add_column(value_label, justify="right")
    table.add_column("PV(FCFF)", justify="right")
    table.add_column("PV(TV)", justify="right")
    table.add_column("TV Share", justify="right")
    table.add_column("Delta", justify="right")
    table.add_column("Delta %", justify="right")
    
    for s in sensitivity.scenarios:
        val = s.ev if methodology == "enterprise" else s.equity_value
        delta = s.delta_ev if methodology == "enterprise" else s.delta_from_base
        delta_pct = s.delta_ev_pct if methodology == "enterprise" else s.delta_pct
        
        delta_style = "green" if delta >= 0 else "red"
        
        pv_fcff = getattr(s, 'pv_fcff', 0)
        pv_tv = getattr(s, 'pv_tv', 0)
        tv_share = pv_tv / val * 100 if val > 0 else 0
        
        table.add_row(
            s.name,
            f"€{val:,.0f}",
            f"€{pv_fcff:,.0f}",
            f"€{pv_tv:,.0f}",
            f"{tv_share:.1f}%",
            f"[{delta_style}]{delta:+,.0f}[/{delta_style}]",
            f"[{delta_style}]{delta_pct:+.1%}[/{delta_style}]",
        )
    
    console.print(table)
    console.print()


def display_incentives_summary(projections: BiometanoProjections) -> None:
    """Display incentives summary."""
    if not projections.accounting:
        return
    
    acc = projections.accounting
    
    table = Table(
        title="🎁 Incentives Summary",
        box=box.ROUNDED,
        show_header=True,
        header_style="bold green",
    )
    
    table.add_column("Incentive", justify="left")
    table.add_column("Amount", justify="right")
    table.add_column("Status", justify="center")
    
    if acc.total_grant_amount > 0:
        table.add_row("PNRR Grant", f"€{acc.total_grant_amount:,.0f}", "✓ Active")
    
    if acc.total_tax_credit > 0:
        table.add_row("ZES Tax Credit", f"€{acc.total_tax_credit:,.0f}", "✓ Active")
    
    if acc.total_grant_amount > 0 or acc.total_tax_credit > 0:
        console.print(table)
        console.print()


def display_incentives_waterfall(result: IncentiveAllocationResult) -> None:
    """Display Incentives Waterfall Allocation & Riparto section.
    
    5 tables:
    1. PNRR eligibility and grant computation
    2. ESL cap and ZES gap
    3. ZES base allocation (line-by-line)
    4. Riparto stress table
    5. Compliance check
    """
    console.print()
    console.rule("[bold green]📋 Incentives — Waterfall Allocation & Riparto[/bold green]")
    console.print()
    
    # Table 1: PNRR Eligibility & Grant
    t1 = Table(title="PNRR Eligibility & Grant", box=box.ROUNDED, header_style="bold green")
    t1.add_column("Parameter", style="cyan")
    t1.add_column("Value", justify="right")
    
    t1.add_row("Biomethane Smc/h", f"{result.smc_per_hour:,.0f}")
    t1.add_row("Annual Hours", "8,000")
    t1.add_row("CS_max Base", f"€{result.cs_max_base:,.0f}/Smc/h")
    t1.add_row("Inflation Factor", "1.137")
    t1.add_row("CS_max Adjusted", f"€{result.cs_max_adjusted:,.0f}/Smc/h")
    t1.add_row("─" * 20, "─" * 15)
    t1.add_row("Eligible Spend (PNRR)", f"€{result.eligible_spend_pnrr:,.0f}")
    t1.add_row("Grant Rate", f"{result.grant_rate:.0%}")
    t1.add_row("[bold]Grant Amount[/bold]", f"[bold]€{result.grant_pnrr:,.0f}[/bold]")
    t1.add_row("─" * 20, "─" * 15)
    t1.add_row("Tech Costs Limit (12%)", f"€{result.tech_costs_limit:,.0f}")
    t1.add_row("Tech Costs Actual", f"€{result.tech_costs_actual:,.0f}")
    tech_status = "[green]✓ OK[/green]" if not result.tech_costs_warning else "[red]⚠ Over Limit[/red]"
    t1.add_row("Tech Cost Check", tech_status)
    
    console.print(t1)
    console.print()
    
    # Table 2: ESL Cap & ZES Gap
    t2 = Table(title="ESL Cap & ZES Gap", box=box.ROUNDED, header_style="bold green")
    t2.add_column("Item", style="cyan")
    t2.add_column("Value", justify="right")
    
    t2.add_row("Total CAPEX", f"€{result.total_capex:,.0f}")
    t2.add_row("Max ESL Intensity", f"{result.max_esl_intensity:.0%}")
    t2.add_row("Max Aid Amount", f"€{result.max_aid_amount:,.0f}")
    t2.add_row("PNRR Grant", f"€{result.grant_pnrr:,.0f}")
    t2.add_row("─" * 20, "─" * 15)
    t2.add_row("[bold]Gap (ZES Nominal)[/bold]", f"[bold]€{result.gap_zes_nominal:,.0f}[/bold]")
    t2.add_row("ZES Rate", f"{result.zes_rate:.0%}")
    t2.add_row("ZES Base Required", f"€{result.zes_base_required:,.0f}")
    
    console.print(t2)
    console.print()
    
    # Table 3: ZES Base Allocation (Line-by-Line)
    t3 = Table(title="ZES Base Allocation (Waterfall)", box=box.ROUNDED, header_style="bold green")
    t3.add_column("CAPEX Line", style="cyan")
    t3.add_column("Amount", justify="right")
    t3.add_column("ZES Eligible", justify="center")
    t3.add_column("From Over-Cap", justify="right")
    t3.add_column("From Overlap", justify="right")
    t3.add_column("Total Allocated", justify="right")
    
    for alloc in result.allocation_details:
        elig_style = "green" if alloc.zes_eligible.value == "eligible" else ("yellow" if alloc.zes_eligible.value == "partial" else "red")
        t3.add_row(
            alloc.line_name,
            f"€{alloc.line_amount:,.0f}",
            f"[{elig_style}]{alloc.zes_eligible.value}[/{elig_style}]",
            f"€{alloc.allocated_from_overcap:,.0f}" if alloc.allocated_from_overcap > 0 else "-",
            f"€{alloc.allocated_from_overlap:,.0f}" if alloc.allocated_from_overlap > 0 else "-",
            f"€{alloc.total_allocated:,.0f}" if alloc.total_allocated > 0 else "-",
        )
    
    # Totals row
    t3.add_row("─" * 15, "─" * 12, "─" * 10, "─" * 12, "─" * 12, "─" * 12)
    t3.add_row(
        "[bold]TOTAL[/bold]",
        f"[bold]€{result.total_capex:,.0f}[/bold]",
        "",
        f"[bold]€{result.zes_base_from_overcap:,.0f}[/bold]",
        f"[bold]€{result.zes_base_from_overlap:,.0f}[/bold]",
        f"[bold]€{result.zes_base_from_overcap + result.zes_base_from_overlap:,.0f}[/bold]",
    )
    
    console.print(t3)
    console.print()
    
    # Table 4: Riparto Stress
    t4 = Table(title="Riparto (Nominal vs Cash Benefit)", box=box.ROUNDED, header_style="bold green")
    t4.add_column("Item", style="cyan")
    t4.add_column("Value", justify="right")
    
    t4.add_row("ZES Nominal Authorized", f"€{result.zes_nominal_authorized:,.0f}")
    t4.add_row("Riparto Coefficient", f"{result.riparto_coeff:.2%}")
    t4.add_row("─" * 20, "─" * 15)
    t4.add_row("[bold green]ZES Cash Benefit[/bold green]", f"[bold green]€{result.zes_cash_benefit:,.0f}[/bold green]")
    t4.add_row("[red]Benefit Lost[/red]", f"[red]€{result.zes_benefit_lost:,.0f}[/red]")
    
    console.print(t4)
    console.print()
    
    # Table 5: Compliance Check
    t5 = Table(title="Compliance Check", box=box.ROUNDED, header_style="bold green")
    t5.add_column("Check", style="cyan")
    t5.add_column("Value", justify="center")
    t5.add_column("Status", justify="center")
    
    intensity_pass = abs(result.nominal_aid_intensity - result.max_esl_intensity) < 0.001
    t5.add_row(
        "Nominal Aid Intensity",
        f"{result.nominal_aid_intensity:.1%}",
        "[green]✓ PASS[/green]" if intensity_pass else "[red]✗ FAIL[/red]",
    )
    t5.add_row(
        "Connection Excluded from ZES",
        "Yes" if result.connection_excluded else "No",
        "[green]✓ PASS[/green]" if result.connection_excluded else "[red]✗ FAIL[/red]",
    )
    t5.add_row("─" * 25, "─" * 10, "─" * 12)
    overall_status = "[bold green]✓ ALL PASS[/bold green]" if result.compliance_pass else "[bold red]✗ FAILED[/bold red]"
    t5.add_row("[bold]Overall Compliance[/bold]", "", overall_status)
    
    console.print(t5)
    console.print()
    
    # Summary panel
    total_cash = result.grant_pnrr + result.zes_cash_benefit
    console.print(Panel(
        f"[bold]Total Cash Benefit:[/bold] €{total_cash:,.0f}\n"
        f"  PNRR Grant: €{result.grant_pnrr:,.0f}\n"
        f"  ZES Cash: €{result.zes_cash_benefit:,.0f}",
        title="[bold green]💰 Incentives Summary[/bold green]",
        border_style="green",
    ))
    console.print()


def display_all_biometano(
    projections: BiometanoProjections,
    statements: FinancialStatements,
    valuation: ValuationOutputs,
    sensitivity: Optional[SensitivityAnalysisOutputs] = None,
    methodology: str = "enterprise",
) -> None:
    """Display all Biometano outputs in correct section order.
    
    Order:
    1. (Inputs Recap - called separately before this)
    2. Production Summary
    3. Revenue by Channel
    4. OPEX Breakdown
    5. Financial Statement (Yearly)
    6. Balance Sheet (Yearly)
    7. FCFF Schedule
    8. Valuation Summary (compact)
    9. Sensitivity Analysis (if provided)
    10. Scenario Comparison (if provided)
    """
    console.print()
    console.rule("[bold cyan]Biometano Project Finance Analysis[/bold cyan]")
    console.print()
    
    # Section 2: Production
    display_production_summary(projections)
    
    # Section 3: Revenue (correct column order)
    display_revenue_breakdown(projections)
    
    # Section 4: OPEX
    display_opex_breakdown(projections)
    
    # Section 5: Income Statement
    display_income_statement(projections, statements)
    
    # Section 6: Financial Statement / FCFF Bridge
    display_fcff_bridge(projections, statements)
    
    # Section 7: Balance Sheet Recap
    display_balance_sheet_recap(statements, projections)
    
    # Section 8: FCFF Schedule
    display_fcff_schedule(projections)
    
    # Section 8: Valuation Summary (no bridge, no commentary)
    display_valuation_summary(valuation, methodology)
    
    # Section 9 & 10: Sensitivity (if provided)
    if sensitivity:
        display_sensitivity_tornado(sensitivity, methodology)
        display_scenario_comparison(sensitivity, methodology)
    
    console.print("[green]✓ Biometano Analysis Complete[/green]")
    console.print()
