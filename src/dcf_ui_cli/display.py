"""
DCF CLI Display

Rich table formatting for terminal output.
"""
from __future__ import annotations

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text

from dcf_engine.models import DCFOutputs


console = Console()


def display_header(title: str) -> None:
    """Display a section header."""
    console.print()
    console.print(Panel(Text(title, style="bold white"), style="blue"))


def display_inputs_summary(outputs: DCFOutputs) -> None:
    """Display inputs/assumptions summary."""
    display_header("📊 Inputs Summary")
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Parameter", style="dim")
    table.add_column("Value", justify="right")
    
    table.add_row("Base Year", str(outputs.base_year))
    table.add_row("Forecast Years", ", ".join(map(str, outputs.forecast_years)))
    table.add_row("Discounting Mode", outputs.discounting_mode.value)
    table.add_row("Cost of Equity (Ke)", f"{outputs.ke:.4f}")
    table.add_row("Terminal Value Method", outputs.terminal_value.method.value)
    
    if outputs.terminal_value.growth_rate is not None:
        table.add_row("Terminal Growth Rate", f"{outputs.terminal_value.growth_rate:.4f}")
    
    console.print(table)


def display_projections(outputs: DCFOutputs) -> None:
    """Display operating projections."""
    display_header("📈 Operating Projections")
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Year", justify="center")
    table.add_column("Revenue", justify="right")
    table.add_column("Op. Costs", justify="right")
    table.add_column("EBITDA", justify="right")
    table.add_column("D&A", justify="right")
    table.add_column("EBIT", justify="right")
    table.add_column("Tax on EBIT", justify="right")
    table.add_column("NOPAT", justify="right")
    
    for p in outputs.projections:
        table.add_row(
            str(p.year),
            f"{p.revenue:,.2f}",
            f"{p.operating_costs:,.2f}",
            f"{p.ebitda:,.2f}",
            f"{p.depreciation_amortization:,.2f}",
            f"{p.ebit:,.2f}",
            f"{p.tax_on_ebit:,.2f}",
            f"{p.nopat:,.2f}",
        )
    
    console.print(table)


def display_nwc(outputs: DCFOutputs) -> None:
    """Display NWC schedule."""
    display_header("💰 Net Working Capital Schedule")
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Year", justify="center")
    table.add_column("NWC", justify="right")
    table.add_column("ΔNWC", justify="right")
    
    for n in outputs.nwc_schedule:
        table.add_row(
            str(n.year),
            f"{n.nwc:,.2f}",
            f"{n.delta_nwc:,.2f}",
        )
    
    console.print(table)


def display_cash_flows(outputs: DCFOutputs) -> None:
    """Display cash flows."""
    display_header("💸 Cash Flows")
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Year", justify="center")
    table.add_column("NOPAT", justify="right")
    table.add_column("+ D&A", justify="right")
    table.add_column("- ΔNWC", justify="right")
    table.add_column("- Capex", justify="right")
    table.add_column("= FCFF", justify="right", style="bold green")
    table.add_column("- Int*(1-t)", justify="right")
    table.add_column("+ NetBorrow", justify="right")
    table.add_column("= FCFE", justify="right", style="bold blue")
    
    for cf in outputs.cash_flows:
        after_tax_int = cf.interest_expense - cf.interest_tax_shield
        table.add_row(
            str(cf.year),
            f"{cf.nopat:,.2f}",
            f"{cf.depreciation_amortization:,.2f}",
            f"{cf.delta_nwc:,.2f}",
            f"{cf.capex:,.2f}",
            f"{cf.fcff:,.2f}",
            f"{after_tax_int:,.2f}",
            f"{cf.net_borrowing:,.2f}",
            f"{cf.fcfe:,.2f}",
        )
    
    console.print(table)


def display_wacc_details(outputs: DCFOutputs) -> None:
    """Display WACC computation details."""
    display_header("📊 WACC Details")
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Year", justify="center")
    table.add_column("Ke", justify="right")
    table.add_column("Rd", justify="right")
    table.add_column("Tax Rate", justify="right")
    table.add_column("Debt", justify="right")
    table.add_column("Equity Book", justify="right")
    table.add_column("wD", justify="right")
    table.add_column("wE", justify="right")
    table.add_column("WACC", justify="right", style="bold yellow")
    
    for w in outputs.wacc_details:
        table.add_row(
            str(w.year),
            f"{w.ke:.4f}",
            f"{w.rd:.4f}",
            f"{w.tax_rate:.4f}",
            f"{w.debt:,.2f}",
            f"{w.equity_book:,.2f}",
            f"{w.weight_debt:.4f}",
            f"{w.weight_equity:.4f}",
            f"{w.wacc:.10f}",
        )
    
    console.print(table)


def display_pv_decomposition(outputs: DCFOutputs) -> None:
    """Display PV decomposition."""
    display_header("📉 Present Value Decomposition")
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Year", justify="center")
    table.add_column("Period", justify="center")
    table.add_column("WACC", justify="right")
    table.add_column("DF(WACC)", justify="right")
    table.add_column("FCFF", justify="right")
    table.add_column("PV(FCFF)", justify="right", style="green")
    table.add_column("Ke", justify="right")
    table.add_column("DF(Ke)", justify="right")
    table.add_column("FCFE", justify="right")
    table.add_column("PV(FCFE)", justify="right", style="blue")
    
    total_pv_fcff = 0.0
    total_pv_fcfe = 0.0
    
    for d in outputs.discount_schedule:
        table.add_row(
            str(d.year),
            str(d.period),
            f"{d.wacc:.6f}",
            f"{d.discount_factor_wacc:.6f}",
            f"{d.fcff:,.2f}",
            f"{d.pv_fcff:,.2f}",
            f"{d.ke:.6f}",
            f"{d.discount_factor_ke:.6f}",
            f"{d.fcfe:,.2f}",
            f"{d.pv_fcfe:,.2f}",
        )
        total_pv_fcff += d.pv_fcff
        total_pv_fcfe += d.pv_fcfe
    
    table.add_row(
        "[bold]Total[/bold]", "", "", "", "",
        f"[bold]{total_pv_fcff:,.2f}[/bold]",
        "", "", "",
        f"[bold]{total_pv_fcfe:,.2f}[/bold]",
    )
    
    console.print(table)


def display_terminal_value(outputs: DCFOutputs) -> None:
    """Display terminal value computation."""
    display_header("🎯 Terminal Value")
    
    tv = outputs.terminal_value
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Item", style="dim")
    table.add_column("Value", justify="right")
    
    table.add_row("Final Year", str(tv.final_year))
    table.add_row("Method", tv.method.value)
    
    if tv.growth_rate is not None:
        table.add_row("Growth Rate (g)", f"{tv.growth_rate:.4f}")
    if tv.exit_multiple:
        table.add_row("Exit Multiple", f"{tv.exit_multiple:.2f}x")
        table.add_row("Exit Metric", tv.exit_metric or "N/A")
    
    table.add_row("", "")
    table.add_row("[bold]Terminal Value (FCFF)[/bold]", f"[bold green]{tv.terminal_value_fcff:,.2f}[/bold green]")
    table.add_row("[bold]Terminal Value (FCFE)[/bold]", f"[bold blue]{tv.terminal_value_fcfe:,.2f}[/bold blue]")
    table.add_row("", "")
    table.add_row("Discount Rate (WACC)", f"{tv.discount_rate_wacc:.6f}")
    table.add_row("Discount Rate (Ke)", f"{tv.discount_rate_ke:.6f}")
    table.add_row("", "")
    table.add_row("[bold]PV of TV (FCFF)[/bold]", f"[bold green]{tv.pv_terminal_value_fcff:,.2f}[/bold green]")
    table.add_row("[bold]PV of TV (FCFE)[/bold]", f"[bold blue]{tv.pv_terminal_value_fcfe:,.2f}[/bold blue]")
    
    console.print(table)


def display_valuation_bridge(outputs: DCFOutputs) -> None:
    """Display valuation bridge."""
    display_header("🌉 Valuation Bridge")
    
    vb = outputs.valuation_bridge
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Item", style="dim")
    table.add_column("FCFF/WACC", justify="right")
    table.add_column("FCFE/Ke", justify="right")
    
    table.add_row("Sum PV(Cash Flows)", f"{vb.sum_pv_fcff:,.2f}", f"{vb.sum_pv_fcfe:,.2f}")
    table.add_row("+ PV(Terminal Value)", f"{vb.pv_terminal_value_fcff:,.2f}", f"{vb.pv_terminal_value_fcfe:,.2f}")
    table.add_row("", "", "")
    table.add_row("[bold]Enterprise Value[/bold]", f"[bold]{vb.enterprise_value:,.2f}[/bold]", "-")
    table.add_row("", "", "")
    table.add_row("Less: Debt at Base", f"({vb.debt_at_base:,.2f})", "-")
    table.add_row("Plus: Cash at Base", f"{vb.cash_at_base:,.2f}", "-")
    table.add_row("Net Debt", f"{vb.net_debt:,.2f}", "-")
    table.add_row("", "", "")
    table.add_row(
        "[bold green]Equity Value[/bold green]",
        f"[bold green]{vb.equity_value_from_ev:,.2f}[/bold green]",
        f"[bold blue]{vb.equity_value_direct:,.2f}[/bold blue]",
    )
    
    console.print(table)


def display_reconciliation(outputs: DCFOutputs) -> None:
    """Display FCFF vs FCFE reconciliation."""
    display_header("🔄 FCFF vs FCFE Reconciliation")
    
    vb = outputs.valuation_bridge
    
    table = Table(show_header=True, header_style="bold cyan")
    table.add_column("Item", style="dim")
    table.add_column("Value", justify="right")
    
    table.add_row("Equity from FCFF/WACC", f"{vb.equity_value_from_ev:,.2f}")
    table.add_row("Equity from FCFE/Ke", f"{vb.equity_value_direct:,.2f}")
    table.add_row("", "")
    table.add_row("[bold]Difference[/bold]", f"[bold]{vb.reconciliation_difference:,.2f}[/bold]")
    
    if vb.equity_value_from_ev != 0:
        pct = abs(vb.reconciliation_difference / vb.equity_value_from_ev * 100)
        table.add_row("Difference (%)", f"{pct:.4f}%")
    
    console.print(table)
    
    # Notes
    if vb.reconciliation_notes:
        console.print()
        for note in vb.reconciliation_notes:
            console.print(f"  • {note}")


def display_all(outputs: DCFOutputs) -> None:
    """Display all DCF outputs."""
    display_inputs_summary(outputs)
    display_projections(outputs)
    display_nwc(outputs)
    display_cash_flows(outputs)
    display_wacc_details(outputs)
    display_pv_decomposition(outputs)
    display_terminal_value(outputs)
    display_valuation_bridge(outputs)
    display_reconciliation(outputs)
    
    console.print()
    console.print("[bold green]✓ DCF Analysis Complete[/bold green]")
