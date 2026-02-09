"""
Biometano CLI Commands

Typer-based CLI commands for Biometano project finance analysis.

Report Section Order:
1. Inputs Recap
2. Production Summary  
3. Revenue by Channel
4. OPEX Breakdown
5. Financial Statement (Yearly)
6. Balance Sheet (Yearly)
7. FCFF Schedule
8. Discounting + PV + Terminal Value
9. Valuation Summary (compact, EV-based)
10. Sensitivity Analysis
11. Scenario Comparison
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer
import yaml
from rich.console import Console

from dcf_projects.biometano.schema import BiometanoCase, DEFAULT_ZES_CREDIT_RATE
from dcf_projects.biometano.builder import build_projections
from dcf_projects.biometano.statements import build_statements
from dcf_projects.biometano.valuation import compute_valuation
from dcf_projects.biometano.sensitivities import run_sensitivity_analysis
from dcf_projects.biometano.inputs_recap import display_inputs_recap
from dcf_ui_cli.biometano_display import (
    display_all_biometano,
    display_sensitivity_tornado,
    display_scenario_comparison,
    display_incentives_waterfall,
)
from dcf_ui_cli.biometano_charts import save_biometano_charts, show_biometano_charts
from dcf_io.writers import export_xlsx_biometano, export_csv_biometano


app = typer.Typer(
    name="biometano",
    help="Biometano project finance commands",
)
console = Console()


def _open_path(path: Path) -> None:
    """Open a file or directory in the system file explorer."""
    try:
        # typer.launch(..., locate=True) reveals file in finder/explorer
        # typer.launch(...) opens the file/folder
        path_str = str(path)
        if path.is_file():
            typer.launch(path_str, locate=True)
            console.print(f"[dim]Revealed in Finder: {path.name}[/dim]")
        else:
            typer.launch(path_str)
            console.print(f"[dim]Opened folder: {path.name}[/dim]")
    except Exception as e:
        console.print(f"[yellow]Could not open path: {e}[/yellow]")


def _resolve_input_file(input_file: Optional[Path], input_option: Optional[Path]) -> Path:
    """Resolve input file from positional arg or --input option."""
    resolved = input_option or input_file
    if resolved is None:
        raise typer.BadParameter("Missing input file. Provide a positional INPUT_FILE or --input.")
    if not resolved.exists():
        raise typer.BadParameter(f"Input file not found: {resolved}")
    if not resolved.is_file():
        raise typer.BadParameter(f"Input path is not a file: {resolved}")
    return resolved


def _load_case(input_file: Path) -> BiometanoCase:
    """Load and validate a BiometanoCase from YAML/JSON."""
    with open(input_file) as f:
        if input_file.suffix in (".yaml", ".yml"):
            data = yaml.safe_load(f)
        else:
            import json
            data = json.load(f)
    
    return BiometanoCase.model_validate(data)


@app.command()
def init(
    output_file: Path = typer.Option(
        Path("case_files/biometano_case.yaml"),
        "--output", "-o",
        help="Output file path",
    ),
) -> None:
    """
    Generate a Biometano case template YAML file.
    
    Creates a template with default values that can be customized.
    Uses ZES credit rate of 14.6189% by default.
    """
    template = {
        "horizon": {
            "base_year": 2024,
            "years_forecast": 15,
            "construction_years": 2,
        },
        "production": {
            "forsu_throughput_tpy": 60000,
            "biomethane_smc_y": 4000000,
            "kwh_per_smc": 10,
            "impurity_rate": 0.20,  # 20% sovvalli
            "availability_profile": [0.75, 0.90, 0.95, 0.95, 0.95],
            "compost_tpy": 12148,
            "co2_tpy": 4560,
        },
        "revenues": {
            "gate_fee": {"price": 190, "payment_delay_days": 45, "escalation_rate": 0},
            "tariff": {"price": 70.16, "payment_delay_days": 90, "escalation_rate": 0},
            "go": {"price": 0.3, "payment_delay_days": 60, "escalation_rate": 0},  # GO before CO2
            "co2": {"price": 120, "payment_delay_days": 45, "escalation_rate": 0},
            "compost": {"price": 5, "payment_delay_days": 30, "escalation_rate": 0},
        },
        "opex": {
            "feedstock_handling": {"fixed_annual": 500000, "variable_per_tonne": 5},
            "utilities": {"fixed_annual": 300000, "variable_per_mwh": 2},
            "chemicals": {"fixed_annual": 100000},
            "maintenance": {"percent_of_capex": 0.02},
            "personnel": {"fixed_annual": 800000},
            "insurance": {"percent_of_capex": 0.005},
            "overheads": {"fixed_annual": 200000},
            "digestate_handling": {"variable_per_tonne": 3},
            "other": {"fixed_annual": 50000},
        },
        "capex": {
            "epc": {"amount": 25000000, "spend_profile": [0.4, 0.6], "useful_life_years": 20, "eligible_for_grant": True},
            "civils": {"amount": 5000000, "spend_profile": [0.7, 0.3], "useful_life_years": 30, "eligible_for_grant": True},
            "upgrading_unit": {"amount": 8000000, "spend_profile": [0.3, 0.7], "useful_life_years": 15, "eligible_for_grant": True},
            "grid_connection": {"amount": 2000000, "spend_profile": [0.5, 0.5], "useful_life_years": 25, "eligible_for_grant": True},
            "engineering": {"amount": 1500000, "spend_profile": [0.8, 0.2], "useful_life_years": 20},
            "permitting": {"amount": 500000, "spend_profile": [1.0], "useful_life_years": 20},
            "contingency": {"amount": 2000000, "spend_profile": [0.3, 0.7], "useful_life_years": 20},
            "startup_costs": {"amount": 500000, "spend_profile": [0.0, 1.0], "capitalize": False},
            "capitalize_idc": False,
        },
        "financing": {
            "debt_amount": 30000000,
            "debt_drawdown_profile": [0.4, 0.6],
            "cost_of_debt": 0.05,
            "debt_repayment_years": 12,
            "cash_at_base": 5000000,
            "equity_book_at_base": 15000000,
            "tax_rate": 0.24,
            "rf": 0.03,
            "rm": 0.08,
            "beta": 1.2,
        },
        "incentives": {
            "capital_grant": {
                "enabled": True,
                "percent_of_eligible": 0.40,  # 40% PNRR grant
                "recognition_trigger": "at_cod",
                "cash_receipt_schedule": [0.5, 0.5],
                "accounting_policy": "A2",
            },
            "tax_credit": {
                "enabled": True,
                "percent_of_eligible": DEFAULT_ZES_CREDIT_RATE,  # 14.6189% ZES
                "eligible_base": "capex",
                "carry_forward_years": 5,
                "accounting_policy": "B1",
            },
        },
        "terminal_value": {
            "method": "perpetuity",
            "perpetuity_growth": 0.0,
        },
    }
    
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w") as f:
        yaml.dump(template, f, default_flow_style=False, sort_keys=False)
    
    console.print(f"[green]✓ Created template: {output_file}[/green]")
    console.print(f"[dim]  ZES Credit: {DEFAULT_ZES_CREDIT_RATE:.4%}[/dim]")


@app.command()
def run(
    input_file: Optional[Path] = typer.Argument(
        None,
        help="Path to Biometano case YAML/JSON file",
    ),
    input_option: Optional[Path] = typer.Option(
        None,
        "--input",
        "-i",
        help="Path to Biometano case YAML/JSON file",
    ),
    value: str = typer.Option(
        "enterprise",  # Default changed to enterprise
        "--value",
        help="Valuation method: 'enterprise' (default), 'equity', or 'both'",
    ),
    charts: bool = typer.Option(
        False,
        "--charts",
        help="Display interactive charts",
    ),
    charts_dir: Optional[Path] = typer.Option(
        None,
        "--charts-dir",
        help="Directory to save chart files",
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Output Excel file path",
    ),
    xlsx_mode: str = typer.Option(
        "formulas",
        "--xlsx-mode",
        help="Excel export mode: formulas or values",
    ),
    quiet: bool = typer.Option(
        False,
        "--quiet", "-q",
        help="Suppress table output",
    ),
    open_folder: bool = typer.Option(
        False,
        "--open",
        help="Open output folder/file after completion",
    ),
) -> None:
    """
    Run Biometano project finance analysis.
    
    Default methodology: Enterprise Value (FCFF/WACC).
    """
    try:
        input_file = _resolve_input_file(input_file, input_option)
        console.print(f"[dim]Loading case: {input_file}[/dim]")
        case = _load_case(input_file)
        
        console.print("[dim]Building projections...[/dim]")
        projections = build_projections(case)
        
        console.print("[dim]Generating statements...[/dim]")
        statements = build_statements(case, projections)
        
        console.print("[dim]Computing valuation...[/dim]")
        valuation = compute_valuation(case, projections)
        
        if not quiet:
            # Section 1: Inputs Recap (first)
            display_inputs_recap(case, console)
            
            # Sections 2-8: Core analysis
            display_all_biometano(projections, statements, valuation, methodology=value)
        
        if charts:
            console.print("[dim]Opening charts in browser...[/dim]")
            show_biometano_charts(projections, valuation)
        
        if charts_dir:
            console.print(f"[dim]Saving charts to: {charts_dir}[/dim]")
            files = save_biometano_charts(projections, valuation, charts_dir)
            console.print(f"[green]✓ Saved {len(files)} chart files[/green]")
            if open_folder and not output:  # prioritize output file if both set
                _open_path(Path(charts_dir))

        if output:
            console.print(f"[dim]Exporting to: {output}[/dim]")
            export_xlsx_biometano(
                projections,
                statements,
                valuation,
                output,
                case=case,
                xlsx_mode=xlsx_mode,
            )
            console.print(f"[green]✓ Exported to {output}[/green]")
            if open_folder:
                _open_path(output)
        
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(code=1)


@app.command()
def sens(
    input_file: Optional[Path] = typer.Argument(
        None,
        help="Path to Biometano case YAML/JSON file",
    ),
    input_option: Optional[Path] = typer.Option(
        None,
        "--input",
        "-i",
        help="Path to Biometano case YAML/JSON file",
    ),
    value: str = typer.Option(
        "enterprise",  # Default to enterprise
        "--value",
        help="Valuation method for sensitivity: 'enterprise' (default) or 'equity'",
    ),
    charts: bool = typer.Option(
        False,
        "--charts",
        help="Display interactive charts",
    ),
    charts_dir: Optional[Path] = typer.Option(
        None,
        "--charts-dir",
        help="Directory to save chart files",
    ),
    open_folder: bool = typer.Option(
        False,
        "--open",
        help="Open output folder after completion",
    ),
) -> None:
    """
    Run sensitivity analysis on a Biometano case.
    
    Default methodology: Enterprise Value (FCFF/WACC).
    Produces tornado chart and scenario comparison.
    """
    try:
        input_file = _resolve_input_file(input_file, input_option)
        console.print(f"[dim]Loading case: {input_file}[/dim]")
        case = _load_case(input_file)
        
        console.print("[dim]Running sensitivity analysis...[/dim]")
        
        # Updated value function signature (equity, ev, pv_fcff, pv_tv)
        def value_fn(c: BiometanoCase) -> tuple[float, float, float, float]:
            proj = build_projections(c)
            val = compute_valuation(c, proj)
            return val.equity_value, val.enterprise_value, val.sum_pv_fcff, val.pv_terminal_value_fcff
        
        sensitivity = run_sensitivity_analysis(case, value_function=value_fn)
        
        display_sensitivity_tornado(sensitivity, methodology=value)
        display_scenario_comparison(sensitivity, methodology=value)
        
        if charts:
            from dcf_ui_cli.biometano_charts import create_tornado_chart, create_scenario_comparison_chart
            create_tornado_chart(sensitivity, methodology=value).show()
            create_scenario_comparison_chart(sensitivity, methodology=value).show()
        
        if charts_dir:
            charts_dir = Path(charts_dir)
            charts_dir.mkdir(parents=True, exist_ok=True)
            
            from dcf_ui_cli.biometano_charts import create_tornado_chart, create_scenario_comparison_chart
            
            create_tornado_chart(sensitivity, methodology=value).write_html(str(charts_dir / "sensitivity_tornado.html"))
            create_scenario_comparison_chart(sensitivity, methodology=value).write_html(str(charts_dir / "scenario_comparison.html"))
            
            console.print(f"[green]✓ Saved sensitivity charts to {charts_dir}[/green]")
            if open_folder:
                _open_path(charts_dir)
        
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(code=1)


@app.command()
def report(
    input_file: Optional[Path] = typer.Argument(
        None,
        help="Path to Biometano case YAML/JSON file",
    ),
    input_option: Optional[Path] = typer.Option(
        None,
        "--input",
        "-i",
        help="Path to Biometano case YAML/JSON file",
    ),
    output_dir: Path = typer.Option(
        Path("output/biometano"),
        "--output", "-o",
        help="Output directory for all exports",
    ),
    xlsx_mode: str = typer.Option(
        "formulas",
        "--xlsx-mode",
        help="Excel export mode: formulas or values",
    ),
    value: str = typer.Option(
        "enterprise",  # Default to enterprise
        "--value",
        help="Valuation method: 'enterprise' (default) or 'equity'",
    ),
    open_folder: bool = typer.Option(
        False,
        "--open",
        help="Open output directory after completion",
    ),
) -> None:
    """
    Generate full Biometano report with statements, valuation, and sensitivities.
    
    Default methodology: Enterprise Value (FCFF/WACC).
    
    Report Section Order:
    1. Inputs Recap
    2. Production Summary
    3. Revenue by Channel  
    4. OPEX Breakdown
    5. Financial Statement (Yearly)
    6. Balance Sheet (Yearly)
    7. FCFF Schedule
    8. Discounting + PV + Terminal Value
    9. Valuation Summary
    10. Sensitivity Analysis
    11. Scenario Comparison
    """
    try:
        input_file = _resolve_input_file(input_file, input_option)
        console.print(f"[dim]Loading case: {input_file}[/dim]")
        case = _load_case(input_file)
        
        console.print("[dim]Building projections...[/dim]")
        projections = build_projections(case)
        
        console.print("[dim]Generating statements...[/dim]")
        statements = build_statements(case, projections)
        
        console.print("[dim]Computing valuation...[/dim]")
        valuation = compute_valuation(case, projections)
        
        console.print("[dim]Running sensitivity analysis...[/dim]")
        
        def value_fn(c: BiometanoCase) -> tuple[float, float, float, float]:
            proj = build_projections(c)
            val = compute_valuation(c, proj)
            return val.equity_value, val.enterprise_value, val.sum_pv_fcff, val.pv_terminal_value_fcff
        
        sensitivity = run_sensitivity_analysis(case, value_function=value_fn)
        
        # Section 1: Inputs Recap (FIRST)
        display_inputs_recap(case, console)
        
        # Section 2: Incentives Waterfall (if enabled)
        if projections.incentive_allocation:
            display_incentives_waterfall(projections.incentive_allocation)
        
        # Sections 3-9: Core analysis (includes Financial Statement, Balance Sheet, no bridge/commentary)
        display_all_biometano(projections, statements, valuation, sensitivity, methodology=value)
        
        # Sections 10-11: Sensitivity (included in display_all_biometano if sensitivity provided)
        
        # Export
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Excel
        xlsx_path = output_dir / "biometano_report.xlsx"
        export_xlsx_biometano(
            projections,
            statements,
            valuation,
            xlsx_path,
            case=case,
            xlsx_mode=xlsx_mode,
        )
        console.print(f"[green]✓ Exported Excel: {xlsx_path}[/green]")
        
        # CSV
        csv_dir = output_dir / "csv"
        export_csv_biometano(projections, statements, valuation, csv_dir)
        console.print(f"[green]✓ Exported CSVs to: {csv_dir}[/green]")
        
        # Charts (no valuation waterfall, no bridge)
        charts_dir = output_dir / "charts"
        files = save_biometano_charts(projections, valuation, charts_dir, sensitivity)
        console.print(f"[green]✓ Saved {len(files)} charts to: {charts_dir}[/green]")
        
        console.print()
        console.print("[bold green]✓ Full report generated successfully[/bold green]")
        
        if open_folder:
            _open_path(output_dir)
        
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(code=1)
