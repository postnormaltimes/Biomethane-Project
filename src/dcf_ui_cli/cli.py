"""
DCF CLI Application

Typer-based command-line interface for DCF modeling tool.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer
from rich.console import Console

from dcf_engine.engine import DCFEngine
from dcf_io.readers import read_input_file
from dcf_io.writers import export_xlsx, export_csv
from dcf_ui_cli.display import display_all
from dcf_ui_cli.charts import save_charts, show_charts


app = typer.Typer(
    name="dcf",
    help="Production-quality DCF modeling tool",
    add_completion=False,
)

# Register biometano subcommand
try:
    from dcf_ui_cli.biometano_cli import app as biometano_app
    app.add_typer(biometano_app, name="biometano")
except ImportError:
    pass  # Biometano module not available



console = Console()


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


@app.command()
def run(
    input_file: Optional[Path] = typer.Argument(
        None,
        help="Path to input file (YAML or JSON)",
    ),
    input_option: Optional[Path] = typer.Option(
        None,
        "--input",
        "-i",
        help="Path to input file (YAML or JSON)",
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Output Excel file path",
    ),
    csv_dir: Optional[Path] = typer.Option(
        None,
        "--csv-dir",
        help="Directory to export CSV files",
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
    quiet: bool = typer.Option(
        False,
        "--quiet", "-q",
        help="Suppress table output",
    ),
) -> None:
    """
    Run DCF analysis on an input file.
    
    Reads a YAML or JSON input file, computes the full DCF model,
    and displays results as rich tables. Optionally exports to Excel/CSV
    and generates charts.
    """
    try:
        input_file = _resolve_input_file(input_file, input_option)
        # Read inputs
        console.print(f"[dim]Reading input file: {input_file}[/dim]")
        inputs = read_input_file(input_file)
        
        # Run engine
        console.print("[dim]Running DCF engine...[/dim]")
        engine = DCFEngine(inputs)
        outputs = engine.run()
        
        # Display results
        if not quiet:
            display_all(outputs)
        
        # Export to Excel
        if output:
            console.print(f"\n[dim]Exporting to Excel: {output}[/dim]")
            export_xlsx(outputs, output)
            console.print(f"[green]✓ Exported to {output}[/green]")
        
        # Export to CSV
        if csv_dir:
            console.print(f"\n[dim]Exporting CSVs to: {csv_dir}[/dim]")
            files = export_csv(outputs, csv_dir)
            console.print(f"[green]✓ Exported {len(files)} CSV files[/green]")
        
        # Charts
        if charts:
            console.print("\n[dim]Opening charts in browser...[/dim]")
            show_charts(outputs)
        
        if charts_dir:
            console.print(f"\n[dim]Saving charts to: {charts_dir}[/dim]")
            files = save_charts(outputs, charts_dir)
            console.print(f"[green]✓ Saved {len(files)} chart files[/green]")
    
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(code=1)


@app.command()
def validate(
    input_file: Optional[Path] = typer.Argument(
        None,
        help="Path to input file (YAML or JSON)",
    ),
    input_option: Optional[Path] = typer.Option(
        None,
        "--input",
        "-i",
        help="Path to input file (YAML or JSON)",
    ),
) -> None:
    """
    Validate an input file without running the full DCF.
    
    Checks that all required fields are present and consistent.
    """
    try:
        input_file = _resolve_input_file(input_file, input_option)
        console.print(f"[dim]Validating: {input_file}[/dim]")
        inputs = read_input_file(input_file)
        
        # Create engine (which validates inputs)
        DCFEngine(inputs)
        
        console.print("[green]✓ Input file is valid[/green]")
        
        # Show summary
        console.print(f"\n  Base year: {inputs.timeline.base_year}")
        console.print(f"  Forecast years: {inputs.timeline.forecast_years}")
        console.print(f"  Discounting mode: {inputs.discounting_mode.value}")
        console.print(f"  Terminal value method: {inputs.terminal_value.method.value}")
        console.print(f"  WACC weighting: {inputs.wacc.weighting_mode.value}")
    
    except Exception as e:
        console.print(f"[red]Validation failed: {e}[/red]")
        raise typer.Exit(code=1)


@app.command()
def export(
    input_file: Optional[Path] = typer.Argument(
        None,
        help="Path to input file (YAML or JSON)",
    ),
    input_option: Optional[Path] = typer.Option(
        None,
        "--input",
        "-i",
        help="Path to input file (YAML or JSON)",
    ),
    output: Path = typer.Option(
        ...,
        "--output", "-o",
        help="Output Excel file path",
    ),
    include_csv: bool = typer.Option(
        False,
        "--csv",
        help="Also export CSV files to same directory",
    ),
    include_charts: bool = typer.Option(
        False,
        "--charts",
        help="Also save chart files to same directory",
    ),
) -> None:
    """
    Run DCF and export results to files.
    
    Primary export is Excel. Optionally also exports CSV and charts.
    """
    try:
        input_file = _resolve_input_file(input_file, input_option)
        # Read and run
        inputs = read_input_file(input_file)
        engine = DCFEngine(inputs)
        outputs = engine.run()
        
        # Export Excel
        export_xlsx(outputs, output)
        console.print(f"[green]✓ Exported to {output}[/green]")
        
        # Export CSV
        if include_csv:
            csv_dir = output.parent / "csv"
            files = export_csv(outputs, csv_dir)
            console.print(f"[green]✓ Exported {len(files)} CSV files to {csv_dir}[/green]")
        
        # Export charts
        if include_charts:
            charts_dir = output.parent / "charts"
            files = save_charts(outputs, charts_dir)
            console.print(f"[green]✓ Saved {len(files)} chart files to {charts_dir}[/green]")
    
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(code=1)


if __name__ == "__main__":
    app()
