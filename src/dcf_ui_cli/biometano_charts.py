"""
Biometano CLI Charts

Plotly visualizations for Biometano project finance analysis.
No valuation bridge/waterfall chart per requirements.
Uses Enterprise Value as default methodology.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import plotly.graph_objects as go
from plotly.subplots import make_subplots

from dcf_projects.biometano.builder import BiometanoProjections
from dcf_projects.biometano.valuation import ValuationOutputs
from dcf_projects.biometano.sensitivities import SensitivityAnalysisOutputs


def create_revenue_stack_chart(projections: BiometanoProjections) -> go.Figure:
    """Create stacked bar chart of revenue by channel over time.
    
    Column order: Gate Fee, Tariff, GO, CO₂, Compost (per requirements)
    """
    years = [r.year for r in projections.revenues if r.year >= projections.cod_year]
    
    # Correct order: Gate Fee, Tariff, GO, CO₂, Compost
    gate_fee = [r.gate_fee for r in projections.revenues if r.year >= projections.cod_year]
    tariff = [r.tariff for r in projections.revenues if r.year >= projections.cod_year]
    go_rev = [r.go for r in projections.revenues if r.year >= projections.cod_year]
    co2 = [r.co2 for r in projections.revenues if r.year >= projections.cod_year]
    compost = [r.compost for r in projections.revenues if r.year >= projections.cod_year]
    
    fig = go.Figure()
    
    # Add in correct order
    fig.add_trace(go.Bar(name="Gate Fee", x=years, y=gate_fee, marker_color="#2ecc71"))
    fig.add_trace(go.Bar(name="Tariff", x=years, y=tariff, marker_color="#3498db"))
    fig.add_trace(go.Bar(name="GO", x=years, y=go_rev, marker_color="#f39c12"))
    fig.add_trace(go.Bar(name="CO₂", x=years, y=co2, marker_color="#9b59b6"))
    fig.add_trace(go.Bar(name="Compost", x=years, y=compost, marker_color="#e74c3c"))
    
    fig.update_layout(
        title="Revenue by Channel",
        xaxis_title="Year",
        yaxis_title="Revenue (€)",
        barmode="stack",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    
    return fig


def create_opex_stack_chart(projections: BiometanoProjections) -> go.Figure:
    """Create stacked bar chart of OPEX by category over time."""
    years = [o.year for o in projections.opex if o.year >= projections.cod_year]
    
    feedstock = [o.feedstock_handling for o in projections.opex if o.year >= projections.cod_year]
    utilities = [o.utilities for o in projections.opex if o.year >= projections.cod_year]
    maintenance = [o.maintenance for o in projections.opex if o.year >= projections.cod_year]
    personnel = [o.personnel for o in projections.opex if o.year >= projections.cod_year]
    other = [
        o.chemicals + o.insurance + o.overheads + o.digestate_handling + o.other
        for o in projections.opex if o.year >= projections.cod_year
    ]
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(name="Feedstock", x=years, y=feedstock, marker_color="#e74c3c"))
    fig.add_trace(go.Bar(name="Utilities", x=years, y=utilities, marker_color="#f39c12"))
    fig.add_trace(go.Bar(name="Maintenance", x=years, y=maintenance, marker_color="#3498db"))
    fig.add_trace(go.Bar(name="Personnel", x=years, y=personnel, marker_color="#9b59b6"))
    fig.add_trace(go.Bar(name="Other", x=years, y=other, marker_color="#95a5a6"))
    
    fig.update_layout(
        title="OPEX by Category",
        xaxis_title="Year",
        yaxis_title="OPEX (€)",
        barmode="stack",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    
    return fig


def create_fcff_timeline_chart(projections: BiometanoProjections) -> go.Figure:
    """Create FCFF timeline chart (EV-focused - FCFF only)."""
    years = projections.operating_years
    fcff = [projections.fcff.get(y, 0) for y in years]
    
    fig = go.Figure()
    
    # FCFF bars
    fig.add_trace(go.Bar(
        name="FCFF",
        x=years,
        y=fcff,
        marker_color="#2ecc71",
        text=[f"€{v:,.0f}" for v in fcff],
        textposition="outside",
    ))
    
    fig.update_layout(
        title="Free Cash Flow to Firm (FCFF)",
        xaxis_title="Year",
        yaxis_title="FCFF (€)",
        template="plotly_white",
        showlegend=False,
    )
    
    return fig


def create_pv_decomposition_chart(valuation: ValuationOutputs) -> go.Figure:
    """Create PV decomposition chart showing explicit FCFFs vs TV contribution."""
    fig = go.Figure()
    
    explicit_share = valuation.sum_pv_fcff / valuation.enterprise_value * 100 if valuation.enterprise_value > 0 else 0
    tv_share = valuation.pv_terminal_value_fcff / valuation.enterprise_value * 100 if valuation.enterprise_value > 0 else 0
    
    fig.add_trace(go.Bar(
        x=["PV(Explicit FCFFs)", "PV(Terminal Value)"],
        y=[valuation.sum_pv_fcff, valuation.pv_terminal_value_fcff],
        marker_color=["#2ecc71", "#3498db"],
        text=[
            f"€{valuation.sum_pv_fcff:,.0f}<br>({explicit_share:.1f}%)",
            f"€{valuation.pv_terminal_value_fcff:,.0f}<br>({tv_share:.1f}%)",
        ],
        textposition="outside",
    ))
    
    fig.update_layout(
        title=f"Enterprise Value Decomposition (€{valuation.enterprise_value:,.0f})",
        yaxis_title="Value (€)",
        template="plotly_white",
        showlegend=False,
    )
    
    return fig


def create_tornado_chart(
    sensitivity: SensitivityAnalysisOutputs, 
    methodology: str = "enterprise",
) -> go.Figure:
    """Create tornado chart for sensitivity analysis (EV-based by default)."""
    # Get top 10 by spread
    if methodology == "enterprise":
        data = sorted(sensitivity.tornado_data, key=lambda x: x.spread_ev, reverse=True)[:10]
    else:
        data = sorted(sensitivity.tornado_data, key=lambda x: x.spread, reverse=True)[:10]
    
    # Reverse for display (largest at top)
    data = list(reversed(data))
    
    parameters = [d.parameter for d in data]
    
    if methodology == "enterprise":
        low_deltas = [d.low_delta_ev for d in data]
        high_deltas = [d.high_delta_ev for d in data]
        base_value = sensitivity.base_ev
        label = "Enterprise Value"
    else:
        low_deltas = [d.low_delta for d in data]
        high_deltas = [d.high_delta for d in data]
        base_value = sensitivity.base_equity_value
        label = "Equity Value"
    
    fig = go.Figure()
    
    # Low shocks (left side)
    fig.add_trace(go.Bar(
        name="Downside",
        y=parameters,
        x=low_deltas,
        orientation="h",
        marker_color="#e74c3c",
    ))
    
    # High shocks (right side)
    fig.add_trace(go.Bar(
        name="Upside",
        y=parameters,
        x=high_deltas,
        orientation="h",
        marker_color="#2ecc71",
    ))
    
    fig.update_layout(
        title=f"Sensitivity Tornado - {label} (Base: €{base_value:,.0f})",
        xaxis_title=f"Change in {label} (€)",
        barmode="overlay",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    
    # Add vertical line at zero
    fig.add_vline(x=0, line_dash="dash", line_color="gray")
    
    return fig


def create_scenario_comparison_chart(
    sensitivity: SensitivityAnalysisOutputs,
    methodology: str = "enterprise",
) -> go.Figure:
    """Create scenario comparison bar chart (EV-based by default)."""
    scenarios = sensitivity.scenarios
    
    names = [s.name for s in scenarios]
    
    if methodology == "enterprise":
        values = [s.ev for s in scenarios]
        deltas = [s.delta_ev for s in scenarios]
        label = "Enterprise Value"
    else:
        values = [s.equity_value for s in scenarios]
        deltas = [s.delta_from_base for s in scenarios]
        label = "Equity Value"
    
    colors = ["#3498db" if d == 0 else ("#2ecc71" if d > 0 else "#e74c3c") for d in deltas]
    
    fig = go.Figure(go.Bar(
        x=names,
        y=values,
        marker_color=colors,
        text=[f"€{v:,.0f}" for v in values],
        textposition="outside",
    ))
    
    fig.update_layout(
        title=f"Scenario Comparison ({label})",
        xaxis_title="Scenario",
        yaxis_title=f"{label} (€)",
        template="plotly_white",
        showlegend=False,
    )
    
    return fig


def save_biometano_charts(
    projections: BiometanoProjections,
    valuation: ValuationOutputs,
    output_dir: Path,
    sensitivity: Optional[SensitivityAnalysisOutputs] = None,
    methodology: str = "enterprise",
) -> list[Path]:
    """Save all biometano charts to files.
    
    No valuation waterfall/bridge chart per requirements.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    files = []
    
    # Revenue stack (correct column order)
    fig = create_revenue_stack_chart(projections)
    path = output_dir / "revenue_stack.html"
    fig.write_html(str(path))
    files.append(path)
    
    # OPEX stack
    fig = create_opex_stack_chart(projections)
    path = output_dir / "opex_stack.html"
    fig.write_html(str(path))
    files.append(path)
    
    # FCFF timeline
    fig = create_fcff_timeline_chart(projections)
    path = output_dir / "fcff_timeline.html"
    fig.write_html(str(path))
    files.append(path)
    
    # PV decomposition (replaces valuation waterfall)
    fig = create_pv_decomposition_chart(valuation)
    path = output_dir / "pv_decomposition.html"
    fig.write_html(str(path))
    files.append(path)
    
    # Sensitivity charts
    if sensitivity:
        fig = create_tornado_chart(sensitivity, methodology)
        path = output_dir / "sensitivity_tornado.html"
        fig.write_html(str(path))
        files.append(path)
        
        fig = create_scenario_comparison_chart(sensitivity, methodology)
        path = output_dir / "scenario_comparison.html"
        fig.write_html(str(path))
        files.append(path)
    
    return files


def show_biometano_charts(
    projections: BiometanoProjections,
    valuation: ValuationOutputs,
    sensitivity: Optional[SensitivityAnalysisOutputs] = None,
    methodology: str = "enterprise",
) -> None:
    """Show biometano charts in browser.
    
    No valuation waterfall/bridge chart per requirements.
    """
    create_revenue_stack_chart(projections).show()
    create_opex_stack_chart(projections).show()
    create_fcff_timeline_chart(projections).show()
    create_pv_decomposition_chart(valuation).show()
    
    if sensitivity:
        create_tornado_chart(sensitivity, methodology).show()
        create_scenario_comparison_chart(sensitivity, methodology).show()
