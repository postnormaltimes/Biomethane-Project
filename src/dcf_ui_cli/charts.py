"""
DCF Visualizations

Plotly charts for DCF analysis.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import plotly.graph_objects as go
import plotly.express as px
import numpy as np

from dcf_engine.models import DCFOutputs


def create_waterfall_chart(outputs: DCFOutputs) -> go.Figure:
    """
    Create PV composition waterfall chart.
    
    Shows: PV(Flows) + PV(TV) → EV → Equity
    """
    vb = outputs.valuation_bridge
    
    # Waterfall data
    labels = [
        "PV(FCFF Flows)",
        "PV(Terminal Value)",
        "Enterprise Value",
        "Less: Net Debt",
        "Equity Value",
    ]
    
    values = [
        vb.sum_pv_fcff,
        vb.pv_terminal_value_fcff,
        0,  # Subtotal
        -vb.net_debt,
        0,  # Total
    ]
    
    measure = ["relative", "relative", "total", "relative", "total"]
    
    fig = go.Figure(go.Waterfall(
        name="Valuation Bridge",
        orientation="v",
        measure=measure,
        x=labels,
        textposition="outside",
        text=[f"{v:,.0f}" if v != 0 else "" for v in values],
        y=values,
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#2E86AB"}},
        decreasing={"marker": {"color": "#E94F37"}},
        totals={"marker": {"color": "#44AF69"}},
    ))
    
    fig.update_layout(
        title="DCF Valuation Waterfall",
        showlegend=False,
        yaxis_title="Value",
        template="plotly_white",
        height=500,
    )
    
    return fig


def create_cashflow_timeline(outputs: DCFOutputs) -> go.Figure:
    """
    Create cash flow timeline chart.
    
    Shows FCFF and FCFE by year as bar chart.
    """
    years = [cf.year for cf in outputs.cash_flows]
    fcff = [cf.fcff for cf in outputs.cash_flows]
    fcfe = [cf.fcfe for cf in outputs.cash_flows]
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name="FCFF",
        x=years,
        y=fcff,
        marker_color="#2E86AB",
        text=[f"{v:,.0f}" for v in fcff],
        textposition="outside",
    ))
    
    fig.add_trace(go.Bar(
        name="FCFE",
        x=years,
        y=fcfe,
        marker_color="#44AF69",
        text=[f"{v:,.0f}" for v in fcfe],
        textposition="outside",
    ))
    
    fig.update_layout(
        title="Free Cash Flow Timeline",
        xaxis_title="Year",
        yaxis_title="Cash Flow",
        barmode="group",
        template="plotly_white",
        height=400,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
    )
    
    return fig


def create_sensitivity_heatmap(
    outputs: DCFOutputs,
    wacc_range: Optional[list[float]] = None,
    growth_range: Optional[list[float]] = None,
) -> go.Figure:
    """
    Create sensitivity heatmap for WACC vs Growth Rate.
    
    Shows how Equity Value changes with different assumptions.
    """
    # Get base values
    base_wacc = outputs.wacc_details[-1].wacc  # Final year WACC
    base_g = outputs.terminal_value.growth_rate or 0.0
    final_fcff = outputs.cash_flows[-1].fcff
    sum_pv_fcff = outputs.valuation_bridge.sum_pv_fcff
    net_debt = outputs.valuation_bridge.net_debt
    n_periods = len(outputs.forecast_years)
    
    # Default ranges
    if wacc_range is None:
        wacc_range = [
            base_wacc - 0.02,
            base_wacc - 0.01,
            base_wacc,
            base_wacc + 0.01,
            base_wacc + 0.02,
        ]
    
    if growth_range is None:
        growth_range = [
            max(0, base_g - 0.01),
            base_g,
            base_g + 0.01,
            base_g + 0.02,
        ]
    
    # Compute sensitivity matrix
    equity_values = []
    for g in growth_range:
        row = []
        for wacc in wacc_range:
            if g >= wacc:
                row.append(float('inf'))
            else:
                if g == 0:
                    tv = final_fcff / wacc
                else:
                    tv = final_fcff * (1 + g) / (wacc - g)
                pv_tv = tv / ((1 + wacc) ** n_periods)
                ev = sum_pv_fcff + pv_tv
                equity = ev - net_debt
                row.append(equity)
        equity_values.append(row)
    
    # Create heatmap
    fig = go.Figure(data=go.Heatmap(
        z=equity_values,
        x=[f"{w:.2%}" for w in wacc_range],
        y=[f"{g:.2%}" for g in growth_range],
        colorscale="RdYlGn",
        text=[[f"{v:,.0f}" if v != float('inf') else "N/A" for v in row] for row in equity_values],
        texttemplate="%{text}",
        textfont={"size": 10},
        hovertemplate="WACC: %{x}<br>Growth: %{y}<br>Equity: %{text}<extra></extra>",
    ))
    
    fig.update_layout(
        title="Equity Value Sensitivity: WACC vs Growth Rate",
        xaxis_title="WACC",
        yaxis_title="Growth Rate (g)",
        template="plotly_white",
        height=400,
    )
    
    return fig


def create_pv_composition_pie(outputs: DCFOutputs) -> go.Figure:
    """
    Create pie chart showing PV composition.
    """
    vb = outputs.valuation_bridge
    
    labels = ["PV(FCFF Flows)", "PV(Terminal Value)"]
    values = [vb.sum_pv_fcff, vb.pv_terminal_value_fcff]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.4,
        marker_colors=["#2E86AB", "#44AF69"],
        textinfo="percent+label",
        textposition="outside",
    )])
    
    fig.update_layout(
        title="Enterprise Value Composition",
        template="plotly_white",
        height=400,
        annotations=[dict(
            text=f"EV<br>{vb.enterprise_value:,.0f}",
            x=0.5, y=0.5,
            font_size=14,
            showarrow=False,
        )],
    )
    
    return fig


def save_charts(
    outputs: DCFOutputs,
    output_dir: str | Path,
    format: str = "html",
) -> list[Path]:
    """
    Generate and save all charts.
    
    Args:
        outputs: DCF outputs
        output_dir: Directory to save charts
        format: Output format ("html", "png", "svg")
    
    Returns:
        List of created file paths
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    charts = {
        "waterfall": create_waterfall_chart(outputs),
        "cashflow_timeline": create_cashflow_timeline(outputs),
        "sensitivity_heatmap": create_sensitivity_heatmap(outputs),
    }
    
    created_files = []
    for name, fig in charts.items():
        file_path = output_dir / f"{name}.{format}"
        if format == "html":
            fig.write_html(str(file_path))
        else:
            fig.write_image(str(file_path))
        created_files.append(file_path)
    
    return created_files


def show_charts(outputs: DCFOutputs) -> None:
    """
    Display all charts (opens in browser).
    """
    charts = [
        create_waterfall_chart(outputs),
        create_cashflow_timeline(outputs),
        create_sensitivity_heatmap(outputs),
    ]
    
    for chart in charts:
        chart.show()
