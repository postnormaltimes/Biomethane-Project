"""
DCF I/O Writers

Export to XLSX and CSV formats.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

from dcf_engine.models import DCFOutputs


def _create_inputs_summary(outputs: DCFOutputs) -> pd.DataFrame:
    """Create inputs/assumptions summary table."""
    data = [
        ["Base Year", outputs.base_year],
        ["Forecast Years", ", ".join(map(str, outputs.forecast_years))],
        ["Discounting Mode", outputs.discounting_mode.value],
        ["Cost of Equity (Ke)", f"{outputs.ke:.4f}"],
        ["Terminal Value Method", outputs.terminal_value.method.value],
    ]
    
    if outputs.terminal_value.growth_rate is not None:
        data.append(["Terminal Growth Rate", f"{outputs.terminal_value.growth_rate:.4f}"])
    
    return pd.DataFrame(data, columns=["Parameter", "Value"])


def _create_projections_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create operating projections table."""
    rows = []
    for p in outputs.projections:
        rows.append({
            "Year": p.year,
            "Revenue": p.revenue,
            "Operating Costs": p.operating_costs,
            "EBITDA": p.ebitda,
            "D&A": p.depreciation_amortization,
            "EBIT": p.ebit,
            "Tax on EBIT": p.tax_on_ebit,
            "NOPAT": p.nopat,
            "Interest Expense": p.interest_expense,
            "EBT": p.ebt,
            "Taxes on EBT": p.taxes_on_ebt,
            "Net Income": p.net_income,
        })
    return pd.DataFrame(rows)


def _create_nwc_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create NWC schedule table."""
    rows = []
    for n in outputs.nwc_schedule:
        rows.append({
            "Year": n.year,
            "NWC": n.nwc,
            "ΔNWC": n.delta_nwc,
        })
    return pd.DataFrame(rows)


def _create_investments_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create investment schedule table."""
    rows = []
    for cf in outputs.cash_flows:
        rows.append({
            "Year": cf.year,
            "Capex": cf.capex,
        })
    return pd.DataFrame(rows)


def _create_cashflows_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create cash flows table."""
    rows = []
    for cf in outputs.cash_flows:
        rows.append({
            "Year": cf.year,
            "NOPAT": cf.nopat,
            "D&A": cf.depreciation_amortization,
            "ΔNWC": cf.delta_nwc,
            "Capex": cf.capex,
            "FCFF": cf.fcff,
            "Interest Expense": cf.interest_expense,
            "Interest Tax Shield": cf.interest_tax_shield,
            "Net Borrowing": cf.net_borrowing,
            "FCFE": cf.fcfe,
        })
    return pd.DataFrame(rows)


def _create_discount_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create discount factors and PV table."""
    rows = []
    for d in outputs.discount_schedule:
        rows.append({
            "Year": d.year,
            "Period": d.period,
            "WACC": d.wacc,
            "Ke": d.ke,
            "DF (WACC)": d.discount_factor_wacc,
            "DF (Ke)": d.discount_factor_ke,
            "FCFF": d.fcff,
            "PV(FCFF)": d.pv_fcff,
            "FCFE": d.fcfe,
            "PV(FCFE)": d.pv_fcfe,
        })
    
    # Add totals row
    sum_pv_fcff = sum(d.pv_fcff for d in outputs.discount_schedule)
    sum_pv_fcfe = sum(d.pv_fcfe for d in outputs.discount_schedule)
    rows.append({
        "Year": "Total",
        "Period": "",
        "WACC": "",
        "Ke": "",
        "DF (WACC)": "",
        "DF (Ke)": "",
        "FCFF": "",
        "PV(FCFF)": sum_pv_fcff,
        "FCFE": "",
        "PV(FCFE)": sum_pv_fcfe,
    })
    
    return pd.DataFrame(rows)


def _create_terminal_value_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create terminal value computation table."""
    tv = outputs.terminal_value
    data = [
        ["Final Year", tv.final_year],
        ["Method", tv.method.value],
        ["Growth Rate (g)", tv.growth_rate if tv.growth_rate is not None else "N/A"],
        ["Exit Multiple", tv.exit_multiple if tv.exit_multiple is not None else "N/A"],
        ["Exit Metric", tv.exit_metric if tv.exit_metric else "N/A"],
        ["---", "---"],
        ["Terminal Value (FCFF)", tv.terminal_value_fcff],
        ["Terminal Value (FCFE)", tv.terminal_value_fcfe],
        ["Discount Rate (WACC)", tv.discount_rate_wacc],
        ["Discount Rate (Ke)", tv.discount_rate_ke],
        ["PV of TV (FCFF)", tv.pv_terminal_value_fcff],
        ["PV of TV (FCFE)", tv.pv_terminal_value_fcfe],
    ]
    return pd.DataFrame(data, columns=["Item", "Value"])


def _create_valuation_bridge_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create valuation bridge table."""
    vb = outputs.valuation_bridge
    data = [
        ["FCFF/WACC Approach", ""],
        ["Sum PV(FCFF)", vb.sum_pv_fcff],
        ["PV(Terminal Value)", vb.pv_terminal_value_fcff],
        ["Enterprise Value", vb.enterprise_value],
        ["Less: Debt at Base", -vb.debt_at_base],
        ["Plus: Cash at Base", vb.cash_at_base],
        ["Net Debt", vb.net_debt],
        ["Equity Value (from EV)", vb.equity_value_from_ev],
        ["---", "---"],
        ["FCFE/Ke Approach", ""],
        ["Sum PV(FCFE)", vb.sum_pv_fcfe],
        ["PV(Terminal Value)", vb.pv_terminal_value_fcfe],
        ["Equity Value (Direct)", vb.equity_value_direct],
        ["---", "---"],
        ["Reconciliation", ""],
        ["Difference", vb.reconciliation_difference],
    ]
    return pd.DataFrame(data, columns=["Item", "Value"])


def _create_reconciliation_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create FCFF vs FCFE reconciliation report."""
    vb = outputs.valuation_bridge
    
    data = [
        ["Equity from FCFF/WACC", vb.equity_value_from_ev],
        ["Equity from FCFE/Ke", vb.equity_value_direct],
        ["Difference", vb.reconciliation_difference],
        ["Difference (%)", abs(vb.reconciliation_difference / vb.equity_value_from_ev * 100) if vb.equity_value_from_ev else 0],
    ]
    
    # Add notes
    for i, note in enumerate(vb.reconciliation_notes):
        data.append([f"Note {i+1}", note])
    
    return pd.DataFrame(data, columns=["Item", "Value"])


def _create_wacc_details_table(outputs: DCFOutputs) -> pd.DataFrame:
    """Create WACC details table."""
    rows = []
    for w in outputs.wacc_details:
        rows.append({
            "Year": w.year,
            "Ke": w.ke,
            "Rd": w.rd,
            "Tax Rate": w.tax_rate,
            "Debt": w.debt,
            "Equity Book": w.equity_book,
            "Weight Debt": w.weight_debt,
            "Weight Equity": w.weight_equity,
            "WACC": w.wacc,
        })
    return pd.DataFrame(rows)


def format_tables(outputs: DCFOutputs) -> dict[str, pd.DataFrame]:
    """
    Convert DCF outputs to display-ready DataFrames.
    
    Returns:
        Dict mapping table name to DataFrame
    """
    return {
        "1_Inputs_Summary": _create_inputs_summary(outputs),
        "2_Operating_Projections": _create_projections_table(outputs),
        "3_NWC_Schedule": _create_nwc_table(outputs),
        "4_Investments": _create_investments_table(outputs),
        "5_Cash_Flows": _create_cashflows_table(outputs),
        "6_WACC_Details": _create_wacc_details_table(outputs),
        "7_PV_Decomposition": _create_discount_table(outputs),
        "8_Terminal_Value": _create_terminal_value_table(outputs),
        "9_Valuation_Bridge": _create_valuation_bridge_table(outputs),
        "10_Reconciliation": _create_reconciliation_table(outputs),
    }


def _style_xlsx_sheet(ws, df: pd.DataFrame):
    """Apply styling to Excel worksheet."""
    # Header style
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Style header row
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")
    
    # Style data cells
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00' if isinstance(cell.value, float) else '#,##0'
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 40)
        ws.column_dimensions[column_letter].width = adjusted_width


def export_xlsx(outputs: DCFOutputs, path: str | Path) -> None:
    """
    Export DCF outputs to Excel file with one sheet per table.
    
    Args:
        outputs: DCF outputs to export
        path: Output file path
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tables = format_tables(outputs)
    
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    for sheet_name, df in tables.items():
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel limits sheet names to 31 chars
        
        # Write data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Apply styling
        _style_xlsx_sheet(ws, df)
    
    wb.save(path)


def export_csv(outputs: DCFOutputs, output_dir: str | Path) -> list[Path]:
    """
    Export DCF outputs to CSV files (one per table).
    
    Args:
        outputs: DCF outputs to export
        output_dir: Directory to write CSV files
    
    Returns:
        List of created file paths
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    tables = format_tables(outputs)
    created_files = []
    
    for table_name, df in tables.items():
        file_path = output_dir / f"{table_name}.csv"
        df.to_csv(file_path, index=False)
        created_files.append(file_path)
    
    return created_files


# ============================================================================
# BIOMETANO EXPORTS
# ============================================================================

def _create_biometano_production_table(projections) -> pd.DataFrame:
    """Create production table for biometano."""
    rows = []
    for p in projections.production:
        rows.append({
            "Year": p.year,
            "Availability": p.availability,
            "FORSU (t)": p.forsu_tonnes,
            "Biomethane (MWh)": p.biomethane_mwh,
            "CO2 (t)": p.co2_tonnes,
            "Compost (t)": p.compost_tonnes,
        })
    return pd.DataFrame(rows)


def _create_biometano_revenue_table(projections) -> pd.DataFrame:
    """Create revenue by channel table."""
    rows = []
    for r in projections.revenues:
        rows.append({
            "Year": r.year,
            "Gate Fee": r.gate_fee,
            "Tariff": r.tariff,
            "CO2": r.co2,
            "GO": r.go,
            "Compost": r.compost,
            "Total": r.total,
        })
    return pd.DataFrame(rows)


def _create_biometano_opex_table(projections) -> pd.DataFrame:
    """Create OPEX by category table."""
    rows = []
    for o in projections.opex:
        rows.append({
            "Year": o.year,
            "Feedstock": o.feedstock_handling,
            "Utilities": o.utilities,
            "Chemicals": o.chemicals,
            "Maintenance": o.maintenance,
            "Personnel": o.personnel,
            "Insurance": o.insurance,
            "Overheads": o.overheads,
            "Digestate": o.digestate_handling,
            "Other": o.other,
            "Total": o.total,
        })
    return pd.DataFrame(rows)


def _create_biometano_income_statement(statements) -> pd.DataFrame:
    """Create income statement table."""
    rows = []
    for s in statements.income_statements:
        rows.append({
            "Year": s.year,
            "Revenue": s.total_revenue,
            "OPEX": s.total_opex,
            "EBITDA": s.ebitda,
            "D&A": s.depreciation,
            "Grant Release": s.grant_income_release,
            "EBIT": s.ebit,
            "Interest": s.interest_expense,
            "EBT": s.ebt,
            "Taxes Before Credit": s.taxes_before_credit,
            "Tax Credit Used": s.tax_credit_utilization,
            "Taxes Paid": s.taxes_paid,
            "Net Income": s.net_income,
        })
    return pd.DataFrame(rows)


def _create_biometano_cash_flow(statements) -> pd.DataFrame:
    """Create cash flow statement table."""
    rows = []
    for c in statements.cash_flows:
        rows.append({
            "Year": c.year,
            "NOPAT": c.nopat,
            "D&A": c.depreciation,
            "Change in NWC": c.change_in_nwc,
            "CFO": c.cfo,
            "CAPEX": c.capex,
            "Grant Cash": c.grant_cash_received,
            "CFI": c.cfi,
            "Debt Draw": c.debt_drawdown,
            "Debt Repay": c.debt_repayment,
            "Interest Paid": c.interest_paid,
            "CFF": c.cff,
            "Net CF": c.net_cash_flow,
            "Closing Cash": c.closing_cash,
        })
    return pd.DataFrame(rows)


def _create_biometano_fcff_table(projections) -> pd.DataFrame:
    """Create FCFF breakdown table."""
    rows = []
    for year in projections.operating_years:
        rows.append({
            "Year": year,
            "EBIT": projections.ebit.get(year, 0),
            "D&A": projections.depreciation.get(year, 0),
            "ΔNWC": projections.delta_nwc.get(year, 0),
            "FCFF": projections.fcff.get(year, 0),
            "FCFE": projections.fcfe.get(year, 0),
        })
    return pd.DataFrame(rows)


def _create_biometano_valuation_table(valuation) -> pd.DataFrame:
    """Create valuation summary table."""
    data = [
        ["Cost of Equity (Ke)", valuation.ke],
        ["Sum PV(FCFF)", valuation.sum_pv_fcff],
        ["Terminal Value (FCFF)", valuation.terminal_value_fcff],
        ["PV Terminal Value", valuation.pv_terminal_value_fcff],
        ["Enterprise Value", valuation.enterprise_value],
        ["Debt at Base", valuation.debt_at_base],
        ["Cash at Base", valuation.cash_at_base],
        ["Net Debt", valuation.net_debt],
        ["Equity Value (FCFF/WACC)", valuation.equity_value],
        ["Equity Value (FCFE/Ke)", valuation.equity_value_direct],
        ["Reconciliation Difference", valuation.reconciliation_difference],
    ]
    return pd.DataFrame(data, columns=["Item", "Value"])


def format_biometano_tables(projections, statements, valuation) -> dict[str, pd.DataFrame]:
    """Format all biometano tables for export."""
    return {
        "1_Production": _create_biometano_production_table(projections),
        "2_Revenue": _create_biometano_revenue_table(projections),
        "3_OPEX": _create_biometano_opex_table(projections),
        "4_Income_Statement": _create_biometano_income_statement(statements),
        "5_Cash_Flow": _create_biometano_cash_flow(statements),
        "6_FCFF_FCFE": _create_biometano_fcff_table(projections),
        "7_Valuation": _create_biometano_valuation_table(valuation),
    }


def export_xlsx_biometano(projections, statements, valuation, path: str | Path) -> None:
    """Export biometano outputs to Excel."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tables = format_biometano_tables(projections, statements, valuation)
    
    wb = Workbook()
    wb.remove(wb.active)
    
    for sheet_name, df in tables.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        _style_xlsx_sheet(ws, df)
    
    wb.save(path)


def export_csv_biometano(projections, statements, valuation, output_dir: str | Path) -> list[Path]:
    """Export biometano outputs to CSV files."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    tables = format_biometano_tables(projections, statements, valuation)
    created_files = []
    
    for table_name, df in tables.items():
        file_path = output_dir / f"{table_name}.csv"
        df.to_csv(file_path, index=False)
        created_files.append(file_path)
    
    return created_files
