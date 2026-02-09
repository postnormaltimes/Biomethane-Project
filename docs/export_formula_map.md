# Export Formula Map

This document summarizes the execution path and Excel formula layout for the DCF and Biometano exports.

## CLI → Engine → Export Flow

### Generic DCF

1. `dcf_ui_cli.cli:app`
   * `run` / `export` commands call `DCFEngine(inputs).run()`.【F:src/dcf_ui_cli/cli.py†L21-L236】【F:src/dcf_engine/engine.py†L67-L199】
2. `DCFEngine` returns `DCFOutputs` containing projections, cash flows, discounting, and valuation bridge data.【F:src/dcf_engine/engine.py†L88-L199】【F:src/dcf_engine/models.py†L306-L331】
3. `dcf_io.writers.export_xlsx(..., xlsx_mode="formulas")` builds the formula-driven workbook via `_write_dcf_formula_workbook`.【F:src/dcf_io/writers.py†L799-L829】

### Biometano

1. `dcf_ui_cli.biometano_cli:app` commands call:
   * `build_projections` → `build_statements` → `compute_valuation`.【F:src/dcf_ui_cli/biometano_cli.py†L244-L475】【F:src/dcf_projects/biometano/builder.py†L717-L728】【F:src/dcf_projects/biometano/statements.py†L460-L475】【F:src/dcf_projects/biometano/valuation.py†L283-L298】
2. `dcf_io.writers.export_xlsx_biometano(..., xlsx_mode="formulas", case=case)` writes the formula model using `_write_biometano_formula_workbook`.【F:src/dcf_ui_cli/biometano_cli.py†L465-L475】【F:src/dcf_io/writers.py†L1002-L1737】

## Sheet Mapping (Formula Exports)

### DCF Formula Workbook

| Sheet | Source + Notes |
| --- | --- |
| `Assumptions` | Input series copied from `DCFOutputs` for revenue, OPEX, D&A, capex, NWC, tax, WACC, etc. |
| `Revenue_By_Channel` | Links to `Assumptions` revenue series. |
| `OPEX` | Links to `Assumptions` OPEX and computes EBITDA. |
| `Income_Statement` | Formulas for EBITDA → EBIT → taxes → net income. |
| `Balance_Sheet` | NWC (revenue-driven), ΔNWC, debt/equity links. |
| `Balance_Sheet_Reclass` | CIN / NFP / Equity roll-up and identity checks. |
| `Cash_Flow` | FCFF/FCFE formulas driven off income statement and assumptions. |
| `FCFF` | Summary of FCFF/FCFE from `Cash_Flow`. |
| `Discounting` | DF, PV, and sums computed with formulas. |
| `Valuation_Summary` | TV, PV TV, EV, equity value formulas. |
| `Audit_Checks` | Formula-driven identity checks (FCFF, CIN, PV roll-up). |
| `Audit_Notes` | Formula coverage + exceptions and conventions. |

### Biometano Formula Workbook

| Sheet | Source + Notes |
| --- | --- |
| `Assumptions` | Case parameters + series inputs (production, OPEX categories, balance sheet components). |
| `Production` | Availability + throughput-driven volumes. |
| `Revenue_By_Channel` | Volume × price × escalation formulas. |
| `OPEX` | Category links + totals. |
| `Income_Statement` | Revenue → EBITDA → EBIT → taxes → net income formulas. |
| `Balance_Sheet` | Component links + total formulas (AR/AP via DSO/DPO). |
| `Balance_Sheet_Reclass` | CIN / NFP / Equity roll-up and identity checks. |
| `Cash_Flow` | CFO/CFI/CFF formulas using assumptions + income statement. |
| `FCFF` | FCFF/FCFE formulas and net borrowing link. |
| `Discounting` | DF/PV formulas based on WACC/Ke. |
| `Valuation_Summary` | TV, PV TV, EV, equity, reconciliation formulas. |
| `Audit_Checks` | Formula-driven identity checks (BS, CIN, cash, FCFF, PV). |
| `Audit_Notes` | Formula coverage + exceptions and conventions. |
