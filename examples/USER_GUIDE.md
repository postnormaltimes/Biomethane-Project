# DCF Modeling Tool — User Guide

Complete step-by-step instructions to solve any DCF valuation exercise from the Terminal.

---

## Part 1: One-Time Setup

### Step 1: Navigate to the Project

```bash
cd /Users/emilianobarin/.gemini/antigravity/playground/pyro-prominence
```

### Step 2: Install the Tool (if not already installed)

```bash
pip install -e ".[dev]"
```

### Step 3: Verify Installation

```bash
dcf --help
```

Expected output:
```
Usage: dcf [OPTIONS] COMMAND [ARGS]...

  Production-quality DCF modeling tool

Options:
  --help  Show this message and exit.

Commands:
  export    Run DCF and export results to files.
  run       Run DCF analysis on an input file.
  validate  Validate an input file without running the full DCF.
```

---

## Part 2: Creating an Input File for a New Exercise

### Step 1: Copy the Template

```bash
cp examples/example_input.yaml examples/my_exercise.yaml
```

### Step 2: Edit the Input File

Open `examples/my_exercise.yaml` in any text editor and fill in your exercise data:

```yaml
# ============================================================
# SECTION 1: DISCOUNTING MODE
# ============================================================
# Options: "year_specific_flat" (MEM-style) or "constant"
discounting_mode: "year_specific_flat"

# ============================================================
# SECTION 2: TIMELINE
# ============================================================
timeline:
  base_year: 2023                        # Valuation date (end of this year)
  forecast_years: [2024, 2025, 2026, 2027]  # Explicit forecast years

# ============================================================
# SECTION 3: REVENUE
# ============================================================
revenue:
  # Option A: Base revenue + growth rates
  base_revenue: 9800.0                   # Revenue at base year
  growth_rates:
    2024: 0.12                           # +12% growth
    2025: 0.09                           # +9% growth
    2026: 0.07                           # +7% growth
    2027: 0.05                           # +5% growth

  # Option B: Explicit revenue (overrides growth rates if provided)
  # explicit_revenue:
  #   2024: 10976.0
  #   2025: 11963.84
  #   ...

# ============================================================
# SECTION 4: OPERATING COSTS & D&A
# ============================================================
operating:
  # Operating costs as % of revenue (EBITDA margin = 1 - cost%)
  cost_ratios:
    2024: 0.86                           # 86% costs → 14% EBITDA margin
    2025: 0.84
    2026: 0.82
    2027: 0.81

  # D&A by year (include base year for context)
  depreciation_amortization:
    2023: 420.0
    2024: 450.0
    2025: 520.0
    2026: 560.0
    2027: 590.0

# ============================================================
# SECTION 5: NET WORKING CAPITAL
# ============================================================
nwc:
  # NWC as % of revenue (include base year!)
  nwc_percent:
    2023: 0.14                           # 14% of revenue
    2024: 0.15
    2025: 0.13
    2026: 0.12
    2027: 0.11

# ============================================================
# SECTION 6: INVESTMENTS
# ============================================================
investments:
  capex:
    2024: 650.0                          # Positive = cash outflow
    2025: 720.0
    2026: 820.0
    2027: 900.0

# ============================================================
# SECTION 7: TAX
# ============================================================
tax:
  tax_rate: 0.28                         # 28% corporate tax

# ============================================================
# SECTION 8: CAPM (Cost of Equity)
# ============================================================
capm:
  rf: 0.035                              # Risk-free rate (3.5%)
  rm: 0.09                               # Market return (9.0%)
  beta: 1.10                             # Equity beta

# ============================================================
# SECTION 9: DEBT
# ============================================================
debt:
  # Debt balances by year (include base year!)
  debt_balances:
    2023: 1200.0
    2024: 1550.0
    2025: 1600.0
    2026: 1550.0
    2027: 1500.0

  # Cost of debt by year
  rd:
    2023: 0.048                          # 4.8%
    2024: 0.056                          # 5.6%
    2025: 0.061
    2026: 0.062
    2027: 0.062

# ============================================================
# SECTION 10: WACC CONFIGURATION
# ============================================================
wacc:
  weighting_mode: "book_value"           # "book_value" or "target"

  # For book_value mode:
  equity_book_inputs:
    base_equity_book: 5000.0             # Shareholders' equity at base year

  # For target mode (uncomment if using):
  # wE: 0.80                             # Target equity weight
  # wD: 0.20                             # Target debt weight

# ============================================================
# SECTION 11: TERMINAL VALUE
# ============================================================
terminal_value:
  method: "perpetuity"                   # "perpetuity" or "exit_multiple"
  g: 0.02                                # Perpetuity growth rate (2%)

  # For exit_multiple method (uncomment if using):
  # method: "exit_multiple"
  # exit_multiple: 8.0
  # exit_metric: "ebitda"                # or "ebit", "revenue"

# ============================================================
# SECTION 12: NET DEBT COMPONENTS
# ============================================================
net_debt:
  cash_and_equivalents: 950.0            # Cash at base year
```

---

## Part 3: Running the Analysis

### Option A: Display Results in Terminal

```bash
dcf run examples/my_exercise.yaml
```

This displays all 9 tables:
1. 📊 Inputs Summary
2. 📈 Operating Projections (Revenue, EBITDA, EBIT)
3. 💰 Net Working Capital Schedule
4. 💸 Cash Flows (FCFF and FCFE)
5. 📊 WACC Details
6. 📉 Present Value Decomposition
7. 🎯 Terminal Value
8. 🌉 Valuation Bridge (EV → Equity)
9. 🔄 FCFF vs FCFE Reconciliation

### Option B: Export to Excel

```bash
dcf run examples/my_exercise.yaml -o output/my_results.xlsx
```

Creates Excel file with 10 worksheets (one per table).

### Option C: Export Everything (Excel + CSV + Charts)

```bash
dcf run examples/my_exercise.yaml \
    -o output/my_results.xlsx \
    --csv-dir output/csv \
    --charts-dir output/charts
```

### Option D: Quiet Mode (Exports Only, No Terminal Output)

```bash
dcf run examples/my_exercise.yaml -o output/my_results.xlsx --quiet
```

### Option E: Display Interactive Charts

```bash
dcf run examples/my_exercise.yaml --charts
```

Opens 3 Plotly charts in your browser:
- Waterfall: PV(Flows) + PV(TV) → EV → Equity
- Timeline: FCFF and FCFE by year
- Sensitivity: WACC vs Growth Rate heatmap

---

## Part 4: Validating Before Running

Check if your input file is correctly formatted:

```bash
dcf validate examples/my_exercise.yaml
```

Expected output:
```
Validating: examples/my_exercise.yaml
✓ Input file is valid

  Base year: 2023
  Forecast years: [2024, 2025, 2026, 2027]
  Discounting mode: year_specific_flat
  Terminal value method: perpetuity
  WACC weighting: book_value
```

---

## Part 5: Example Workflow for a Complete Exercise

```bash
# 1. Navigate to project
cd /Users/emilianobarin/.gemini/antigravity/playground/pyro-prominence

# 2. Copy template
cp examples/example_input.yaml examples/exam_q3.yaml

# 3. Edit with your exercise data
nano examples/exam_q3.yaml   # or use any editor

# 4. Validate
dcf validate examples/exam_q3.yaml

# 5. Run and display
dcf run examples/exam_q3.yaml

# 6. Export to Excel for submission
dcf run examples/exam_q3.yaml -o output/exam_q3.xlsx --quiet
```

---

## Part 6: Common Input Patterns

### Pattern 1: Constant Tax Rate
```yaml
tax:
  tax_rate: 0.28
```

### Pattern 2: Variable Tax Rate by Year
```yaml
tax:
  tax_rate:
    2024: 0.28
    2025: 0.27
    2026: 0.25
    2027: 0.25
```

### Pattern 3: Constant Cost of Debt
```yaml
debt:
  rd: 0.06  # 6% for all years
```

### Pattern 4: Exit Multiple Instead of Perpetuity
```yaml
terminal_value:
  method: "exit_multiple"
  exit_multiple: 8.0
  exit_metric: "ebitda"
```

### Pattern 5: Target WACC Weights
```yaml
wacc:
  weighting_mode: "target"
  wE: 0.75
  wD: 0.25
```

---

## Part 7: Troubleshooting

### Error: "Missing growth rates for years: [2024, 2025]"
→ Add all forecast years to `growth_rates` section

### Error: "Missing nwc_percent for years: [2023]"
→ Include base year in `nwc_percent` (needed for ΔNWC calculation)

### Error: "Missing debt_balances for years: [2023]"
→ Include base year in `debt_balances`

### Error: "Growth rate must be less than discount rate"
→ Your perpetuity growth rate `g` is ≥ WACC (reduces to absurd valuation)

---

## Quick Reference Card

```bash
# Run with display
dcf run INPUT.yaml

# Run + Excel export
dcf run INPUT.yaml -o OUTPUT.xlsx

# Run + all exports
dcf run INPUT.yaml -o OUTPUT.xlsx --csv-dir CSV_DIR --charts-dir CHART_DIR

# Validate only
dcf validate INPUT.yaml

# Show help
dcf --help
dcf run --help
```
