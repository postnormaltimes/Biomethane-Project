# DCF Valuation Exercise: NordSilica S.r.l.

**Course**: Financial Valuation and Modeling  
**Exercise Type**: Unlevered DCF (FCFF → Enterprise Value → Equity Value)  
**Units**: All figures in **€ million** unless stated otherwise

---

## Scenario

A private equity analyst is valuing **NordSilica S.r.l.**, an unlisted Italian specialty chemicals company, at the **end of 2023** (valuation date: 31 December 2023).

The forecast horizon covers **2024–2027** (4 explicit years). All cash flows are assumed to occur at year-end. Beyond 2027, the company is assumed to grow in perpetuity at a constant rate.

**Key assumptions**:
- No dividends are paid during 2024–2027
- No equity injections or capital increases
- Shareholders' Equity evolves only via retained Net Income

---

## Input Data

### A. Base Year Data (End of 2023)

| Item | Value (€M) |
|------|------------|
| Revenues 2023 | 9,800 |
| NWC as % of Revenues (2023) | 14% |
| Fixed Assets (Net PP&E + Intangibles) | 6,200 |
| Cash & Equivalents | 950 |
| Financial Debt (gross) | 1,200 |
| Shareholders' Equity (book value) | 5,000 |

---

### B. Operating Assumptions (2024–2027)

#### Revenue Growth (Year-over-Year)

| Year | Growth Rate |
|------|-------------|
| 2024 | +12% |
| 2025 | +9% |
| 2026 | +7% |
| 2027 | +5% |

#### Operating Costs as % of Revenues

| Year | Cost % |
|------|--------|
| 2024 | 86% |
| 2025 | 84% |
| 2026 | 82% |
| 2027 | 81% |

> *Note: EBITDA = Revenues × (1 − Cost%)*

#### Depreciation & Amortization (€M)

| Year | D&A |
|------|-----|
| 2024 | 450 |
| 2025 | 520 |
| 2026 | 560 |
| 2027 | 590 |

#### Net CAPEX (€M)

| Year | CAPEX |
|------|-------|
| 2024 | 650 |
| 2025 | 720 |
| 2026 | 820 |
| 2027 | 900 |

#### NWC as % of Revenues

| Year | NWC % |
|------|-------|
| 2024 | 15% |
| 2025 | 13% |
| 2026 | 12% |
| 2027 | 11% |

---

### C. Debt Schedule and Cost of Debt

#### Financial Debt (end of year, €M)

| Year | Debt |
|------|------|
| 2023 | 1,200 |
| 2024 | 1,550 |
| 2025 | 1,600 |
| 2026 | 1,550 |
| 2027 | 1,500 |

#### Cost of Debt (annual rate)

| Year | rd |
|------|-----|
| 2023 | 4.8% |
| 2024 | 5.6% |
| 2025 | 6.1% |
| 2026 | 6.2% |
| 2027 | 6.2% |

> *Convention: Interest expense = End-of-year Debt × Cost of Debt for that year*

---

### D. Tax and Discount Rate Parameters

| Parameter | Value |
|-----------|-------|
| Corporate Tax Rate (t) | 28% |
| Risk-free Rate (rf) | 3.5% |
| Expected Market Return (rm) | 9.0% |
| Equity Beta (β) | 1.10 |
| Terminal Growth Rate (g) | 2.0% |

---

## Required Tasks

Complete the following analysis using the **Unlevered DCF (FCFF)** approach:

### Task 1: Operating Projections (2024–2027)
Build a projection table showing for each year:
- Revenues
- EBITDA (using: EBITDA = Revenues × (1 − Cost%))
- EBIT (using: EBIT = EBITDA − D&A)

### Task 2: Working Capital Schedule
Calculate for each year:
- NWC (= Revenues × NWC%)
- ΔNWC (= NWC_t − NWC_{t-1})

### Task 3: Free Cash Flow to Firm (FCFF)
Compute FCFF for each year using:
$$\text{FCFF}_t = \text{EBIT}_t \times (1 - t) + \text{D\&A}_t - \Delta\text{NWC}_t - \text{CAPEX}_t$$

### Task 4: Income Statement & Equity Roll-Forward
For Net Income calculation:
- Interest_t = Debt_t × rd_t
- EBT_t = EBIT_t − Interest_t
- Taxes_t = EBT_t × t
- Net Income_t = EBT_t − Taxes_t

Roll forward Shareholders' Equity:
$$\text{Equity}_t = \text{Equity}_{t-1} + \text{Net Income}_t$$

### Task 5: WACC Calculation
For each year, compute:
- Cost of Equity: Ke = rf + β × (rm − rf)
- Weights (book-value based):
  - wD_t = Debt_t / (Debt_t + Equity_t)
  - wE_t = Equity_t / (Debt_t + Equity_t)
- WACC_t = wE_t × Ke + wD_t × rd_t × (1 − t)

### Task 6: Present Value of FCFFs
Use the **year-specific-flat discounting convention**:
- PV(FCFF_2024) = FCFF_2024 / (1 + WACC_2024)^1
- PV(FCFF_2025) = FCFF_2025 / (1 + WACC_2025)^2
- PV(FCFF_2026) = FCFF_2026 / (1 + WACC_2026)^3
- PV(FCFF_2027) = FCFF_2027 / (1 + WACC_2027)^4

### Task 7: Terminal Value
Calculate Terminal Value at end of 2027:
$$\text{TV}_{2027} = \frac{\text{FCFF}_{2027} \times (1 + g)}{\text{WACC}_{2027} - g}$$

Then discount to valuation date:
$$\text{PV(TV)} = \frac{\text{TV}_{2027}}{(1 + \text{WACC}_{2027})^4}$$

### Task 8: Enterprise Value and Equity Value
- Enterprise Value = Sum of PV(FCFFs) + PV(TV)
- Net Financial Position at 2023 = Debt_2023 − Cash_2023
- Equity Value = Enterprise Value − NFP_2023

---

## Deliverables

1. **Projection Table**: Revenues, EBITDA, EBIT, D&A (2024–2027)
2. **NWC Schedule**: NWC and ΔNWC by year
3. **FCFF Schedule**: FCFF calculation by year
4. **Income Statement Summary**: Interest, EBT, Taxes, Net Income by year
5. **Equity Roll-Forward**: Shareholders' Equity by year
6. **WACC Table**: Weights and WACC by year
7. **Valuation Summary**: PV of FCFFs, Terminal Value, EV, NFP, Equity Value
8. **Interpretation**: Brief commentary on key value drivers

---

*End of Exercise*
