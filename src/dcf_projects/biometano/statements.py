"""
Biometano Financial Statements Module

Generates Income Statement, Balance Sheet, and Cash Flow Statement
with proper classification of grants and tax credits per OIC standards.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from dcf_projects.biometano.schema import (
    BiometanoCase,
    GrantAccountingPolicy,
)
from dcf_projects.biometano.builder import BiometanoProjections


@dataclass
class IncomeStatementLine:
    """Single year income statement."""
    year: int
    
    # Revenue by channel
    revenue_gate_fee: float = 0.0
    revenue_tariff: float = 0.0
    revenue_co2: float = 0.0
    revenue_go: float = 0.0
    revenue_compost: float = 0.0
    total_revenue: float = 0.0
    
    # OPEX by category
    opex_feedstock: float = 0.0
    opex_utilities: float = 0.0
    opex_chemicals: float = 0.0
    opex_maintenance: float = 0.0
    opex_personnel: float = 0.0
    opex_insurance: float = 0.0
    opex_overheads: float = 0.0
    opex_digestate: float = 0.0
    opex_other: float = 0.0
    total_opex: float = 0.0
    
    # Operating results
    ebitda: float = 0.0
    depreciation: float = 0.0
    grant_income_release: float = 0.0  # Policy A2
    ebit: float = 0.0
    
    # Financing
    interest_expense: float = 0.0
    ebt: float = 0.0
    
    # Taxes
    taxes_before_credit: float = 0.0
    tax_credit_utilization: float = 0.0
    taxes_paid: float = 0.0
    
    # Bottom line
    net_income: float = 0.0


@dataclass
class BalanceSheetLine:
    """Single year balance sheet."""
    year: int
    
    # Assets
    cash: float = 0.0
    trade_receivables: float = 0.0
    grant_receivable: float = 0.0
    tax_credit_receivable: float = 0.0
    total_current_assets: float = 0.0
    
    fixed_assets_gross: float = 0.0
    accumulated_depreciation: float = 0.0
    fixed_assets_net: float = 0.0
    total_assets: float = 0.0
    
    # Liabilities
    trade_payables: float = 0.0
    taxes_payable: float = 0.0
    total_current_liabilities: float = 0.0
    
    debt: float = 0.0
    deferred_income: float = 0.0  # Policy A2
    total_non_current_liabilities: float = 0.0
    total_liabilities: float = 0.0
    
    # Equity
    share_capital: float = 0.0
    retained_earnings: float = 0.0
    current_year_profit: float = 0.0
    total_equity: float = 0.0
    
    # Check
    total_liabilities_and_equity: float = 0.0
    balance_check: float = 0.0  # Should be zero


@dataclass
class CashFlowLine:
    """Single year cash flow statement."""
    year: int
    
    # Operating activities (indirect method)
    ebit: float = 0.0
    tax_on_ebit: float = 0.0
    nopat: float = 0.0
    depreciation: float = 0.0
    grant_income_release: float = 0.0  # Non-cash, reverse out
    change_in_trade_receivables: float = 0.0
    change_in_trade_payables: float = 0.0
    change_in_nwc: float = 0.0
    taxes_paid_adjustment: float = 0.0  # Difference from accrual
    cfo: float = 0.0
    
    # Investing activities
    capex: float = 0.0
    grant_cash_received: float = 0.0
    cfi: float = 0.0
    
    # Financing activities
    debt_drawdown: float = 0.0
    debt_repayment: float = 0.0
    interest_paid: float = 0.0
    equity_contribution: float = 0.0  # Equity injection for CAPEX funding
    cff: float = 0.0
    
    # Summary
    net_cash_flow: float = 0.0
    opening_cash: float = 0.0
    closing_cash: float = 0.0


@dataclass
class FinancialStatements:
    """Complete financial statements."""
    income_statements: list[IncomeStatementLine] = field(default_factory=list)
    balance_sheets: list[BalanceSheetLine] = field(default_factory=list)
    cash_flows: list[CashFlowLine] = field(default_factory=list)
    
    def get_income_statement(self, year: int) -> Optional[IncomeStatementLine]:
        for stmt in self.income_statements:
            if stmt.year == year:
                return stmt
        return None
    
    def get_balance_sheet(self, year: int) -> Optional[BalanceSheetLine]:
        for stmt in self.balance_sheets:
            if stmt.year == year:
                return stmt
        return None
    
    def get_cash_flow(self, year: int) -> Optional[CashFlowLine]:
        for stmt in self.cash_flows:
            if stmt.year == year:
                return stmt
        return None
    
    def validate_balance_sheets(self) -> list[tuple[int, float]]:
        """Check that balance sheets balance. Returns [(year, imbalance), ...]"""
        errors = []
        for bs in self.balance_sheets:
            if abs(bs.balance_check) > 0.01:
                errors.append((bs.year, bs.balance_check))
        return errors


class StatementsBuilder:
    """
    Builds financial statements from projections.
    
    Handles:
    - Revenue/OPEX allocation to income statement
    - Fixed asset and liability schedules to balance sheet
    - Cash flow classification per OIC standards
    """
    
    def __init__(self, case: BiometanoCase, projections: BiometanoProjections):
        self.case = case
        self.projections = projections
        
    def build(self) -> FinancialStatements:
        """Build all three statements."""
        statements = FinancialStatements()
        
        statements.income_statements = self._build_income_statements()
        statements.cash_flows = self._build_cash_flows()
        statements.balance_sheets = self._build_balance_sheets()
        
        return statements
    
    def _build_income_statements(self) -> list[IncomeStatementLine]:
        """Build income statements from projections."""
        statements = []
        proj = self.projections
        
        for year in proj.all_forecast_years:
            stmt = IncomeStatementLine(year=year)
            
            # Revenue
            rev = proj.get_revenue(year)
            if rev:
                stmt.revenue_gate_fee = rev.gate_fee
                stmt.revenue_tariff = rev.tariff
                stmt.revenue_co2 = rev.co2
                stmt.revenue_go = rev.go
                stmt.revenue_compost = rev.compost
                stmt.total_revenue = rev.total
            
            # OPEX
            opex = proj.get_opex(year)
            if opex:
                stmt.opex_feedstock = opex.feedstock_handling
                stmt.opex_utilities = opex.utilities
                stmt.opex_chemicals = opex.chemicals
                stmt.opex_maintenance = opex.maintenance
                stmt.opex_personnel = opex.personnel
                stmt.opex_insurance = opex.insurance
                stmt.opex_overheads = opex.overheads
                stmt.opex_digestate = opex.digestate_handling
                stmt.opex_other = opex.other
                stmt.total_opex = opex.total
            
            # Operating
            stmt.ebitda = proj.ebitda.get(year, 0.0)
            stmt.depreciation = proj.depreciation.get(year, 0.0)
            stmt.grant_income_release = proj.grant_income_release.get(year, 0.0)
            stmt.ebit = proj.ebit.get(year, 0.0)
            
            # Financing
            stmt.interest_expense = proj.interest.get(year, 0.0)
            stmt.ebt = proj.ebt.get(year, 0.0)
            
            # Taxes
            stmt.taxes_before_credit = proj.taxes_before_credit.get(year, 0.0)
            stmt.tax_credit_utilization = proj.tax_credit_utilization.get(year, 0.0)
            stmt.taxes_paid = proj.taxes_paid.get(year, 0.0)
            
            stmt.net_income = proj.net_income.get(year, 0.0)
            
            statements.append(stmt)
        
        return statements
    
    def _build_cash_flows(self) -> list[CashFlowLine]:
        """Build cash flow statements from projections."""
        statements = []
        proj = self.projections
        tax_rate = self.case.financing.tax_rate
        
        prev_ar = 0.0
        prev_ap = 0.0
        prev_cash = self.case.financing.cash_at_base
        
        for year in proj.all_forecast_years:
            stmt = CashFlowLine(year=year)
            stmt.opening_cash = prev_cash
            
            # CFO
            stmt.ebit = proj.ebit.get(year, 0.0)
            stmt.tax_on_ebit = stmt.ebit * tax_rate if stmt.ebit > 0 else 0.0
            stmt.nopat = stmt.ebit - stmt.tax_on_ebit
            stmt.depreciation = proj.depreciation.get(year, 0.0)
            
            # Grant income release is non-cash, reverse it out
            stmt.grant_income_release = proj.grant_income_release.get(year, 0.0)
            
            # Changes in working capital
            curr_ar = proj.ar_trade.get(year, 0.0)
            curr_ap = proj.ap_trade.get(year, 0.0)
            stmt.change_in_trade_receivables = -(curr_ar - prev_ar)  # Increase is cash outflow
            stmt.change_in_trade_payables = curr_ap - prev_ap  # Increase is cash inflow
            stmt.change_in_nwc = stmt.change_in_trade_receivables + stmt.change_in_trade_payables
            
            # Tax adjustment (tax credit utilization is non-cash benefit already in stmt)
            stmt.taxes_paid_adjustment = proj.tax_credit_utilization.get(year, 0.0)
            
            stmt.cfo = (
                proj.net_income.get(year, 0.0)  # Start from Net Income
                + proj.interest.get(year, 0.0)  # Add back interest (paid in CFF)
                + stmt.depreciation
                - stmt.grant_income_release  # Reverse non-cash income
                + stmt.change_in_nwc
                # Note: tax_credit_utilization NOT added - already in Net Income
            )
            
            # CFI
            capex = proj.get_capex(year)
            stmt.capex = -(capex.total) if capex else 0.0  # Outflow
            
            # Grant cash received (classified in CFI per policy)
            grant_cash = 0.0
            if proj.accounting:
                gr = proj.accounting.get_grant_receivable(year)
                if gr:
                    grant_cash = gr.cash_received
            stmt.grant_cash_received = grant_cash
            
            stmt.cfi = stmt.capex + stmt.grant_cash_received
            
            # CFF - Financing activities
            fin = proj.get_financing(year)
            if fin:
                stmt.debt_drawdown = fin.debt_drawdown
                stmt.debt_repayment = -fin.debt_repayment  # Outflow
                stmt.interest_paid = -fin.interest_expense  # Outflow
            
            # Equity contribution: funds CAPEX not covered by debt and grants
            # Only applies during construction years
            capex_amount = abs(stmt.capex)  # Make positive for comparison
            debt_this_year = stmt.debt_drawdown
            grant_this_year = stmt.grant_cash_received
            
            if year in proj.construction_years or year == proj.cod_year:
                equity_needed = max(0, capex_amount - debt_this_year - grant_this_year)
                stmt.equity_contribution = equity_needed
            
            stmt.cff = stmt.debt_drawdown + stmt.debt_repayment + stmt.interest_paid + stmt.equity_contribution
            
            # Summary
            stmt.net_cash_flow = stmt.cfo + stmt.cfi + stmt.cff
            stmt.closing_cash = stmt.opening_cash + stmt.net_cash_flow
            
            statements.append(stmt)
            
            prev_ar = curr_ar
            prev_ap = curr_ap
            prev_cash = stmt.closing_cash
        
        return statements
    
    def _build_balance_sheets(self) -> list[BalanceSheetLine]:
        """Build balance sheets from projections and cash flow statements."""
        statements = []
        proj = self.projections
        
        # Build cash flows first (we need them for cash position)
        cash_flows = self._build_cash_flows()
        cf_by_year = {cf.year: cf for cf in cash_flows}
        
        # Track cumulative share capital
        # Start with cash_at_base (actual paid-in capital at start)
        # Equity contributions during construction add to this
        cumulative_equity_contrib = self.case.financing.cash_at_base
        prev_retained = 0.0
        
        for i, year in enumerate(proj.all_forecast_years):
            stmt = BalanceSheetLine(year=year)
            
            # Cash from cash flow statement
            cf = cf_by_year.get(year)
            stmt.cash = cf.closing_cash if cf else 0.0
            
            # Trade receivables
            stmt.trade_receivables = proj.ar_trade.get(year, 0.0)
            
            # Grant receivable
            if proj.accounting:
                gr = proj.accounting.get_grant_receivable(year)
                if gr:
                    stmt.grant_receivable = gr.closing_balance
            
            # Tax credit receivable
            # Only show on BS under B2 policy (dta_receivable)
            # Under B1 (reduce_tax), credit directly reduces tax expense - no BS entry
            from dcf_projects.biometano.schema import TaxCreditPolicy
            if self.case.incentives.tax_credit.accounting_policy != TaxCreditPolicy.REDUCE_TAX_EXPENSE:
                if proj.accounting:
                    tc = proj.accounting.get_tax_credit(year)
                    if tc:
                        stmt.tax_credit_receivable = tc.closing_balance
            
            stmt.total_current_assets = (
                stmt.cash + stmt.trade_receivables + 
                stmt.grant_receivable + stmt.tax_credit_receivable
            )
            
            # Fixed assets
            if proj.accounting:
                fa = proj.accounting.get_fixed_asset(year)
                if fa:
                    stmt.fixed_assets_gross = fa.closing_gross
                    stmt.accumulated_depreciation = fa.closing_accum_depr
                    stmt.fixed_assets_net = fa.net_book_value
            
            stmt.total_assets = stmt.total_current_assets + stmt.fixed_assets_net
            
            # Liabilities - Trade payables
            stmt.trade_payables = proj.ap_trade.get(year, 0.0)
            stmt.taxes_payable = 0.0  # Assumed paid in period
            stmt.total_current_liabilities = stmt.trade_payables + stmt.taxes_payable
            
            # Debt
            fin = proj.get_financing(year)
            stmt.debt = fin.debt_balance_closing if fin else 0.0
            
            # Deferred income (grant policy A2)
            if proj.accounting:
                di = proj.accounting.get_deferred_income(year)
                if di:
                    stmt.deferred_income = di.closing_balance
            
            stmt.total_non_current_liabilities = stmt.debt + stmt.deferred_income
            stmt.total_liabilities = stmt.total_current_liabilities + stmt.total_non_current_liabilities
            
            # Equity calculation
            # Share capital = initial equity + cumulative equity contributions from cash flow
            net_income = proj.net_income.get(year, 0.0)
            stmt.current_year_profit = net_income
            stmt.retained_earnings = prev_retained
            
            # Accumulate equity contributions from cash flows
            cumulative_equity_contrib += cf.equity_contribution if cf else 0.0
            stmt.share_capital = cumulative_equity_contrib
            
            stmt.total_equity = stmt.share_capital + stmt.retained_earnings + stmt.current_year_profit
            
            stmt.total_liabilities_and_equity = stmt.total_liabilities + stmt.total_equity
            stmt.balance_check = stmt.total_assets - stmt.total_liabilities_and_equity
            
            statements.append(stmt)
            
            # Roll forward retained earnings
            prev_retained = stmt.retained_earnings + stmt.current_year_profit
        
        return statements
    
    def _get_cash_flow_for_year(
        self, 
        statements: list[BalanceSheetLine], 
        year: int
    ) -> Optional[CashFlowLine]:
        """Helper to get cash flow for a year."""
        return None  # Not used
    
    def _find_cash_flow(self, year: int, proj: BiometanoProjections) -> float:
        """Compute net cash flow for a year from projections."""
        # Simplified: use FCFF as proxy, adjusted for financing
        fcff = proj.fcff.get(year, 0.0)
        fin = proj.get_financing(year)
        
        net_borrowing = 0.0
        interest = 0.0
        if fin:
            net_borrowing = fin.debt_drawdown - fin.debt_repayment
            interest = fin.interest_expense
        
        # Grant cash
        grant_cash = 0.0
        if proj.accounting:
            gr = proj.accounting.get_grant_receivable(year)
            if gr:
                grant_cash = gr.cash_received
        
        return fcff + net_borrowing - interest + grant_cash


def build_statements(
    case: BiometanoCase, 
    projections: BiometanoProjections
) -> FinancialStatements:
    """
    Convenience function to build statements.
    
    Args:
        case: Biometano case inputs
        projections: Built projections from builder
        
    Returns:
        Complete FinancialStatements
    """
    builder = StatementsBuilder(case, projections)
    return builder.build()
