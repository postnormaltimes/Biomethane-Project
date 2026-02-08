"""
DCF I/O Readers

YAML and JSON input file parsing.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import yaml

from dcf_engine.models import (
    DCFInputs,
    TimelineInputs,
    RevenueInputs,
    OperatingInputs,
    NWCInputs,
    InvestmentInputs,
    TaxInputs,
    CAPMInputs,
    DebtInputs,
    WACCInputs,
    EquityBookInputs,
    TerminalValueInputs,
    NetDebtInputs,
    DiscountingMode,
    TerminalValueMethod,
    WeightingMode,
    ExitMultipleMetric,
)


def _convert_int_keys(data: dict) -> dict:
    """Convert string keys to int for year-indexed dicts."""
    if not isinstance(data, dict):
        return data
    
    result = {}
    for key, value in data.items():
        # Try to convert key to int if it looks like a year
        try:
            int_key = int(key)
            result[int_key] = value
        except (ValueError, TypeError):
            result[key] = value
    
    return result


def _parse_timeline(data: dict) -> TimelineInputs:
    """Parse timeline section."""
    return TimelineInputs(
        base_year=data["base_year"],
        forecast_years=data["forecast_years"],
    )


def _parse_revenue(data: dict) -> RevenueInputs:
    """Parse revenue section."""
    return RevenueInputs(
        base_revenue=data.get("base_revenue"),
        explicit_revenue=_convert_int_keys(data.get("explicit_revenue")) if data.get("explicit_revenue") else None,
        growth_rates=_convert_int_keys(data.get("growth_rates")) if data.get("growth_rates") else None,
    )


def _parse_operating(data: dict) -> OperatingInputs:
    """Parse operating section."""
    return OperatingInputs(
        cost_ratios=_convert_int_keys(data.get("cost_ratios")) if data.get("cost_ratios") else None,
        explicit_ebitda=_convert_int_keys(data.get("explicit_ebitda")) if data.get("explicit_ebitda") else None,
        depreciation_amortization=_convert_int_keys(data["depreciation_amortization"]),
    )


def _parse_nwc(data: dict) -> NWCInputs:
    """Parse NWC section."""
    return NWCInputs(
        base_nwc=data.get("base_nwc"),
        nwc_percent=_convert_int_keys(data.get("nwc_percent")) if data.get("nwc_percent") else None,
        explicit_nwc=_convert_int_keys(data.get("explicit_nwc")) if data.get("explicit_nwc") else None,
    )


def _parse_investments(data: dict) -> InvestmentInputs:
    """Parse investments section."""
    return InvestmentInputs(
        capex=_convert_int_keys(data["capex"]),
    )


def _parse_tax(data: dict) -> TaxInputs:
    """Parse tax section."""
    tax_rate = data["tax_rate"]
    if isinstance(tax_rate, dict):
        tax_rate = _convert_int_keys(tax_rate)
    return TaxInputs(tax_rate=tax_rate)


def _parse_capm(data: dict) -> CAPMInputs:
    """Parse CAPM section."""
    return CAPMInputs(
        rf=data["rf"],
        rm=data["rm"],
        beta=data["beta"],
        ke_override=data.get("ke_override"),
    )


def _parse_debt(data: dict) -> DebtInputs:
    """Parse debt section."""
    rd = data["rd"]
    if isinstance(rd, dict):
        rd = _convert_int_keys(rd)
    return DebtInputs(
        debt_balances=_convert_int_keys(data["debt_balances"]),
        rd=rd,
    )


def _parse_wacc(data: dict) -> WACCInputs:
    """Parse WACC section."""
    weighting_mode = WeightingMode(data.get("weighting_mode", "target"))
    
    equity_book_inputs = None
    if "equity_book_inputs" in data:
        eb_data = data["equity_book_inputs"]
        equity_book_inputs = EquityBookInputs(
            base_equity_book=eb_data["base_equity_book"],
            dividends=_convert_int_keys(eb_data.get("dividends")) if eb_data.get("dividends") else None,
            new_equity=_convert_int_keys(eb_data.get("new_equity")) if eb_data.get("new_equity") else None,
        )
    
    return WACCInputs(
        weighting_mode=weighting_mode,
        wE=data.get("wE"),
        wD=data.get("wD"),
        equity_book_inputs=equity_book_inputs,
    )


def _parse_terminal_value(data: dict) -> TerminalValueInputs:
    """Parse terminal value section."""
    method = TerminalValueMethod(data["method"])
    
    exit_metric = None
    if data.get("exit_metric"):
        exit_metric = ExitMultipleMetric(data["exit_metric"])
    
    return TerminalValueInputs(
        method=method,
        g=data.get("g"),
        exit_multiple=data.get("exit_multiple"),
        exit_metric=exit_metric,
    )


def _parse_net_debt(data: dict) -> NetDebtInputs:
    """Parse net debt section."""
    return NetDebtInputs(
        cash_and_equivalents=data["cash_and_equivalents"],
    )


def parse_input_dict(data: dict[str, Any]) -> DCFInputs:
    """
    Parse a dictionary of inputs into DCFInputs model.
    
    This is the core parsing function used by both YAML and JSON readers.
    
    Args:
        data: Raw input dictionary
    
    Returns:
        DCFInputs model
    """
    discounting_mode = DiscountingMode(
        data.get("discounting_mode", "year_specific_flat")
    )
    
    return DCFInputs(
        timeline=_parse_timeline(data["timeline"]),
        revenue=_parse_revenue(data["revenue"]),
        operating=_parse_operating(data["operating"]),
        nwc=_parse_nwc(data["nwc"]),
        investments=_parse_investments(data["investments"]),
        tax=_parse_tax(data["tax"]),
        capm=_parse_capm(data["capm"]),
        debt=_parse_debt(data["debt"]),
        wacc=_parse_wacc(data["wacc"]),
        terminal_value=_parse_terminal_value(data["terminal_value"]),
        net_debt=_parse_net_debt(data["net_debt"]),
        discounting_mode=discounting_mode,
    )


def read_yaml(path: str | Path) -> DCFInputs:
    """
    Read DCF inputs from a YAML file.
    
    Args:
        path: Path to YAML file
    
    Returns:
        DCFInputs model
    """
    path = Path(path)
    with open(path, "r") as f:
        data = yaml.safe_load(f)
    
    return parse_input_dict(data)


def read_json(path: str | Path) -> DCFInputs:
    """
    Read DCF inputs from a JSON file.
    
    Args:
        path: Path to JSON file
    
    Returns:
        DCFInputs model
    """
    path = Path(path)
    with open(path, "r") as f:
        data = json.load(f)
    
    return parse_input_dict(data)


def read_input_file(path: str | Path) -> DCFInputs:
    """
    Read DCF inputs from a file (auto-detects format).
    
    Args:
        path: Path to input file (YAML or JSON)
    
    Returns:
        DCFInputs model
    """
    path = Path(path)
    suffix = path.suffix.lower()
    
    if suffix in (".yaml", ".yml"):
        return read_yaml(path)
    elif suffix == ".json":
        return read_json(path)
    else:
        raise ValueError(f"Unsupported file format: {suffix}")
