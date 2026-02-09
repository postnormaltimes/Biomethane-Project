"""
XLSX Formula Export Layout

Centralized layout specification for formula-driven Excel exports.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from openpyxl.utils import get_column_letter


@dataclass(frozen=True)
class SheetLayout:
    """Layout mapping for year-based tables."""
    start_col: int = 2
    header_row: int = 1

    def year_to_col(self, years: Iterable[int]) -> dict[int, int]:
        return {year: self.start_col + idx for idx, year in enumerate(years)}

    @staticmethod
    def cell(col: int, row: int) -> str:
        return f"{get_column_letter(col)}{row}"
