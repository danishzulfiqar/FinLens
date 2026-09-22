"""Reusable bank-statement analysis library (parsing, analytics, Excel export)."""

from .analysis import (
    CATEGORY_COLORS,
    categorize,
    category_breakdown,
    kpi_metrics,
    monthly_summary,
    reconcile,
    statement_to_dataframe,
)
from .dashboard import build_dashboard_workbook
from .excel import build_excel_report
from .parser import LAYOUTS, Statement, parse_statement

__all__ = [
    "LAYOUTS",
    "Statement",
    "parse_statement",
    "statement_to_dataframe",
    "reconcile",
    "monthly_summary",
    "categorize",
    "category_breakdown",
    "kpi_metrics",
    "CATEGORY_COLORS",
    "build_excel_report",
    "build_dashboard_workbook",
]
