"""Interactive Excel dashboard builder (the same workbook the notebook produces).

Takes finlib-shaped, categorized transaction data and writes a filterable,
chart-rich ``.xlsx`` dashboard driven entirely by native Excel formulas.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- palette -----------------------------------------------------------------
NAVY = "1F3864"
BLUE = "2E75B6"
GREEN = "548235"
RED = "C00000"
LIGHT = "D9E1F2"
WHITE = "FFFFFF"

_thin = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_MONEY_FMT = "#,##0.00"
_INT_FMT = "#,##0"
_DATE_FMT = "yyyy-mm-dd"
_TABLE_STYLE = TableStyleInfo(
    name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
)

# Column order the dashboard formulas rely on (month=B, account=C, debit=F, credit=G
# after "account" is inserted at position 2).
_CANONICAL = [
    "date", "month", "description", "type",
    "debit", "credit", "amount", "balance", "source_file", "category",
]


def _acct_labels(transactions: pd.DataFrame, report: pd.DataFrame) -> pd.Series:
    """Friendly account label per row, e.g. 'MYABL (1789...pdf)'."""
    layout_by_file = (
        dict(zip(report["source_file"], report["layout"])) if not report.empty else {}
    )

    def label(sf: str) -> str:
        lay = layout_by_file.get(sf, "")
        return f"{lay.upper()} ({sf})" if lay else str(sf)

    return transactions["source_file"].map(label)


def _style_table_sheet(ws, df, table_name, *, tab_color=NAVY):
    """Give a data sheet the shared look: table, banding, formats, freeze."""
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab_color
    nrows, ncols = df.shape
    ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = _TABLE_STYLE
    ws.add_table(table)
    ws.freeze_panes = "A2"
    for idx, col in enumerate(df.columns, start=1):
        letter = get_column_letter(idx)
        if col == "date":
            fmt = _DATE_FMT
        elif pd.api.types.is_float_dtype(df[col].dtype):
            fmt = _MONEY_FMT
        elif pd.api.types.is_integer_dtype(df[col].dtype):
            fmt = _INT_FMT
        else:
            fmt = None
        if fmt:
            for cell in ws[letter][1:]:
                cell.number_format = fmt
        width = max([len(str(col))] + df[col].astype(str).str.len().tolist())
        ws.column_dimensions[letter].width = min(max(width + 2, 10), 55)


def _kpi_card(ws, cell_label, cell_value, title, formula, *, fmt=_MONEY_FMT, color=NAVY):
    """Draw a titled KPI card (label cell above value cell)."""
    lab = ws[cell_label]
    lab.value = title
    lab.font = Font(bold=True, color=WHITE, size=10)
    lab.fill = PatternFill("solid", fgColor=color)
    lab.alignment = Alignment(horizontal="center", vertical="center")
    lab.border = _BORDER
    val = ws[cell_value]
    val.value = formula
    val.font = Font(bold=True, size=14, color=color)
    val.alignment = Alignment(horizontal="center", vertical="center")
    val.number_format = fmt
    val.fill = PatternFill("solid", fgColor=LIGHT)
    val.border = _BORDER


def build_dashboard_workbook(transactions, monthly, report, path):
    """Write a filterable, chart-rich Excel dashboard from the extracted data.

    ``transactions`` must be a categorized DataFrame (as produced by
    ``finlib.categorize``); ``monthly`` from ``finlib.monthly_summary``; and
    ``report`` a DataFrame of ``finlib.reconcile`` dicts (one row per statement).
    """
    path = Path(path)
    if transactions.empty:
        raise ValueError("No transactions to export")

    # Enforce the column order the formulas depend on.
    cols = [c for c in _CANONICAL if c in transactions.columns]
    tx = transactions[cols].copy()
    tx.insert(2, "account", _acct_labels(tx, report))
    n = len(tx)
    last = n + 1  # last data row (row 1 is the header)

    months = sorted(tx["month"].unique().tolist())
    years = sorted({m[:4] for m in months})
    periods = ["All"] + years + months
    accounts = ["All"] + sorted(tx["account"].unique().tolist())

    with pd.ExcelWriter(path, engine="openpyxl", datetime_format=_DATE_FMT) as writer:
        book = writer.book
        # Force Excel to recalculate every formula on open so the KPIs, charts and
        # the filtered transaction list populate immediately (openpyxl caches none).
        book.calculation.fullCalcOnLoad = True

        # --- data sheets (shared styling) ------------------------------------
        tx_out = tx.copy()
        tx_out["date"] = pd.to_datetime(tx_out["date"]).dt.date
        tx_out.to_excel(writer, sheet_name="Transactions", index=False)
        monthly.to_excel(writer, sheet_name="Monthly Summary", index=False)
        report.to_excel(writer, sheet_name="Reconciliation", index=False)

        _style_table_sheet(writer.sheets["Transactions"], tx_out, "tblTxns")
        _style_table_sheet(writer.sheets["Monthly Summary"], monthly, "tblMonthly", tab_color=BLUE)
        _style_table_sheet(writer.sheets["Reconciliation"], report, "tblRecon", tab_color=GREEN)

        # --- hidden list + calc sheets ---------------------------------------
        ws_lists = book.create_sheet("Lists")
        for i, p in enumerate(periods, start=1):
            ws_lists.cell(row=i, column=1, value=p)
        for i, a in enumerate(accounts, start=1):
            ws_lists.cell(row=i, column=3, value=a)
        ws_lists.sheet_state = "hidden"

        ws_calc = book.create_sheet("Calc")
        # Selection -> SUMIFS criteria (wildcard handles All / year / month).
        ws_calc["B1"] = '=IF(Dashboard!$C$4="All","*",Dashboard!$C$4&"*")'
        ws_calc["B2"] = '=IF(Dashboard!$C$5="All","*",Dashboard!$C$5)'
        # Pie source (income vs expense for the current selection).
        ws_calc["D1"], ws_calc["E1"] = "Category", "Amount"
        ws_calc["D2"], ws_calc["E2"] = "Income", (
            "=SUMIFS(Transactions!$G:$G,Transactions!$B:$B,$B$1,Transactions!$C:$C,$B$2)"
        )
        ws_calc["D3"], ws_calc["E3"] = "Expense", (
            "=SUMIFS(Transactions!$F:$F,Transactions!$B:$B,$B$1,Transactions!$C:$C,$B$2)"
        )
        # Monthly trend block (account-filtered, every month).
        ws_calc["A6"], ws_calc["B6"], ws_calc["C6"], ws_calc["D6"] = (
            "Month", "Income", "Expense", "Net",
        )
        for j, month in enumerate(months):
            r = 7 + j
            ws_calc.cell(row=r, column=1, value=month)
            ws_calc.cell(
                row=r, column=2,
                value=f"=SUMIFS(Transactions!$G:$G,Transactions!$B:$B,$A{r},Transactions!$C:$C,$B$2)",
            )
            ws_calc.cell(
                row=r, column=3,
                value=f"=SUMIFS(Transactions!$F:$F,Transactions!$B:$B,$A{r},Transactions!$C:$C,$B$2)",
            )
            ws_calc.cell(row=r, column=4, value=f"=B{r}-C{r}")
        last_month_row = 6 + len(months)
        ws_calc.sheet_state = "hidden"

        # --- dashboard --------------------------------------------------------
        ws = book.create_sheet("Dashboard")
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = NAVY
        ws.column_dimensions["A"].width = 3
        for letter in "BCDEFGHIJKLMN":
            ws.column_dimensions[letter].width = 15

        ws.merge_cells("B2:H2")
        title = ws["B2"]
        title.value = "FinLens Dashboard"
        title.font = Font(bold=True, size=20, color=NAVY)
        title.alignment = Alignment(vertical="center")

        # Filters
        for cell, text in (("B4", "Period:"), ("B5", "Account:")):
            ws[cell] = text
            ws[cell].font = Font(bold=True)
        for cell, default in (("C4", "All"), ("C5", "All")):
            c = ws[cell]
            c.value = default
            c.fill = PatternFill("solid", fgColor="FFF2CC")
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")
            c.border = _BORDER

        dv_period = DataValidation(
            type="list", formula1=f"=Lists!$A$1:$A${len(periods)}", allow_blank=False
        )
        dv_account = DataValidation(
            type="list", formula1=f"=Lists!$C$1:$C${len(accounts)}", allow_blank=False
        )
        ws.add_data_validation(dv_period)
        ws.add_data_validation(dv_account)
        dv_period.add(ws["C4"])
        dv_account.add(ws["C5"])

        # KPI cards
        income = "=SUMIFS(Transactions!$G:$G,Transactions!$B:$B,Calc!$B$1,Transactions!$C:$C,Calc!$B$2)"
        expense = "=SUMIFS(Transactions!$F:$F,Transactions!$B:$B,Calc!$B$1,Transactions!$C:$C,Calc!$B$2)"
        count = "=COUNTIFS(Transactions!$B:$B,Calc!$B$1,Transactions!$C:$C,Calc!$B$2)"
        _kpi_card(ws, "B7", "B8", "Total Income", income, color=GREEN)
        _kpi_card(ws, "D7", "D8", "Total Expense", expense, color=RED)
        _kpi_card(ws, "F7", "F8", "Net", "=B8-D8", color=NAVY)
        _kpi_card(ws, "H7", "H8", "Transactions", count, fmt=_INT_FMT, color=BLUE)

        # Pie: income vs expense (selection)
        pie = PieChart()
        pie.title = "Income vs Expense (selection)"
        pie.add_data(Reference(ws_calc, min_col=5, min_row=2, max_row=3), titles_from_data=False)
        pie.set_categories(Reference(ws_calc, min_col=4, min_row=2, max_row=3))
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.height, pie.width = 7.5, 11
        ws.add_chart(pie, "B11")

        # Bar: monthly net (account-filtered)
        bar = BarChart()
        bar.type = "col"
        bar.title = "Monthly Net (selected account)"
        bar.add_data(
            Reference(ws_calc, min_col=4, min_row=6, max_row=last_month_row), titles_from_data=True
        )
        bar.set_categories(Reference(ws_calc, min_col=1, min_row=7, max_row=last_month_row))
        bar.legend = None
        bar.height, bar.width = 7.5, 16
        ws.add_chart(bar, "F11")

        # Line: monthly income vs expense
        line = LineChart()
        line.title = "Monthly Income vs Expense"
        line.add_data(
            Reference(ws_calc, min_col=2, max_col=3, min_row=6, max_row=last_month_row),
            titles_from_data=True,
        )
        line.set_categories(Reference(ws_calc, min_col=1, min_row=7, max_row=last_month_row))
        line.height, line.width = 8, 27
        ws.add_chart(line, "B27")

        # Filtered transaction list — universal INDEX/MATCH (works in every Excel).
        # Helper columns on Transactions: L flags rows matching the dashboard
        # selection, M numbers each match 1..k so INDEX/MATCH can pull the k-th row.
        ws_tx = writer.sheets["Transactions"]
        ws_tx["L1"], ws_tx["M1"] = "_match", "_idx"
        for r in range(2, last + 1):
            ws_tx.cell(
                row=r, column=12,
                value=(
                    '=IF(AND('
                    f'OR(Dashboard!$C$4="All",LEFT($B{r},LEN(Dashboard!$C$4))=Dashboard!$C$4),'
                    f'OR(Dashboard!$C$5="All",$C{r}=Dashboard!$C$5)),1,0)'
                ),
            )
            ws_tx.cell(row=r, column=13, value=f'=IF(L{r}=1,SUM($L$2:L{r}),"")')
        ws_tx.column_dimensions["L"].hidden = True
        ws_tx.column_dimensions["M"].hidden = True

        ws["B44"] = "Transactions for current selection"
        ws["B44"].font = Font(bold=True, size=12, color=NAVY)
        ws["B44"].alignment = Alignment(vertical="center")
        headers = list(tx_out.columns)
        money_idx = {headers.index(c) for c in ("debit", "credit", "amount", "balance")}
        for i, h in enumerate(headers):
            c = ws.cell(row=45, column=2 + i, value=h)
            c.font = Font(bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=BLUE)
            c.alignment = Alignment(horizontal="center")
            c.border = _BORDER
        detail_start = 46
        for k in range(n):
            r = detail_start + k
            for i in range(len(headers)):
                src = get_column_letter(i + 1)  # Transactions column A..J
                cell = ws.cell(
                    row=r, column=2 + i,
                    value=(
                        f'=IFERROR(INDEX(Transactions!${src}$2:${src}${last},'
                        f'MATCH({k + 1},Transactions!$M$2:$M${last},0)),"")'
                    ),
                )
                if i == 0:
                    cell.number_format = _DATE_FMT
                elif i in money_idx:
                    cell.number_format = _MONEY_FMT

        # Sheet order + active
        order = ["Dashboard", "Transactions", "Monthly Summary", "Reconciliation", "Calc", "Lists"]
        book._sheets.sort(key=lambda s: order.index(s.title))
        book.active = 0

    return path
