"""Analytics: typed DataFrame, reconciliation, monthly rollups, categorization, KPIs."""

from __future__ import annotations

import pandas as pd

from .parser import Statement

TX_COLUMNS = [
    "date",
    "month",
    "description",
    "type",
    "debit",
    "credit",
    "amount",
    "balance",
    "source_file",
]

# Keyword rules used to auto-classify transactions. First match wins; order matters.
# Expense rules apply to debits, income rules to credits. Everything else falls back
# to "Other Expense" / "Other Income" and can be overridden by the user in the app.
EXPENSE_RULES: list[tuple[str, list[str]]] = [
    ("Bills & Utilities", ["electric", "k-electric", "kelectric", "wapda", "sui", "gas bill",
                            "ptcl", "wifi", "internet", "lesco", "iesco", "gepco", "utility", "bill"]),
    ("Telecom & Mobile", ["jazz", "zong", "ufone", "telenor", "easypaisa", "jazzcash", "mobile",
                           "recharge", "top up", "topup", "airtime", "scratch", "load"]),
    ("Food & Dining", ["food", "restaurant", "cafe", "foodpanda", "kfc", "mcdonald", "pizza",
                        "bakery", "dine", "eat", "burger", "coffee"]),
    ("Groceries & Shopping", ["mart", "store", "shop", "daraz", "amazon", "grocery", "bazar",
                              "supermarket", "metro", "imtiaz", "carrefour", "pos"]),
    ("Fuel & Transport", ["fuel", "petrol", "pso", "shell", "parco", "byco", "careem", "uber",
                          "indrive", "bykea", "toll", "transport"]),
    ("Cash & ATM", ["atm", "cash withdraw", "withdrawal", "cash wdl", "cwd"]),
    ("Transfers Out", ["ibft", "funds transfer", "fund transfer", "1link", "raast", "ift",
                       "transfer to", "sent to", "remit", "outward"]),
    ("Fees & Charges", ["charge", "fee", "wht", "service charge", "duty", "stamp", "tax",
                        "fed ", "excise", "penalty"]),
    ("Health", ["pharmacy", "hospital", "clinic", "medical", "lab", "dawa"]),
    ("Education", ["school", "university", "college", "tuition", "academy", "course"]),
    ("Subscriptions", ["netflix", "spotify", "youtube", "subscription", "google", "apple.com",
                       "microsoft", "openai", "prime"]),
]
INCOME_RULES: list[tuple[str, list[str]]] = [
    ("Salary", ["salary", "payroll", "wages", "pay"]),
    ("Profit & Returns", ["profit", "return", "markup", "interest", "dividend", "reward"]),
    ("Refunds", ["refund", "reversal", "cashback", "chargeback"]),
    ("Transfers In", ["ibft", "funds transfer", "transfer from", "received", "raast", "inward",
                      "credit from", "deposit"]),
]

# Stable colors so categories look the same across every chart.
CATEGORY_COLORS: dict[str, str] = {
    "Salary": "#2E7D32",
    "Profit & Returns": "#43A047",
    "Refunds": "#66BB6A",
    "Transfers In": "#81C784",
    "Other Income": "#A5D6A7",
    "Bills & Utilities": "#1F4E79",
    "Telecom & Mobile": "#2E75B6",
    "Food & Dining": "#E64A19",
    "Groceries & Shopping": "#F57C00",
    "Fuel & Transport": "#6D4C41",
    "Cash & ATM": "#8E24AA",
    "Transfers Out": "#5E35B1",
    "Fees & Charges": "#C62828",
    "Health": "#00838F",
    "Education": "#3949AB",
    "Subscriptions": "#AD1457",
    "Other Expense": "#90A4AE",
}


def _to_number(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.replace(",", "", regex=False)
        .replace("", "0")
        .astype(float)
    )


def statement_to_dataframe(stmt: Statement) -> pd.DataFrame:
    """Turn a parsed Statement into a typed, analysis-ready DataFrame."""
    records = stmt.transactions
    if stmt.metadata.get("order") == "desc":
        records = list(reversed(records))

    df = pd.DataFrame(records)
    if df.empty:
        return df

    df["date"] = pd.to_datetime(df["date"], format=stmt.metadata.get("date_fmt"))
    for col in ("debit", "credit", "balance"):
        df[col] = _to_number(df[col])

    df["amount"] = df["credit"] - df["debit"]
    df["type"] = df["amount"].apply(lambda v: "Credit" if v >= 0 else "Debit")
    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["source_file"] = stmt.metadata.get("source_file")
    return df.reset_index(drop=True)[TX_COLUMNS]


def categorize(df: pd.DataFrame) -> pd.DataFrame:
    """Add a rule-based ``category`` column (income vs expense aware)."""
    if df.empty:
        df = df.copy()
        df["category"] = pd.Series(dtype="object")
        return df

    desc = df["description"].fillna("").str.lower()
    is_credit = df["amount"] >= 0
    out = []
    for credit, text in zip(is_credit, desc):
        rules = INCOME_RULES if credit else EXPENSE_RULES
        fallback = "Other Income" if credit else "Other Expense"
        match = fallback
        for name, keywords in rules:
            if any(k in text for k in keywords):
                match = name
                break
        out.append(match)
    df = df.copy()
    df["category"] = out
    return df


def reconcile(stmt: Statement, df: pd.DataFrame, tol: float = 0.01) -> dict:
    """Replay from the opening balance and compare against printed balances."""
    opening = float((stmt.metadata.get("opening_balance") or "0").replace(",", ""))
    closing = float((stmt.metadata.get("closing_balance") or "0").replace(",", ""))
    running = opening + df["amount"].cumsum()
    net = round(float(df["amount"].sum()), 2)
    return {
        "source_file": stmt.metadata.get("source_file"),
        "layout": stmt.metadata.get("layout"),
        "transactions": len(df),
        "opening_balance": opening,
        "closing_balance": closing,
        "computed_closing": round(float(running.iloc[-1]) if len(running) else opening, 2),
        "row_mismatches": int((running.sub(df["balance"]).abs() > tol).sum()),
        "closing_ok": abs(opening + net - closing) <= tol,
    }


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Per-month income, expense, net and transaction count."""
    if df.empty:
        return pd.DataFrame(
            columns=["month", "transactions", "income", "expense", "net", "closing_balance"]
        )
    summary = (
        df.groupby("month")
        .agg(
            transactions=("amount", "size"),
            income=("credit", "sum"),
            expense=("debit", "sum"),
            net=("amount", "sum"),
        )
        .reset_index()
    )
    summary["closing_balance"] = df.groupby("month")["balance"].last().values
    return summary.round(2)


def category_breakdown(df: pd.DataFrame, kind: str = "expense") -> pd.DataFrame:
    """Total value and share per category. ``kind`` is 'expense' or 'income'."""
    if df.empty or "category" not in df:
        return pd.DataFrame(columns=["category", "total", "transactions", "share"])
    value_col = "debit" if kind == "expense" else "credit"
    subset = df[df[value_col] > 0]
    if subset.empty:
        return pd.DataFrame(columns=["category", "total", "transactions", "share"])
    grouped = (
        subset.groupby("category")
        .agg(total=(value_col, "sum"), transactions=(value_col, "size"))
        .reset_index()
        .sort_values("total", ascending=False)
    )
    grouped["share"] = (grouped["total"] / grouped["total"].sum() * 100).round(1)
    return grouped.round(2)


def kpi_metrics(df: pd.DataFrame) -> dict:
    """Headline numbers for the current (already-filtered) transaction set."""
    if df.empty:
        return {
            "income": 0.0, "expense": 0.0, "net": 0.0, "count": 0,
            "avg_expense": 0.0, "savings_rate": 0.0,
            "closing_balance": 0.0, "months": 0,
        }
    income = float(df["credit"].sum())
    expense = float(df["debit"].sum())
    debits = df[df["debit"] > 0]
    return {
        "income": income,
        "expense": expense,
        "net": income - expense,
        "count": int(len(df)),
        "avg_expense": float(debits["debit"].mean()) if not debits.empty else 0.0,
        "savings_rate": ((income - expense) / income * 100) if income else 0.0,
        "closing_balance": float(df.sort_values("date")["balance"].iloc[-1]),
        "months": int(df["month"].nunique()),
    }
