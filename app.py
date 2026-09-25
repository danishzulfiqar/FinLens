"""FinLens — upload bank statements, explore, categorize, and export.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import io

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from finlib import (
    CATEGORY_COLORS,
    build_excel_report,
    categorize,
    category_breakdown,
    forecast_balance,
    kpi_metrics,
    monthly_summary,
    parse_statement,
    reconcile,
    statement_to_dataframe,
)

NAVY, GREEN, RED, BLUE = "#1F3864", "#2E7D32", "#C62828", "#2E75B6"
CATEGORY_OPTIONS = list(CATEGORY_COLORS.keys())

st.set_page_config(page_title="FinLens", layout="wide")


def _palette(dark: bool) -> dict:
    """Return the colour tokens for the requested light/dark theme."""
    if dark:
        return {
            "bg": "#0B0F17", "panel": "#111726", "text": "#E8ECF4", "muted": "#93A0B7",
            "border": "#232B3D", "card": "#141B2B", "card_text": "#F3F6FC",
            "heading": "#F1F5FF", "accent": "#5E8BFF", "accent_soft": "rgba(94,139,255,.16)",
            "shadow": "0 1px 2px rgba(0,0,0,.5), 0 4px 12px rgba(0,0,0,.32)",
            "shadow_hover": "0 10px 26px rgba(0,0,0,.5)", "hover": "#1A2234",
        }
    return {
        "bg": "#FFFFFF", "panel": "#F6F8FC", "text": "#1B2333", "muted": "#667085",
        "border": "#E6EAF2", "card": "#FFFFFF", "card_text": "#101828",
        "heading": "#1F3864", "accent": "#2F5FD0", "accent_soft": "rgba(47,95,208,.10)",
        "shadow": "0 1px 3px rgba(16,24,40,.08), 0 1px 2px rgba(16,24,40,.04)",
        "shadow_hover": "0 6px 18px rgba(16,24,40,.10)", "hover": "#F0F4FB",
    }


def _inject_css(p: dict) -> None:
    st.markdown(
        f"""
        <style>
          :root {{
            --bg:{p['bg']}; --panel:{p['panel']}; --text:{p['text']}; --muted:{p['muted']};
            --border:{p['border']}; --card:{p['card']}; --card-text:{p['card_text']};
            --heading:{p['heading']}; --accent:{p['accent']}; --accent-soft:{p['accent_soft']};
            --shadow:{p['shadow']}; --shadow-hover:{p['shadow_hover']}; --hover:{p['hover']};
          }}
          [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] .main {{
            background: var(--bg); color: var(--text);
          }}
          [data-testid="stHeader"] {{ background: transparent; }}
          section[data-testid="stSidebar"] {{
            background: var(--panel); border-right: 1px solid var(--border);
          }}
          section[data-testid="stSidebar"] * {{ color: var(--text); }}
          .block-container {{ padding-top: 1.6rem; padding-bottom: 2.5rem; max-width: 1360px; }}
          h1, h2, h3, h4, h5, h6, p, label, li {{ color: var(--text); }}

          .app-title {{ font-size: 1.7rem; font-weight: 700; color: var(--heading);
                       letter-spacing: -.01em; margin-bottom: .15rem; }}
          .app-sub {{ color: var(--muted); font-size: .95rem; margin-bottom: .6rem; }}
          .app-divider {{ height: 1px; background: var(--border); margin: .2rem 0 1.4rem; }}

          .kpi-card {{ background: var(--card); border: 1px solid var(--border);
                      border-radius: 14px; padding: 16px 18px; box-shadow: var(--shadow);
                      position: relative; overflow: hidden;
                      transition: transform .12s ease, box-shadow .12s ease; }}
          .kpi-card::before {{ content: ""; position: absolute; left: 0; top: 0; bottom: 0;
                      width: 3px; background: var(--accent); }}
          .kpi-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-hover); }}
          .kpi-label {{ font-size: .72rem; color: var(--muted); font-weight: 600;
                       letter-spacing: .06em; text-transform: uppercase; }}
          .kpi-value {{ font-size: 1.5rem; font-weight: 700; color: var(--card-text);
                       margin-top: 4px; letter-spacing: -.02em; }}
          .kpi-sub {{ font-size: .78rem; color: var(--muted); margin-top: 3px; }}

          .stTabs [data-baseweb="tab-list"] {{ gap: 2px; border-bottom: 1px solid var(--border); }}
          .stTabs [data-baseweb="tab"] {{ padding: 9px 18px; font-weight: 600;
                      color: var(--muted); }}
          .stTabs [aria-selected="true"] {{ color: var(--accent) !important; }}
          .stTabs [data-baseweb="tab-highlight"] {{ background: var(--accent); }}

          /* Inputs, selects and uploader share the card surface for a cohesive look. */
          [data-baseweb="input"], [data-baseweb="select"] > div,
          [data-baseweb="textarea"], .stTextInput input, .stNumberInput input {{
            background: var(--card) !important; border-color: var(--border) !important;
          }}
          input, textarea {{ color: var(--text) !important; }}
          [data-testid="stFileUploaderDropzone"] {{
            background: var(--panel); border: 1px dashed var(--border); border-radius: 10px;
          }}
          .stButton > button {{ border-radius: 9px; border: 1px solid var(--border);
                      background: var(--card); color: var(--text); font-weight: 600;
                      transition: background .12s ease, border-color .12s ease; }}
          .stButton > button:hover {{ background: var(--hover); border-color: var(--accent); }}
          .stDownloadButton > button {{ border-radius: 9px; background: var(--accent);
                      border: 1px solid var(--accent); color: #fff; font-weight: 600; }}

          /* Multiselect chips (categories, accounts) — keep them on-theme. */
          span[data-baseweb="tag"] {{ background: var(--accent-soft) !important;
                      border: 1px solid var(--accent) !important; border-radius: 7px !important; }}
          span[data-baseweb="tag"], span[data-baseweb="tag"] span,
          span[data-baseweb="tag"] div {{ color: var(--accent) !important; }}
          span[data-baseweb="tag"] svg {{ fill: var(--accent) !important; color: var(--accent) !important; }}

          /* Dropdown / selectbox popovers follow the active theme. */
          [data-baseweb="popover"] [role="listbox"],
          [data-baseweb="popover"] ul {{ background: var(--card) !important;
                      border: 1px solid var(--border) !important; }}
          [role="option"] {{ color: var(--text) !important; }}
          [role="option"]:hover, li[aria-selected="true"] {{ background: var(--hover) !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


with st.sidebar:
    dark_mode = st.toggle("Dark mode", value=False, key="dark_mode")

THEME = _palette(dark_mode)
PLOTLY_TEMPLATE = "plotly_dark" if dark_mode else "plotly_white"
ACCENT = THEME["accent"]
_inject_css(THEME)


def style_fig(fig, height: int, *, top: int = 50, **layout):
    """Apply the active theme (template, transparent bg, font colour) to a figure."""
    fig.update_layout(
        template=PLOTLY_TEMPLATE, height=height,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font_color=THEME["text"], margin=dict(t=top, l=10, r=10, b=10), **layout,
    )
    return fig


def money(value: float, currency: str = "PKR", decimals: int = 0) -> str:
    return f"{currency} {value:,.{decimals}f}"


def kpi_card(col, label: str, value: str, accent: str, sub: str | None = None) -> None:
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    col.markdown(
        f'<div class="kpi-card" style="--accent:{accent}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f"{sub_html}"
        f"</div>",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def _parse_one(data: bytes, name: str, password: str | None):
    """Parse a single uploaded PDF -> (dataframe, reconcile dict, metadata)."""
    stmt = parse_statement(io.BytesIO(data), password=password or None, source_name=name)
    df = statement_to_dataframe(stmt)
    rec = reconcile(stmt, df) if not df.empty else None
    return df, rec, stmt.metadata


def load_uploads(files, password: str):
    """Parse all uploads into one categorized DataFrame + a reconciliation report."""
    frames, recs, errors = [], [], []
    for f in files:
        data = f.getvalue()
        try:
            df, rec, meta = _parse_one(data, f.name, password)
        except Exception as exc:  # noqa: BLE001 - surface to the user, keep going
            errors.append((f.name, f"{type(exc).__name__}: {exc}"))
            continue
        if df.empty:
            errors.append((f.name, "0 transactions extracted (unrecognised layout?)"))
            continue
        df = df.copy()
        df["account"] = f"{meta['layout'].upper()} ({f.name})"
        frames.append(df)
        if rec:
            recs.append(rec)

    if not frames:
        return pd.DataFrame(), pd.DataFrame(), errors

    tx = pd.concat(frames, ignore_index=True)
    tx = categorize(tx)
    report = pd.DataFrame(recs)
    return tx, report, errors


# --- Header ------------------------------------------------------------------
st.markdown('<div class="app-title">FinLens</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-sub">Bank-statement analysis for myABL and HBL — automatic '
    "categorization, filtering, charts and Excel export.</div>",
    unsafe_allow_html=True,
)
st.markdown('<div class="app-divider"></div>', unsafe_allow_html=True)

# --- Sidebar: upload ---------------------------------------------------------
with st.sidebar:
    st.header("Statements")
    uploads = st.file_uploader(
        "Upload PDF statement(s)", type="pdf", accept_multiple_files=True
    )
    password = st.text_input(
        "Password (for protected PDFs)", type="password",
        help="Applied to every file; ignored by non-protected PDFs.",
    )

if not uploads:
    st.info("Upload one or more bank-statement PDFs from the sidebar to begin.")
    st.markdown(
        "**What you get:** headline KPIs • income vs expense trends • automatic "
        "spending categories you can correct • date/account/category filtering • "
        "reconciliation check • one-click Excel report."
    )
    st.stop()

# --- Parse (re-parse only when the upload set or password changes) ------------
signature = tuple((f.name, f.size) for f in uploads) + (password,)
if st.session_state.get("signature") != signature:
    with st.spinner("Reading statements…"):
        tx, report, errors = load_uploads(uploads, password)
    st.session_state.signature = signature
    st.session_state.tx = tx
    st.session_state.report = report
    st.session_state.errors = errors

tx: pd.DataFrame = st.session_state.tx
report: pd.DataFrame = st.session_state.report
errors: list = st.session_state.errors

for name, msg in errors:
    st.warning(f"**{name}** — {msg}")

if tx.empty:
    st.error("No transactions could be extracted from the uploaded file(s).")
    st.stop()

currency = tx.attrs.get("currency", "PKR")

# --- Sidebar: filters --------------------------------------------------------
with st.sidebar:
    st.header("Filters")
    min_d, max_d = tx["date"].min().date(), tx["date"].max().date()
    if min_d == max_d:
        date_range = (min_d, max_d)
        st.caption(f"Single day: {min_d}")
    else:
        date_range = st.slider(
            "Date range", min_value=min_d, max_value=max_d, value=(min_d, max_d), format="YYYY-MM-DD"
        )
    accounts = sorted(tx["account"].unique())
    sel_accounts = st.multiselect("Accounts", accounts, default=accounts)
    cats = sorted(tx["category"].unique())
    sel_cats = st.multiselect("Categories", cats, default=cats)
    flow = st.radio("Flow", ["All", "Credit (in)", "Debit (out)"], horizontal=False)
    search = st.text_input("Search description")

mask = (
    (tx["date"].dt.date >= date_range[0])
    & (tx["date"].dt.date <= date_range[1])
    & (tx["account"].isin(sel_accounts))
    & (tx["category"].isin(sel_cats))
)
if flow.startswith("Credit"):
    mask &= tx["amount"] >= 0
elif flow.startswith("Debit"):
    mask &= tx["amount"] < 0
if search.strip():
    mask &= tx["description"].str.contains(search.strip(), case=False, na=False)

fdf = tx[mask].copy()

# --- KPIs --------------------------------------------------------------------
k = kpi_metrics(fdf)
c1, c2, c3, c4 = st.columns(4)
kpi_card(c1, "Total Income", money(k["income"], currency), GREEN, f'{k["count"]} transactions')
kpi_card(c2, "Total Expense", money(k["expense"], currency), RED, f'avg {money(k["avg_expense"], currency)}')
kpi_card(c3, "Net", money(k["net"], currency), ACCENT, f'savings rate {k["savings_rate"]:.0f}%')
kpi_card(c4, "Closing Balance", money(k["closing_balance"], currency), BLUE, f'{k["months"]} month(s)')

st.write("")

overview, txns, categories_tab, trends, forecast_tab, recon_tab, export_tab = st.tabs(
    ["Overview", "Transactions", "Categories", "Trends", "Forecast", "Reconciliation", "Export"]
)

# --- Overview ----------------------------------------------------------------
with overview:
    if fdf.empty:
        st.info("No transactions match the current filters.")
    else:
        msum = monthly_summary(fdf)
        left, right = st.columns([1.3, 1])
        with left:
            mlong = msum.melt(
                id_vars="month", value_vars=["income", "expense"],
                var_name="Flow", value_name="Amount",
            )
            fig = px.bar(
                mlong, x="month", y="Amount", color="Flow", barmode="group",
                color_discrete_map={"income": GREEN, "expense": RED},
                title="Income vs Expense by Month",
            )
            style_fig(fig, 360, legend_title="")
            st.plotly_chart(fig, use_container_width=True)
        with right:
            cat_exp = category_breakdown(fdf, "expense")
            if not cat_exp.empty:
                fig = px.pie(
                    cat_exp, names="category", values="total", hole=0.55,
                    color="category", color_discrete_map=CATEGORY_COLORS,
                    title="Expense by Category",
                )
                fig.update_traces(textposition="inside", textinfo="percent")
                style_fig(fig, 360)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No expenses in the current selection.")

        fig = px.line(
            msum, x="month", y="net", markers=True, title="Net Cash Flow by Month",
        )
        fig.update_traces(line_color=ACCENT)
        fig.add_hline(y=0, line_dash="dot", line_color="#9098A5")
        style_fig(fig, 300)
        st.plotly_chart(fig, use_container_width=True)

# --- Transactions (editable categories) -------------------------------------
with txns:
    st.caption(
        "Edit the **category** cell to reclassify a transaction — charts and exports "
        "update automatically."
    )
    top = st.columns([3, 1])
    with top[1]:
        if st.button("Auto-categorize (reset)", use_container_width=True):
            st.session_state.tx = categorize(st.session_state.tx.drop(columns=["category"]))
            st.rerun()

    view_cols = ["date", "account", "description", "category", "type",
                 "debit", "credit", "amount", "balance"]
    edited = st.data_editor(
        fdf[view_cols],
        hide_index=True,
        use_container_width=True,
        height=460,
        column_config={
            "date": st.column_config.DateColumn("Date", format="YYYY-MM-DD"),
            "category": st.column_config.SelectboxColumn(
                "Category", options=sorted(set(CATEGORY_OPTIONS) | set(cats)), required=True
            ),
            "debit": st.column_config.NumberColumn("Debit", format="%.2f"),
            "credit": st.column_config.NumberColumn("Credit", format="%.2f"),
            "amount": st.column_config.NumberColumn("Amount", format="%.2f"),
            "balance": st.column_config.NumberColumn("Balance", format="%.2f"),
        },
        disabled=["date", "account", "description", "type",
                  "debit", "credit", "amount", "balance"],
        key="tx_editor",
    )
    changed = edited["category"].values != fdf["category"].values
    if changed.any():
        st.session_state.tx.loc[edited.index[changed], "category"] = \
            edited["category"].values[changed]
        st.rerun()

    st.download_button(
        "Download filtered transactions (CSV)",
        fdf[view_cols].to_csv(index=False).encode(),
        file_name="transactions.csv",
        mime="text/csv",
    )

# --- Categories --------------------------------------------------------------
with categories_tab:
    if fdf.empty:
        st.info("No transactions match the current filters.")
    else:
        exp_col, inc_col = st.columns(2)
        with exp_col:
            st.subheader("Spending")
            cat_exp = category_breakdown(fdf, "expense")
            if not cat_exp.empty:
                fig = px.bar(
                    cat_exp.sort_values("total"), x="total", y="category", orientation="h",
                    color="category", color_discrete_map=CATEGORY_COLORS, text="share",
                )
                fig.update_traces(texttemplate="%{text}%", showlegend=False)
                style_fig(fig, 380, top=20, yaxis_title="", xaxis_title=currency)
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(cat_exp, hide_index=True, use_container_width=True)
            else:
                st.info("No expenses in the current selection.")
        with inc_col:
            st.subheader("Income")
            cat_inc = category_breakdown(fdf, "income")
            if not cat_inc.empty:
                fig = px.pie(cat_inc, names="category", values="total", hole=0.55,
                             color="category", color_discrete_map=CATEGORY_COLORS)
                fig.update_traces(textinfo="percent")
                style_fig(fig, 380, top=20)
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(cat_inc, hide_index=True, use_container_width=True)
            else:
                st.info("No income in the current selection.")

# --- Trends ------------------------------------------------------------------
with trends:
    if fdf.empty:
        st.info("No transactions match the current filters.")
    else:
        ordered = fdf.sort_values("date")
        fig = px.area(ordered, x="date", y="balance", title="Balance Over Time")
        fig.update_traces(line_color=ACCENT, fillcolor="rgba(91,141,239,0.15)")
        style_fig(fig, 320)
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Monthly summary")
        st.dataframe(monthly_summary(fdf), hide_index=True, use_container_width=True)

        st.subheader("Top 10 expenses")
        top_exp = (
            fdf[fdf["debit"] > 0]
            .nlargest(10, "debit")[["date", "description", "category", "debit"]]
            .reset_index(drop=True)
        )
        st.dataframe(top_exp, hide_index=True, use_container_width=True)

# --- Forecast ----------------------------------------------------------------
with forecast_tab:
    st.caption(
        "Monte Carlo projection of the combined liquid balance. We bootstrap your "
        "historical monthly net cash-flow into 2,000 simulated futures and show the "
        "P10-P90 range around the median. Uses full history, ignoring the sidebar filters."
    )
    horizon = st.slider("Months ahead", min_value=3, max_value=12, value=6, key="fc_horizon")
    res = forecast_balance(tx, horizon=horizon)
    fc, hist, summary = res["forecast"], res["history"], res["summary"]

    if not summary or summary["months_history"] < 2:
        st.info("Need at least two months of history to build a forecast.")
    else:
        if not summary["reliable"]:
            st.warning(
                f"Only {summary['months_history']} months of history — bands are wide "
                "and the projection is indicative rather than reliable."
            )

        f1, f2, f3, f4 = st.columns(4)
        kpi_card(f1, "Current balance", money(summary["current_balance"], currency), ACCENT)
        kpi_card(
            f2, f"Projected in {horizon}m (P50)",
            money(summary["proj_median_end"], currency),
            GREEN if summary["proj_median_end"] >= summary["current_balance"] else RED,
            sub="Median outcome",
        )
        kpi_card(
            f3, "Downside (P10)", money(summary["proj_p10_end"], currency), RED,
            sub="1-in-10 worse case",
        )
        kpi_card(
            f4, "Chance of going negative",
            f"{summary['prob_negative'] * 100:.0f}%",
            RED if summary["prob_negative"] > 0.1 else GREEN,
            sub=f"Within {horizon} months",
        )

        anchor_month = hist["month"].iloc[-1]
        anchor_val = summary["current_balance"]
        xs = [anchor_month] + list(fc["month"])
        p50s = [anchor_val] + list(fc["p50"])
        p10s = [anchor_val] + list(fc["p10"])
        p90s = [anchor_val] + list(fc["p90"])

        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=xs + xs[::-1], y=p90s + p10s[::-1], fill="toself",
                fillcolor=THEME["accent_soft"], line=dict(width=0),
                hoverinfo="skip", name="P10-P90",
            )
        )
        fig.add_trace(
            go.Scatter(
                x=hist["month"], y=hist["balance"], mode="lines+markers",
                line=dict(color=THEME["muted"], width=2), name="History",
            )
        )
        fig.add_trace(
            go.Scatter(
                x=xs, y=p50s, mode="lines+markers",
                line=dict(color=ACCENT, width=2, dash="dash"), name="Forecast (P50)",
            )
        )
        if summary["proj_p10_end"] < 0:
            fig.add_hline(y=0, line_dash="dot", line_color=RED)
        style_fig(fig, 380, legend_title="", title="Balance Forecast")
        st.plotly_chart(fig, use_container_width=True)

        table = fc.rename(
            columns={"month": "Month", "p10": "P10", "p50": "P50 (median)", "p90": "P90"}
        )[["Month", "P10", "P50 (median)", "P90"]].copy()
        table["Month"] = pd.to_datetime(table["Month"]).dt.strftime("%b %Y")
        for col in ("P10", "P50 (median)", "P90"):
            table[col] = table[col].round(0)
        st.dataframe(table, hide_index=True, use_container_width=True)


# --- Reconciliation ----------------------------------------------------------
with recon_tab:
    st.caption(
        "Replays every transaction from each statement's opening balance. "
        "`closing_ok = True` means the debit/credit classification is correct end-to-end."
    )
    if report.empty:
        st.info("No reconciliation data available.")
    else:
        all_ok = bool(report["closing_ok"].all())
        (st.success if all_ok else st.error)(
            "All statements reconciled successfully." if all_ok
            else "One or more statements did not reconcile — check the flagged file."
        )
        st.dataframe(report, hide_index=True, use_container_width=True)

# --- Export ------------------------------------------------------------------
with export_tab:
    st.caption("Download a styled Excel report for the current filter selection.")
    period_label = f"{date_range[0]} to {date_range[1]}"
    if fdf.empty:
        st.info("Nothing to export for the current filters.")
    else:
        msum = monthly_summary(fdf)
        cat_exp = category_breakdown(fdf, "expense")
        export_cols = ["date", "month", "account", "description", "category", "type",
                       "debit", "credit", "amount", "balance", "source_file"]
        xlsx = build_excel_report(
            fdf[export_cols], msum, cat_exp, kpi_metrics(fdf),
            currency=currency, period_label=period_label,
        )
        st.download_button(
            "Download Excel report",
            xlsx,
            file_name="FinLens-Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.write(f"**Rows in report:** {len(fdf)}  •  **Period:** {period_label}")
