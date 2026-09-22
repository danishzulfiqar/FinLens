# FinLens

Turn PDF bank statements into an interactive dashboard and a styled Excel report.
Upload one or more statements, and the app parses every transaction, categorizes it,
reconciles the running balance, and lets you filter, edit, and export the results.

Supported statement layouts out of the box: **myABL** (Allied Bank) and **HBL**
(HBL Mobile / Konnect). Password-protected PDFs are supported.

---

## Features

- Upload multiple PDF statements at once (password-protected files supported)
- Automatic transaction categorization (editable per row)
- Overview, category, and trend charts
- Balance reconciliation (verifies debit/credit classification end-to-end)
- Light and dark mode
- Export a styled multi-sheet Excel report or CSV for any filter selection

---

## Project structure

```
Fin-Analysis/
├── app.py                     # Streamlit web app (entry point)
├── cli.py                     # Shell command: PDFs -> Excel dashboard
├── finlib/                    # Reusable library
│   ├── parser.py              # PDF -> Statement (per-bank layouts)
│   ├── analysis.py            # Categorization, reconciliation, KPIs
│   ├── excel.py               # Styled Excel report builder
│   └── dashboard.py           # Interactive Excel dashboard builder
├── notebooks/
│   └── PDF-Extraction.ipynb   # Batch/offline pipeline for local PDFs
├── data/
│   ├── statements/            # Put your local PDF statements here (git-ignored)
│   └── output/                # Generated Excel reports land here (git-ignored)
├── .streamlit/config.toml     # Theme + upload-size settings
└── requirements.txt
```

---

## Setup

Requires Python 3.10+.

```bash
# From the project root
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## Run the web app

```bash
streamlit run app.py
```

Then open the URL Streamlit prints (usually http://localhost:8501).

### Using the app

1. **Upload** one or more PDF statements in the sidebar.
2. If a statement is password-protected, type the password in the sidebar
   (it is used only in memory and never saved).
3. Use the sidebar **filters** — date range, account, category, flow (in/out),
   and free-text search — to narrow the view.
4. **Overview / Categories / Trends** tabs show charts for the current selection.
5. In the **Transactions** tab you can edit any row's category inline, or click
   **Auto-categorize (reset)** to re-run the rules. Download the table as CSV.
6. The **Reconciliation** tab confirms each statement balances from its opening
   figure to its closing figure.
7. The **Export** tab builds a styled Excel report for the current filters.
8. Toggle **Dark mode** at the top of the sidebar any time.

> **Note on dark mode:** the toggle restyles the app surfaces, widgets, and charts.
> Streamlit's native data-grid tables follow the theme in `.streamlit/config.toml`.
> For fully native dark tables you can also pick **Dark** from the app's
> **menu (top-right) > Settings**.

---

## Build the Excel dashboard from the shell

Prefer the terminal? `cli.py` produces the same interactive Excel dashboard as the
notebook — no browser needed.

```bash
# Interactive: it asks you for the PDF(s)
python cli.py

# Or pass files directly
python cli.py "data/statements/statement.pdf"
python cli.py a.pdf b.pdf -o my-report.xlsx
```

- You can pass several PDFs (space or comma separated); globs like `data/statements/*.pdf` work too.
- Password-protected PDFs are detected automatically and you'll be prompted for the
  password (typing is hidden).
- Output defaults to `data/output/Bank-Statement-Analysis.xlsx`; override with `-o`.

The workbook has a **Dashboard** sheet with Period/Account dropdown filters, KPI cards,
and charts, all driven by native Excel formulas, plus **Transactions**,
**Monthly Summary**, and **Reconciliation** sheets.

---

## Batch / offline workflow (notebook)

`notebooks/PDF-Extraction.ipynb` runs the same pipeline over local files without the
web UI — handy for bulk processing.

1. Drop your PDFs into `data/statements/`.
2. For protected files, add an entry to the `PASSWORDS` dict in the config cell,
   e.g. `"Account Statement.pdf": "905003"`.
3. Run all cells. The styled workbook is written to
   `data/output/Bank-Statement-Analysis.xlsx`.

The notebook resolves the project root automatically, so it works regardless of the
folder you launch it from.

---

## Adding a new bank

Statement layouts are defined declaratively in `finlib/parser.py` (`LAYOUTS`). Add a
new entry with a `detect` function, the amount-column x-anchors, and the date format
for your bank, and the parser will pick it up automatically.
