"""FinLens — build the Excel dashboard from bank-statement PDFs, straight from the shell.

Examples
--------
    python cli.py                          # interactive: prompts for the PDF(s)
    python cli.py statement.pdf            # one file
    python cli.py a.pdf b.pdf -o out.xlsx  # several files, custom output

Password-protected PDFs are detected automatically and you'll be asked for the
password (input is hidden). The output defaults to
``data/output/Bank-Statement-Analysis.xlsx``.
"""

from __future__ import annotations

import argparse
import getpass
import glob
from pathlib import Path

import pandas as pd

from finlib import (
    build_dashboard_workbook,
    categorize,
    monthly_summary,
    parse_statement,
    reconcile,
    statement_to_dataframe,
)

try:
    from pdfminer.pdfdocument import PDFPasswordIncorrect
except Exception:  # pragma: no cover - fallback if pdfminer internals move
    PDFPasswordIncorrect = Exception

DEFAULT_OUTPUT = Path("data/output/Bank-Statement-Analysis.xlsx")


def _expand(patterns: list[str]) -> list[str]:
    """Expand globs / ~ and de-duplicate while preserving order."""
    files: list[str] = []
    for p in patterns:
        matches = glob.glob(str(Path(p).expanduser()))
        files.extend(matches or [p])
    seen: set[str] = set()
    out: list[str] = []
    for f in files:
        if f not in seen:
            seen.add(f)
            out.append(f)
    return out


def _prompt_files() -> list[str]:
    raw = input("Enter path to statement PDF(s) (space or comma separated): ").strip()
    parts = [p for chunk in raw.split(",") for p in chunk.split()]
    return _expand(parts)


def _parse_with_password(path: str):
    """Parse one PDF, prompting for a password if the file is protected."""
    name = Path(path).name
    try:
        return parse_statement(path, source_name=name)
    except PDFPasswordIncorrect:
        for _ in range(3):
            pw = getpass.getpass(f"  Password for {name}: ")
            try:
                return parse_statement(path, password=pw, source_name=name)
            except PDFPasswordIncorrect:
                print("  Incorrect password, try again.")
        raise SystemExit(f"Could not unlock {path}.")


def main(argv: list[str] | None = None) -> None:
    ap = argparse.ArgumentParser(
        description="FinLens — build an Excel dashboard from bank-statement PDFs."
    )
    ap.add_argument(
        "files", nargs="*", help="PDF statement path(s). If omitted, you'll be prompted."
    )
    ap.add_argument(
        "-o", "--output", default=None,
        help=f"Output .xlsx path (default: {DEFAULT_OUTPUT}).",
    )
    args = ap.parse_args(argv)

    files = _expand(args.files) if args.files else _prompt_files()
    if not files:
        raise SystemExit("No files provided.")

    frames: list[pd.DataFrame] = []
    reports: list[dict] = []
    for path in files:
        if not Path(path).is_file():
            print(f"Skipping (not found): {path}")
            continue
        print(f"Parsing {path} ...")
        stmt = _parse_with_password(path)
        df = statement_to_dataframe(stmt)
        if df.empty:
            print(f"  No transactions found in {path}.")
            continue
        df = categorize(df)
        frames.append(df)
        reports.append(reconcile(stmt, df))
        print(f"  {len(df)} transactions ({str(stmt.metadata.get('layout', '')).upper()}).")

    if not frames:
        raise SystemExit("Nothing to export.")

    transactions = pd.concat(frames, ignore_index=True)
    monthly = monthly_summary(transactions)
    report = pd.DataFrame(reports)

    out_path = Path(args.output).expanduser() if args.output else DEFAULT_OUTPUT
    out_path.parent.mkdir(parents=True, exist_ok=True)
    saved = build_dashboard_workbook(transactions, monthly, report, out_path)
    print(f"\nSaved dashboard: {saved.resolve()}")


if __name__ == "__main__":
    main()
