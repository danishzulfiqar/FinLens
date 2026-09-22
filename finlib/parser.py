"""PDF statement parsing — ported from the notebook, extended to accept file objects.

Amounts are placed into Debit / Credit / Balance purely by their x-position, so we
never guess whether a value is money-in or money-out. New banks can be supported by
adding an entry to ``LAYOUTS``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import IO

import pdfplumber

# Money token, e.g. "1,234.56" or "-1,234.56".
MONEY_RE = re.compile(r"^-?[\d,]+\.\d{2}$")

# Per-bank layouts. `anchors` are the right-edge x-coordinates of each amount column
# (amounts are right-aligned). `detect` picks the layout from the first page's text.
LAYOUTS: dict[str, dict] = {
    "myabl": {
        "detect": lambda t: "Opening Balance:" in t and "Account Statement" in t,
        "anchors": {"debit": 372.68, "credit": 471.89, "balance": 571.10},
        "date_re": re.compile(r"^\d{2}\s[A-Z][a-z]{2}\s\d{4}$"),
        "date_fmt": "%d %b %Y",
        "date_x_max": 100.0,
        "desc_x0_min": 100.0,
        "desc_x1_max": 340.0,
        "header_bottom": 205.0,
        "footer_re": re.compile(r"^\d+\s+\d{2}\s\w{3}\s\d{4},"),
        "order": "asc",
    },
    "hbl": {
        "detect": lambda t: "HBL Mobile" in t or "Account Activity generated" in t,
        "anchors": {"debit": 375.0, "credit": 446.0, "balance": 550.0},
        "date_re": re.compile(r"^\d{2}-\d{2}-\d{4}$"),
        "date_fmt": "%d-%m-%Y",
        "date_x_max": 90.0,
        "desc_x0_min": 150.0,
        "desc_x1_max": 340.0,
        "header_bottom": 218.0,
        "footer_re": None,
        "order": "desc",
    },
}


@dataclass
class Statement:
    """Parsed contents of one statement PDF."""

    metadata: dict = field(default_factory=dict)
    transactions: list[dict] = field(default_factory=list)


def _detect_layout(page_text: str) -> str:
    for name, cfg in LAYOUTS.items():
        if cfg["detect"](page_text):
            return name
    raise ValueError("Unrecognised statement layout (not myABL or HBL).")


def _classify_amount(x1: float, anchors: dict) -> str:
    debit, credit, balance = anchors["debit"], anchors["credit"], anchors["balance"]
    if x1 < (debit + credit) / 2:
        return "debit"
    if x1 < (credit + balance) / 2:
        return "credit"
    return "balance"


def _group_words_into_lines(words: list[dict], tol: float = 3.0) -> list[list[dict]]:
    lines: dict[int, list[dict]] = {}
    for w in words:
        lines.setdefault(round(w["top"] / tol), []).append(w)
    return [sorted(lines[k], key=lambda w: w["x0"]) for k in sorted(lines)]


def _extract_metadata(page, layout_name: str) -> dict:
    text = page.extract_text() or ""
    meta: dict = {}
    if layout_name == "myabl":
        patterns = {
            "account_number": r"Account Number:\s*(\S+)",
            "account_title": r"Account Title:\s*(.+)",
            "currency": r"Currency:\s*(\S+)",
            "opening_balance": r"Opening Balance:\s*([\d,]+\.\d{2})",
            "closing_balance": r"Closing Balance:\s*([\d,]+\.\d{2})",
        }
        for key, pat in patterns.items():
            m = re.search(pat, text)
            meta[key] = m.group(1).strip() if m else None
    else:  # hbl
        m_title = re.search(r"AccountTitle:(.+)", text)
        m_iban = re.search(r"IBAN:(\S+)", text)
        m_row = re.search(
            r"(\d{6,})\s+\d+\s+([A-Z]{3})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", text
        )
        meta["account_number"] = m_row.group(1) if m_row else None
        meta["account_title"] = m_title.group(1).strip() if m_title else None
        meta["currency"] = m_row.group(2) if m_row else None
        meta["opening_balance"] = m_row.group(3) if m_row else None
        meta["closing_balance"] = m_row.group(4) if m_row else None
        meta["iban"] = m_iban.group(1) if m_iban else None
    return meta


def parse_statement(
    source: str | Path | IO[bytes],
    password: str | None = None,
    source_name: str | None = None,
) -> Statement:
    """Extract account metadata and transactions from a statement PDF.

    ``source`` may be a path or a binary file-like object (e.g. an upload).
    """
    is_fileobj = hasattr(source, "read")
    if source_name:
        name = source_name
    elif is_fileobj:
        name = getattr(source, "name", "statement.pdf")
    else:
        name = Path(source).name

    stmt = Statement()
    pdf = pdfplumber.open(source if is_fileobj else str(source), password=password)
    with pdf:
        layout_name = _detect_layout(pdf.pages[0].extract_text() or "")
        layout = LAYOUTS[layout_name]
        anchors = layout["anchors"]

        stmt.metadata = _extract_metadata(pdf.pages[0], layout_name)
        stmt.metadata.update(
            source_file=name,
            pages=len(pdf.pages),
            layout=layout_name,
            date_fmt=layout["date_fmt"],
            order=layout["order"],
        )

        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False)
            header_tops = [w["top"] for w in words if w["text"] == "Description"]
            start_y = max(header_tops) if header_tops else layout["header_bottom"]

            for row in _group_words_into_lines(words):
                if not row or row[0]["top"] <= start_y + 2:
                    continue
                line_text = " ".join(w["text"] for w in row)
                if layout["footer_re"] and layout["footer_re"].match(line_text):
                    continue

                date_str = " ".join(
                    w["text"] for w in row if w["x0"] < layout["date_x_max"]
                )
                desc = " ".join(
                    w["text"]
                    for w in row
                    if layout["desc_x0_min"] <= w["x0"] and w["x1"] <= layout["desc_x1_max"]
                )
                amounts = {
                    _classify_amount(w["x1"], anchors): w["text"]
                    for w in row
                    if MONEY_RE.match(w["text"])
                }

                if layout["date_re"].match(date_str):
                    stmt.transactions.append(
                        {
                            "date": date_str,
                            "description": desc,
                            "debit": amounts.get("debit"),
                            "credit": amounts.get("credit"),
                            "balance": amounts.get("balance"),
                        }
                    )
                elif stmt.transactions and desc:
                    stmt.transactions[-1]["description"] += " " + desc

    return stmt
