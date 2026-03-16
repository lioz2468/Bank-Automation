"""
Parser for Bank Hapoalim transaction export files.

Bank Hapoalim online banking ("ניהול חשבון" → "תנועות בחשבון") lets you
download transaction history as an Excel workbook.  Two variants exist:

  Standard export  – columns: תאריך, תאריך ערך, פרטי פעולה, אסמכתא,
                               חובה, זכות, יתרה
  Extended export  – adds:    קוד פעולה, צרור

This parser auto-detects whichever columns are present, so both variants
work.  Plain CSV files (UTF-8 or Windows-1255) are also accepted.

Date formats handled: DD/MM/YYYY, DD/MM/YY, DD.MM.YYYY, DD.MM.YY
Number format: Israeli convention — comma thousands-separator, period decimal
               (e.g. "1,234,567.89").  Parentheses for negatives are also
               handled (e.g. "(1,234.56)" → -1234.56).
"""

from __future__ import annotations

import csv
import re
from datetime import datetime as _datetime, date as _date
from typing import Dict, List, Optional

import openpyxl

from models.statement import BankTransaction

# ── Column-name → BankTransaction field mapping ───────────────────────────────
# Keys are normalised header strings (stripped of whitespace and RTL markers).
# Multiple aliases per field handle variations across export types.

_COL_MAP: Dict[str, str] = {
    # date
    "תאריך":              "date",
    "תאריך ביצוע":        "date",
    # operation code  (extended export)
    "קוד פעולה":          "code",
    "קוד":                "code",
    # operation description
    "פרטי פעולה":         "operation",
    "תיאור פעולה":        "operation",
    "הפעולה":             "operation",
    "פרטים":              "operation",
    "תיאור":              "operation",
    "שם פעולה":           "operation",
    # details / sub-description
    "פרטי המסמך":         "details",
    "פרטים נוספים":       "details",
    "תאור נוסף":          "details",
    # reference
    "אסמכתא":             "reference",
    "מסמך":               "reference",
    "מספר מסמך":          "reference",
    # batch  (extended export)
    "צרור":               "batch",
    "אצווה":              "batch",
    # debit
    "חובה":               "debit",
    "חיוב":               "debit",
    'חובה (-)':           "debit",
    # credit
    "זכות":               "credit",
    "זיכוי":              "credit",
    'זכות (+)':           "credit",
    # running balance
    "יתרה":               "balance",
    'יתרה בש"ח':          "balance",
    "יתרה בש''ח":         "balance",   # two single-quotes variant
    "יתרה לאחר פעולה":    "balance",
    "יתרה נוכחית":        "balance",
}

# RTL / BOM / invisible Unicode characters that may appear in headers
_STRIP_CHARS = "\u200f\u200e\u202b\u202a\ufeff\u00a0"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _normalise(text: str) -> str:
    return text.strip(_STRIP_CHARS).strip()


def _to_str(val) -> str:
    # openpyxl returns date cells as datetime / date objects — format them
    # as DD.MM.YYYY so _normalise_date can parse them correctly.
    if isinstance(val, _datetime):
        return val.strftime("%d.%m.%Y")
    if isinstance(val, _date):
        return val.strftime("%d.%m.%Y")
    return "" if val is None else _normalise(str(val))


def _to_float(val) -> Optional[float]:
    """Convert an Israeli-formatted number cell to float, or None."""
    if val is None:
        return None
    s = _normalise(str(val)).replace(",", "").replace(" ", "")
    if not s or s in ("-", "–", "—"):
        return None
    # Parentheses = negative  e.g. "(1234.56)"
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        result = float(s)
        return -result if negative else result
    except ValueError:
        return None


def _normalise_date(raw: str) -> str:
    """
    Normalise a date string to DD.MM.YYYY.
    Accepts DD/MM/YYYY, DD/MM/YY, DD.MM.YYYY, DD.MM.YY.
    Returns the original string unchanged if it doesn't match.
    """
    m = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$", raw.strip())
    if not m:
        return raw
    d, mo, y = m.group(1), m.group(2), m.group(3)
    if len(y) == 2:
        y = ("20" + y) if int(y) < 70 else ("19" + y)
    return f"{int(d):02d}.{int(mo):02d}.{y}"


def _map_headers(raw_headers: list) -> Dict[int, str]:
    """Return {column_index: field_name} for every recognised header."""
    mapping: Dict[int, str] = {}
    for i, h in enumerate(raw_headers):
        norm = _normalise(_to_str(h))
        field = _COL_MAP.get(norm)
        if field and field not in mapping.values():   # first match wins
            mapping[i] = field
    return mapping


def _row_to_tx(row: list, mapping: Dict[int, str]) -> Optional[BankTransaction]:
    """Convert one data row to a BankTransaction, or None if the row is empty/invalid."""
    fields: Dict[str, object] = {
        "date": "", "code": "", "operation": "", "details": "",
        "reference": "", "batch": "", "debit": None, "credit": None, "balance": None,
    }
    for col_idx, field in mapping.items():
        if col_idx < len(row):
            raw = row[col_idx]
            if field in ("debit", "credit", "balance"):
                fields[field] = _to_float(raw)
            else:
                fields[field] = _to_str(raw)

    date_str = str(fields["date"]).strip()
    if not date_str:
        return None   # skip header-repetition or totals rows

    fields["date"] = _normalise_date(date_str)
    return BankTransaction(**fields)  # type: ignore[arg-type]


def _find_header_row(rows: list) -> tuple[Optional[int], Dict[int, str]]:
    """
    Scan rows top-to-bottom for the first row that contains at least one
    recognised column name.  Returns (row_index, mapping) or (None, {}).
    """
    for i, row in enumerate(rows):
        mapping = _map_headers([_to_str(c) for c in row])
        if mapping:
            return i, mapping
    return None, {}


# ── Public API ────────────────────────────────────────────────────────────────

def parse_transactions(file_path: str) -> List[BankTransaction]:
    """
    Parse a Bank Hapoalim transaction export file (Excel or CSV).

    Returns a list of BankTransaction objects in the order they appear in
    the file (typically newest-first, matching the bank's default sort).
    Unknown columns are silently ignored; missing optional columns default
    to empty string / None.
    """
    if file_path.lower().endswith(".csv"):
        return _parse_csv(file_path)
    return _parse_excel(file_path)


def _parse_excel(file_path: str) -> List[BankTransaction]:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]

    header_idx, mapping = _find_header_row(rows)
    if header_idx is None:
        return []

    txs: List[BankTransaction] = []
    for row in rows[header_idx + 1:]:
        tx = _row_to_tx(row, mapping)
        if tx:
            txs.append(tx)
    return txs


def _parse_csv(file_path: str) -> List[BankTransaction]:
    # Try UTF-8-with-BOM first (common for Israeli bank exports), then Windows-1255
    for encoding in ("utf-8-sig", "windows-1255", "utf-8"):
        try:
            with open(file_path, encoding=encoding, newline="") as f:
                rows = list(csv.reader(f))
            break
        except (UnicodeDecodeError, LookupError):
            continue
    else:
        return []

    header_idx, mapping = _find_header_row(rows)
    if header_idx is None:
        return []

    txs: List[BankTransaction] = []
    for row in rows[header_idx + 1:]:
        tx = _row_to_tx(row, mapping)
        if tx:
            txs.append(tx)
    return txs
