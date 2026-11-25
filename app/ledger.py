"""
Step 3 - Save the results to Excel.

  - validated_invoices.xlsx: one row per invoice that matched exactly, with its
    VALIDATION TIMESTAMP (when it passed the control). An invoice already in the
    ledger is never added again (duplicate check by invoice number).
  - exceptions.xlsx: one row per mismatching field. Without it, rejected
    invoices would only exist in an email. This way there is a trace of what
    failed, when, and whether an alert was sent.
    The same exception (same invoice, field and values) is recorded only once.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .validation import ValidationResult

LEDGER_COLUMNS = [
    ("Invoice no.", 14), ("Invoice date", 13), ("Customer ID", 13),
    ("Company name", 34), ("VAT ID", 18), ("Street", 30), ("Postcode", 11),
    ("Town", 22), ("Country", 9), ("Taxable base", 15), ("VAT amount", 13),
    ("Total", 15), ("File", 16), ("Validated at", 18),
]

EXCEPTION_COLUMNS = [
    ("Detected at", 18), ("Invoice no.", 14), ("Invoice date", 13),
    ("File", 16), ("Field", 15), ("Invoice value", 34),
    ("Master value", 40), ("Alert", 40),
]

HEADER_FILL = PatternFill("solid", start_color="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY = '#,##0.00 "€"'


def _create_workbook(path: Path, sheet_title: str, columns) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    for i, (name, width) in enumerate(columns, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _open(path: Path, sheet_title: str, columns):
    if not path.exists():
        _create_workbook(path, sheet_title, columns)
    return openpyxl.load_workbook(path)


def _update_filter(ws, n_columns: int) -> None:
    ws.auto_filter.ref = f"A1:{get_column_letter(n_columns)}{ws.max_row}"


def logged_invoices(path: Path) -> set[str]:
    if not path.exists():
        return set()
    wb = openpyxl.load_workbook(path, read_only=True)
    numbers = {str(r[0]) for r in wb.active.iter_rows(min_row=2, values_only=True) if r and r[0]}
    wb.close()
    return numbers


def alerted_exceptions(path: Path) -> dict[tuple, datetime]:
    """
    Exceptions already recorded: (invoice no., field, invoice value, master value)
    -> when they were first detected.
    Used to avoid recording and alerting the same exception on every run.
    """
    if not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, read_only=True)
    seen = {}
    for r in wb.active.iter_rows(min_row=2, values_only=True):
        if r and r[1]:
            key = (str(r[1]), str(r[4]), str(r[5]), str(r[6]))
            seen.setdefault(key, r[0])
    wb.close()
    return seen


def log_validated(path: Path, r: ValidationResult, when: datetime) -> None:
    wb = _open(path, "Validated invoices", LEDGER_COLUMNS)
    ws = wb.active
    f = r.invoice
    ws.append([
        f.number, f.date, f.customer_id, f.company_name, f.vat_id, f.street,
        f.postcode, f.town, f.country, f.taxable_base, f.vat_amount,
        f.total, f.file, when,
    ])
    row = ws.max_row
    ws.cell(row, 2).number_format = "DD/MM/YYYY"
    for col in (10, 11, 12):
        ws.cell(row, col).number_format = MONEY
    ws.cell(row, 14).number_format = "DD/MM/YYYY HH:MM"
    _update_filter(ws, len(LEDGER_COLUMNS))
    wb.save(path)


def log_exceptions(path: Path, r: ValidationResult, when: datetime, alert: str) -> None:
    wb = _open(path, "Exceptions", EXCEPTION_COLUMNS)
    ws = wb.active
    f = r.invoice
    for m in r.mismatches:
        ws.append([when, f.number, f.date, f.file, m.field, m.invoice_value, m.master_value, alert])
        row = ws.max_row
        ws.cell(row, 1).number_format = "DD/MM/YYYY HH:MM"
        ws.cell(row, 3).number_format = "DD/MM/YYYY"
    _update_filter(ws, len(EXCEPTION_COLUMNS))
    wb.save(path)


def log_read_error(path: Path, file: str, when: datetime, reason: str) -> None:
    wb = _open(path, "Exceptions", EXCEPTION_COLUMNS)
    ws = wb.active
    ws.append([when, "", None, file, "PDF reading", reason, "", "Manual review"])
    ws.cell(ws.max_row, 1).number_format = "DD/MM/YYYY HH:MM"
    _update_filter(ws, len(EXCEPTION_COLUMNS))
    wb.save(path)
