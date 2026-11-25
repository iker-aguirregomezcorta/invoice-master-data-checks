"""
Step 5 - Report of validated invoices by date.

The report can filter by either date:
  - "validation": when the invoice passed the control (stored in the ledger).
  - "invoice":    the invoice issue date (useful for month-end close).
And by a from/to range: a single day is from = to.
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .ledger import LEDGER_COLUMNS, MONEY

HEADER_FILL = PatternFill("solid", start_color="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def filter_ledger(ledger_path: Path, start: date, end: date, date_type: str = "validation") -> list[tuple]:
    """Ledger rows whose date (validation or invoice) falls within [start, end]."""
    if date_type not in ("validation", "invoice"):
        raise ValueError("date_type must be 'validation' or 'invoice'")
    if start > end:
        raise ValueError("The start date is after the end date")
    if not ledger_path.exists():
        raise FileNotFoundError("There are no validated invoices yet (the ledger does not exist).")

    col = 13 if date_type == "validation" else 1  # 0-based index in the row
    wb = openpyxl.load_workbook(ledger_path, read_only=True)
    rows = [
        r for r in wb.active.iter_rows(min_row=2, values_only=True)
        if r and r[0] and (d := _as_date(r[col])) and start <= d <= end
    ]
    wb.close()
    rows.sort(key=lambda r: (_as_date(r[col]), str(r[0])))
    return rows


def build_report(ledger_path: Path, reports_folder: Path,
                 start: date, end: date, date_type: str = "validation") -> tuple[Path, int]:
    rows = filter_ledger(ledger_path, start, end, date_type)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    date_name = "validation date" if date_type == "validation" else "invoice date"
    ws["A1"] = "Validated invoices report"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Filter: {date_name} from {start:%d/%m/%Y} to {end:%d/%m/%Y}"
    ws["A3"] = f"Generated on {datetime.now():%d/%m/%Y %H:%M} · {len(rows)} invoice(s)"

    header_row = 5
    for i, (name, width) in enumerate(LEDGER_COLUMNS, start=1):
        c = ws.cell(header_row, i, name)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = width

    for r in rows:
        ws.append(list(r))
        n = ws.max_row
        ws.cell(n, 2).number_format = "DD/MM/YYYY"
        for c in (10, 11, 12):
            ws.cell(n, c).number_format = MONEY
        ws.cell(n, 14).number_format = "DD/MM/YYYY HH:MM"

    if rows:
        n = ws.max_row + 1
        ws.cell(n, 9, "TOTAL").font = Font(bold=True)
        for c in (10, 11, 12):
            letter = get_column_letter(c)
            cell = ws.cell(n, c, f"=SUM({letter}{header_row + 1}:{letter}{n - 1})")
            cell.font = Font(bold=True)
            cell.number_format = MONEY
    else:
        ws.cell(header_row + 1, 1, "No validated invoices in that period.")

    ws.freeze_panes = ws.cell(header_row + 1, 1)
    reports_folder.mkdir(parents=True, exist_ok=True)
    path = reports_folder / f"report_{date_type}_{start:%Y%m%d}_{end:%Y%m%d}.xlsx"
    try:
        wb.save(path)
    except PermissionError:
        raise PermissionError(
            f"Cannot save {path.name}. It is probably open in Excel: close it and try again."
        ) from None
    return path, len(rows)
