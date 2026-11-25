"""
Process orchestration, shared by the command line and the web app.

Both interfaces call these same functions, so the matching rule and the
ledger are identical whichever way the application is used.
"""
from __future__ import annotations

import json
import sys
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path

import openpyxl

from .alerts import send_alert
from .extraction import ExtractionError, extract_invoice
from .ledger import (alerted_exceptions, log_exceptions, log_read_error,
                     log_validated, logged_invoices)
from .validation import load_master, validate


def base_folder() -> Path:
    # Works the same as a script or packaged as an .exe (PyInstaller).
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


BASE = base_folder()


def load_config() -> dict:
    path = BASE / "config.json"
    cfg = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    paths = cfg.setdefault("paths", {})
    paths.setdefault("customer_master", "data/customer_master.xlsx")
    paths.setdefault("invoices_folder", "data/invoices")
    paths.setdefault("output_folder", "output")
    cfg.setdefault("email", {"enabled": False})
    return cfg


def resolve(p: str | Path) -> Path:
    p = Path(p)
    return p if p.is_absolute() else BASE / p


@dataclass
class Paths:
    master: Path
    invoices: Path
    output: Path

    @property
    def ledger(self) -> Path:
        return self.output / "validated_invoices.xlsx"

    @property
    def exceptions(self) -> Path:
        return self.output / "exceptions.xlsx"

    @property
    def alerts(self) -> Path:
        return self.output / "alerts"

    @property
    def reports(self) -> Path:
        return self.output / "reports"

    @property
    def uploads(self) -> Path:
        return self.output / "uploads"


def paths_from(cfg: dict) -> Paths:
    p = cfg["paths"]
    return Paths(resolve(p["customer_master"]), resolve(p["invoices_folder"]), resolve(p["output_folder"]))


class FileLocked(Exception):
    pass


def check_writable(*files: Path) -> None:
    """
    On Windows, an Excel file that is open is locked and cannot be written.
    This is checked BEFORE processing anything, so a run is never left half done.
    """
    locked = []
    for f in files:
        if f.exists():
            try:
                with open(f, "r+b"):
                    pass
            except PermissionError:
                locked.append(f.name)
    if locked:
        raise FileLocked(
            "Cannot write to " + ", ".join(locked)
            + ". It is probably open in Excel: close it and try again."
        )


def process(pdfs: list[Path], cfg: dict, now: datetime | None = None) -> list[dict]:
    """
    Check a list of PDFs and return one result per invoice:
      status = "validated" | "mismatch" | "duplicate" | "unreadable"
    An exception already alerted (same invoice, field and values) is not
    recorded or alerted again; if anything changes, it is a new exception.
    """
    now = now or datetime.now().replace(microsecond=0)
    pt = paths_from(cfg)
    check_writable(pt.ledger, pt.exceptions)

    master = load_master(pt.master)
    already = logged_invoices(pt.ledger)
    alerted = alerted_exceptions(pt.exceptions)
    results = []

    for pdf in pdfs:
        res = {"file": pdf.name, "number": None, "date": None, "customer": None,
               "vat_id": None, "total": None, "status": None, "mismatches": [],
               "alert": None, "alert_file": None, "reason": None}
        try:
            data = extract_invoice(pdf)
        except ExtractionError as e:
            log_read_error(pt.exceptions, pdf.name, now, str(e))
            res.update(status="unreadable", reason=str(e))
            results.append(res)
            continue
        except Exception as e:  # damaged PDF or not a PDF at all
            log_read_error(pt.exceptions, pdf.name, now, f"Cannot open: {e}")
            res.update(status="unreadable", reason=f"The file cannot be opened ({e.__class__.__name__})")
            results.append(res)
            continue

        res.update(number=data.number, date=data.date.isoformat(), customer=data.company_name,
                   vat_id=data.vat_id, total=data.total)

        if data.number in already:
            res["status"] = "duplicate"
            results.append(res)
            continue

        r = validate(data, master)
        if r.is_valid:
            log_validated(pt.ledger, r, now)
            already.add(data.number)
            res["status"] = "validated"
        else:
            def key(m):
                return (data.number, m.field, str(m.invoice_value), str(m.master_value))

            new = [m for m in r.mismatches if key(m) not in alerted]
            if new:
                # The email lists all current mismatches; the Excel log only the new ones.
                text, eml = send_alert(r, cfg["email"], pt.alerts)
                log_exceptions(pt.exceptions, replace(r, mismatches=new), now, text)
                for m in new:
                    alerted[key(m)] = now
                alert_file = eml.name if eml else None
            else:
                dates = [alerted[key(m)] for m in r.mismatches if isinstance(alerted[key(m)], datetime)]
                text = (f"Already alerted on {min(dates):%d/%m/%Y %H:%M}; not alerted again"
                        if dates else "Already alerted; not alerted again")
                alert_file = None
            res.update(
                status="mismatch", alert=text, alert_file=alert_file,
                mismatches=[{"field": m.field, "invoice": m.invoice_value, "master": m.master_value}
                            for m in r.mismatches],
            )
        results.append(res)
    return results


# ---------------------------------------------------------------- reading

def cell(v):
    if isinstance(v, datetime):
        return v.isoformat(sep=" ", timespec="minutes") if (v.hour or v.minute) else v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return v


def read_table(xlsx_path: Path) -> list[dict]:
    """Return the rows of a ledger/exceptions Excel file as dictionaries."""
    if not xlsx_path.exists():
        return []
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    header = [str(c) for c in rows[0]]
    return [dict(zip(header, map(cell, r))) for r in rows[1:] if any(c is not None for c in r)]


def master_summary(cfg: dict) -> dict:
    pt = paths_from(cfg)
    m = load_master(pt.master)
    return {"file": pt.master.name, "customers": len(m), "missing_vat_id": sum(1 for c in m if not c.vat_id)}
