"""
Invoice checks against the customer master file (command line).

For the web app: python web.py  (or double-click start_web.bat)

Usage:
  python invoice_checks.py                          -> interactive menu
  python invoice_checks.py validate [folder]        -> checks every PDF in the folder
  python invoice_checks.py report --from 2026-10-01 --to 2026-10-31 [--by invoice]

What "validate" does, for each PDF:
  1. Extract the customer details from the invoice     (app/extraction.py)
  2. Compare them with the customer master file        (app/validation.py)
  3a. If they match -> log with the validation timestamp (app/ledger.py)
  3b. If not        -> email alert + exception record   (app/alerts.py, app/ledger.py)
Orchestration lives in app/service.py, shared with the web app.
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from app.report import build_report
from app.service import FileLocked, load_config, paths_from, process, resolve

STATUSES = ("validated", "mismatch", "duplicate", "unreadable")


def cmd_validate(folder: str | None, cfg: dict) -> list[dict]:
    pt = paths_from(cfg)
    pdf_folder = resolve(folder) if folder else pt.invoices
    pdfs = sorted(p for p in pdf_folder.glob("*") if p.suffix.lower() == ".pdf")
    print(f"\n{len(pdfs)} invoice(s) in {pdf_folder}\n")

    results = process(pdfs, cfg)
    for r in results:
        label = f"{r['number']} ({r['file']})" if r["number"] else r["file"]
        if r["status"] == "validated":
            print(f"  ✔  {label}: OK -> logged")
        elif r["status"] == "duplicate":
            print(f"  =  {label}: already logged, skipped")
        elif r["status"] == "unreadable":
            print(f"  ⚠  {label}: cannot be read -> {r['reason']}")
        else:
            print(f"  ✘  {label}: mismatch in {', '.join(m['field'] for m in r['mismatches'])}")
            for m in r["mismatches"]:
                print(f"        {m['field']}: invoice '{m['invoice']}' / master '{m['master']}'")
            print(f"        {r['alert']}")

    n = {s: sum(1 for r in results if r["status"] == s) for s in STATUSES}
    print(f"\nSummary: {n['validated']} validated, {n['mismatch']} with mismatches, "
          f"{n['duplicate']} already logged, {n['unreadable']} unreadable.")
    print(f"Ledger: {pt.ledger}\nExceptions: {pt.exceptions}\n")
    return results


def parse_date(txt: str) -> date:
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(txt.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Invalid date: {txt!r} (use YYYY-MM-DD or DD/MM/YYYY)")


def cmd_report(start: date, end: date, by: str, cfg: dict) -> Path:
    pt = paths_from(cfg)
    path, n = build_report(pt.ledger, pt.reports, start, end, by)
    print(f"\nReport created with {n} invoice(s): {path}\n")
    return path


def menu(cfg: dict) -> None:
    while True:
        print("=== Invoice checks ===")
        print("1. Validate invoices")
        print("2. Report of validated invoices")
        print("3. Exit")
        op = input("Option: ").strip()
        try:
            if op == "1":
                f = input(f"Invoices folder [{cfg['paths']['invoices_folder']}]: ").strip()
                cmd_validate(f or None, cfg)
            elif op == "2":
                today = date.today().isoformat()
                s = input(f"From (YYYY-MM-DD) [{today}]: ").strip() or today
                e = input(f"To (YYYY-MM-DD) [{s}]: ").strip() or s
                t = input("Filter by (V)alidation date or (I)nvoice date [V]: ").strip().upper()
                cmd_report(parse_date(s), parse_date(e), "invoice" if t == "I" else "validation", cfg)
            elif op == "3":
                return
        except Exception as e:
            print(f"\nError: {e}\n")


def main(argv=None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")  # ✔ ✘ symbols in the Windows console
    cfg = load_config()
    p = argparse.ArgumentParser(description="Check sales invoices against the customer master file")
    sub = p.add_subparsers(dest="cmd")
    pv = sub.add_parser("validate", help="Check the PDFs in a folder")
    pv.add_argument("folder", nargs="?")
    pr = sub.add_parser("report", help="Report of validated invoices by date")
    pr.add_argument("--from", dest="start", required=True)
    pr.add_argument("--to", dest="end")
    pr.add_argument("--by", choices=["validation", "invoice"], default="validation")
    a = p.parse_args(argv)

    try:
        if a.cmd == "validate":
            cmd_validate(a.folder, cfg)
        elif a.cmd == "report":
            s = parse_date(a.start)
            cmd_report(s, parse_date(a.end) if a.end else s, a.by, cfg)
        else:
            menu(cfg)
    except (FileLocked, PermissionError, FileNotFoundError, ValueError) as e:
        print(f"\nError: {e}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
