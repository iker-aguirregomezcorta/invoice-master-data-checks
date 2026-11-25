"""
Web interface for the invoice checks (FastAPI).

Start:  python web.py      (or double-click start_web.bat)
The browser opens at http://127.0.0.1:8000

The web app is just a more convenient way to use the tool: it calls the same
functions as the command line (app/service.py), so the matching rule, the
ledger and the alerts are exactly the same.
It only listens on this computer (127.0.0.1): it is not exposed to the network.
"""
from __future__ import annotations

import re
import threading
import webbrowser
from datetime import date, datetime
from pathlib import Path

import uvicorn
from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from app.ledger import LEDGER_COLUMNS
from app.report import build_report, filter_ledger
from app.service import (BASE, FileLocked, cell, load_config, master_summary,
                         paths_from, process, read_table)

HOST, PORT = "127.0.0.1", 8000
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="Invoice checks", docs_url="/api/docs", redoc_url=None)


def current_cfg() -> dict:
    # Re-read on every request: a change in config.json applies without restarting.
    return load_config()


@app.exception_handler(FileLocked)
async def _locked(_, e: FileLocked):
    return JSONResponse({"detail": str(e)}, status_code=409)


@app.get("/", response_class=HTMLResponse)
def home():
    return (BASE / "web" / "index.html").read_text(encoding="utf-8")


@app.get("/api/status")
def status():
    cfg = current_cfg()
    pt = paths_from(cfg)
    try:
        master = master_summary(cfg)
    except Exception as e:
        master = {"error": f"Cannot read the master file ({pt.master.name}): {e}"}
    pending = sorted(p.name for p in pt.invoices.glob("*") if p.suffix.lower() == ".pdf") if pt.invoices.exists() else []
    return {
        "master": master,
        "validated": len(read_table(pt.ledger)),
        "exceptions": len(read_table(pt.exceptions)),
        "invoices_folder": str(pt.invoices),
        "pdfs_in_folder": pending,
        "email_enabled": bool(cfg["email"].get("enabled")),
        "recipients": cfg["email"].get("recipients") or [],
    }


def _safe_name(name: str) -> str:
    name = Path(name or "invoice.pdf").name
    return re.sub(r"[^\w.\- ]", "_", name)[:120] or "invoice.pdf"


@app.post("/api/validate")
async def validate_uploads(files: list[UploadFile] = File(...)):
    """Receive PDFs, keep a copy in output/uploads/<timestamp>/ and check them."""
    cfg = current_cfg()
    pt = paths_from(cfg)
    batch = pt.uploads / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    batch.mkdir(parents=True, exist_ok=True)
    paths = []
    for f in files:
        dest = batch / _safe_name(f.filename)
        if dest.suffix.lower() != ".pdf":
            continue
        dest.write_bytes(await f.read())
        paths.append(dest)
    if not paths:
        raise HTTPException(400, "No PDF received.")
    return {"results": process(paths, cfg)}


@app.post("/api/validate-folder")
def validate_folder():
    """Check the PDFs in the configured input folder (data/invoices)."""
    cfg = current_cfg()
    pt = paths_from(cfg)
    pdfs = sorted(p for p in pt.invoices.glob("*") if p.suffix.lower() == ".pdf")
    if not pdfs:
        raise HTTPException(400, f"No PDFs in {pt.invoices}")
    return {"results": process(pdfs, cfg)}


@app.get("/api/ledger")
def ledger():
    return read_table(paths_from(current_cfg()).ledger)


@app.get("/api/exceptions")
def exceptions():
    return read_table(paths_from(current_cfg()).exceptions)


def _dates(start: str, end: str | None) -> tuple[date, date]:
    try:
        s = date.fromisoformat(start)
        e = date.fromisoformat(end) if end else s
    except ValueError:
        raise HTTPException(400, "Invalid date")
    if s > e:
        raise HTTPException(400, "The start date is after the end date")
    return s, e


@app.get("/api/report")
def report(start: str, end: str | None = None, by: str = Query("validation", pattern="^(validation|invoice)$")):
    s, e = _dates(start, end)
    pt = paths_from(current_cfg())
    if not pt.ledger.exists():
        return {"rows": [], "total": 0, "base": 0, "vat": 0}
    header = [c for c, _ in LEDGER_COLUMNS]
    rows = [dict(zip(header, map(cell, r))) for r in filter_ledger(pt.ledger, s, e, by)]
    total = lambda k: round(sum(r.get(k) or 0 for r in rows), 2)
    return {"rows": rows, "base": total("Taxable base"), "vat": total("VAT amount"), "total": total("Total")}


@app.get("/api/report/download")
def report_download(start: str, end: str | None = None, by: str = Query("validation", pattern="^(validation|invoice)$")):
    s, e = _dates(start, end)
    pt = paths_from(current_cfg())
    try:
        path, _ = build_report(pt.ledger, pt.reports, s, e, by)
    except FileNotFoundError as err:
        raise HTTPException(404, str(err))
    except PermissionError as err:
        raise HTTPException(409, str(err))
    return FileResponse(path, media_type=XLSX, filename=path.name)


@app.get("/api/download/{which}")
def download_excel(which: str):
    pt = paths_from(current_cfg())
    f = {"ledger": pt.ledger, "exceptions": pt.exceptions}.get(which)
    if f is None or not f.exists():
        raise HTTPException(404, "That file does not exist yet")
    return FileResponse(f, media_type=XLSX, filename=f.name)


@app.get("/api/alerts/{name}")
def download_alert(name: str):
    folder = paths_from(current_cfg()).alerts.resolve()
    f = (folder / Path(name).name).resolve()
    if f.parent != folder or f.suffix != ".eml" or not f.exists():
        raise HTTPException(404, "Alert not found")
    return FileResponse(f, media_type="message/rfc822", filename=f.name)


def main():
    url = f"http://{HOST}:{PORT}"
    print(f"\nInvoice checks running at {url}\nClose this window to stop it.\n")
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")


if __name__ == "__main__":
    main()
