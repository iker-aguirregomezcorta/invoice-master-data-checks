"""
Step 4 - Email alert when an invoice does not match the master file.

One email per invoice, with the invoice number and a table of the mismatching
fields (value on the invoice vs. value in the master file).

Two modes (config.json -> "email"):
  - "enabled": true  -> sent via SMTP to the configured recipients.
  - "enabled": false -> the email is saved as an .eml file in output/alerts/.
    It opens in Outlook or any email client. Useful for demos and for
    testing without a mail server.
The SMTP password is never stored in a file: it is read from an environment variable.
"""
from __future__ import annotations

import html
import os
import re
import smtplib
from email.message import EmailMessage
from pathlib import Path

from .validation import ValidationResult


def build_email(r: ValidationResult, cfg: dict) -> EmailMessage:
    f = r.invoice
    msg = EmailMessage()
    msg["Subject"] = f"[Invoice checks] Mismatch in invoice {f.number}"
    msg["From"] = cfg.get("sender") or "invoice-checks@localhost"
    msg["To"] = ", ".join(cfg.get("recipients") or ["recipient@example.com"])

    lines = [
        f"Invoice {f.number} ({f.file}) does not match the customer master file.",
        f"Customer: {f.company_name} - Customer ID {f.customer_id} - Invoice date {f.date:%d/%m/%Y}",
        "",
        "Mismatching fields:",
    ]
    for m in r.mismatches:
        lines.append(f"  - {m.field}: on the invoice '{m.invoice_value}' / in the master file '{m.master_value}'")
    lines += ["", "The invoice has NOT been logged. Fix the invoice or the master file and run the check again."]
    msg.set_content("\n".join(lines))

    rows = "".join(
        f"<tr><td><b>{html.escape(m.field)}</b></td><td>{html.escape(m.invoice_value)}</td>"
        f"<td>{html.escape(m.master_value)}</td></tr>"
        for m in r.mismatches
    )
    msg.add_alternative(f"""\
<p>Invoice <b>{html.escape(f.number)}</b> ({html.escape(f.file)}) does not match the customer master file.</p>
<p>Customer: {html.escape(f.company_name)} &middot; Customer ID {html.escape(f.customer_id)} &middot; Invoice date {f.date:%d/%m/%Y}</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial;font-size:13px">
<tr style="background:#1F3864;color:#fff"><th>Field</th><th>On the invoice</th><th>In the master file</th></tr>
{rows}
</table>
<p>The invoice has <b>not</b> been logged. Fix the invoice or the master file and run the check again.</p>
""", subtype="html")
    return msg


def send_alert(r: ValidationResult, cfg: dict, alerts_folder: Path) -> tuple[str, Path | None]:
    """
    Send (or save) the alert. Returns a text describing what was done and,
    if it was saved as .eml, the file path.
    """
    msg = build_email(r, cfg)

    if cfg.get("enabled"):
        password = os.environ.get(cfg.get("password_env_var", "SMTP_PASSWORD"), "")
        try:
            with smtplib.SMTP(cfg["smtp_server"], int(cfg.get("smtp_port", 587)), timeout=30) as s:
                if cfg.get("use_tls", True):
                    s.starttls()
                if cfg.get("username"):
                    s.login(cfg["username"], password)
                s.send_message(msg)
            return f"Email sent to {msg['To']}", None
        except Exception as e:  # if sending fails, the alert is not lost
            reason = f"SMTP error ({e.__class__.__name__}: {e}); saved as .eml"
    else:
        reason = "Sending disabled; saved as .eml"

    alerts_folder.mkdir(parents=True, exist_ok=True)
    name = re.sub(r"[^\w-]", "_", r.invoice.number)
    path = alerts_folder / f"alert_{name}.eml"
    path.write_bytes(bytes(msg))
    return f"{reason}: {path.name}", path
