"""
Step 1 - Read the PDF invoice and extract the data to be checked.

Invoices are text-based PDFs (not scanned) and always follow the issuer's own
template. They are therefore read with fixed rules (position on the page plus
labels), with no OCR and no AI: the result is exact and repeatable.

What is extracted:
  - Invoice header: number, date, customer VAT ID, customer ID.
  - "Invoice-To" block: company name, street, postcode, town, country.
    This is the billing address, the one that is checked. Only the LEFT column
    is read; the right one is the "Postal address".
  - Amounts (taxable base, VAT amount and total) for the ledger and the report.

Note: the page header is the ISSUER, i.e. our own company. It is not checked:
the customer is.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pdfplumber

BILL_TO_LABEL = "Invoice-To"
POSTAL_ADDRESS_WORD = "Postal"
END_OF_BLOCK_LABEL = "Ship to"
TOTALS_LABEL = "Freight"


class ExtractionError(Exception):
    """The invoice does not have the expected layout and cannot be read reliably."""


@dataclass
class InvoiceData:
    file: str
    number: str
    date: date
    vat_id: str
    customer_id: str
    company_name: str
    street: str
    postcode: str
    town: str
    country: str
    taxable_base: float | None = None
    vat_amount: float | None = None
    total: float | None = None
    reading_warnings: list[str] = field(default_factory=list)


def _find(pattern: str, text: str, name: str) -> str:
    m = re.search(pattern, text)
    if not m:
        raise ExtractionError(f"Field '{name}' not found in the invoice")
    return m.group(1).strip()


def _number(txt: str) -> float:
    """'114,011.15' -> 114011.15 (the template uses comma thousands and dot decimals)."""
    return float(txt.replace(",", ""))


def _lines_by_height(words: list[dict]) -> list[str]:
    """Group words at the same height and return the text lines."""
    lines: dict[int, list[dict]] = {}
    for w in words:
        key = round(w["top"])
        # tolerate 1-2 points of difference between words on the same line
        existing = next((k for k in lines if abs(k - key) <= 2), None)
        lines.setdefault(existing if existing is not None else key, []).append(w)
    return [
        " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"]))
        for _, ws in sorted(lines.items())
    ]


def _bill_to_block(page) -> list[str]:
    """Return the lines of the 'Invoice-To' block (left column)."""
    lines = page.extract_text_lines()

    start = next((l for l in lines if l["text"].startswith(BILL_TO_LABEL)), None)
    if start is None:
        raise ExtractionError("'Invoice-To' block not found")
    end = next(
        (l for l in lines if l["top"] > start["top"] and l["text"].startswith(END_OF_BLOCK_LABEL)),
        None,
    )
    if end is None:
        raise ExtractionError("End of the 'Invoice-To' block not found")

    # The column boundary is where the right-hand label ("Postal address") starts.
    words = page.extract_words()
    right_label = [w for w in words if abs(w["top"] - start["top"]) <= 2 and w["text"] == POSTAL_ADDRESS_WORD]
    limit_x = min(w["x0"] for w in right_label) - 5 if right_label else page.width / 2

    zone = [w for w in words if start["bottom"] < w["top"] < end["top"] and w["x0"] < limit_x]
    return _lines_by_height(zone)


def _split_postcode_town(line: str) -> tuple[str, str]:
    """
    '48013 Bilbao'       -> ('48013', 'Bilbao')
    'LS1 4AP Leeds'      -> ('LS1 4AP', 'Leeds')   (UK postcode with a space)
    """
    parts = line.split()
    if len(parts) >= 3 and any(c.isdigit() for c in parts[1]) and any(c.isdigit() for c in parts[0]):
        return f"{parts[0]} {parts[1]}", " ".join(parts[2:])
    return parts[0], " ".join(parts[1:])


def extract_invoice(path: str | Path) -> InvoiceData:
    path = Path(path)
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            raise ExtractionError("The PDF has no pages")
        first = pdf.pages[0]
        first_text = first.extract_text() or ""
        if not first_text.strip():
            raise ExtractionError("The PDF has no text (scanned?). Manual review needed.")
        last_text = pdf.pages[-1].extract_text() or ""

        number = _find(r"Invoice No:\s*(\S+)", first_text, "Invoice number")
        date_txt = _find(r"Invoice date:\s*(\d{2}\.\d{2}\.\d{4})", first_text, "Invoice date")
        vat_id = _find(r"VAT ID:\s*(\S+)", first_text, "Customer VAT ID")
        customer_id = _find(r"Customer ID:\s*(\S+)", first_text, "Customer ID")

        block = _bill_to_block(first)
        if len(block) != 4:
            raise ExtractionError(
                f"The 'Invoice-To' block has {len(block)} lines and 4 were expected "
                f"(company name, street, postcode + town, country): {block}"
            )
        company_name, street, postcode_town, country = block
        postcode, town = _split_postcode_town(postcode_town)

        data = InvoiceData(
            file=path.name,
            number=number,
            date=datetime.strptime(date_txt, "%d.%m.%Y").date(),
            vat_id=vat_id,
            customer_id=customer_id,
            company_name=company_name,
            street=street,
            postcode=postcode,
            town=town,
            country=country,
        )

        # Amounts: on the last page, on the line after the "Freight / Net amount / ..."
        # header. They are informative (not checked).
        lines = last_text.splitlines()
        idx = next((i for i, l in enumerate(lines) if l.strip().startswith(TOTALS_LABEL)), None)
        if idx is not None:
            numbers = next(
                (re.findall(r"-?[\d,]+\.\d+", l) for l in lines[idx + 1:] if re.findall(r"[\d,]+\.\d{2}", l)),
                [],
            )
            # [freight, net, base, (VAT rate %), VAT amount, total, exchange rate]
            if len(numbers) >= 5:
                data.taxable_base = _number(numbers[2])
                data.vat_amount = _number(numbers[-3])
                data.total = _number(numbers[-2])
        if data.total is None:
            data.reading_warnings.append("Amounts could not be read")

    return data
