"""
Step 2 - Compare the invoice data with the customer master file (Excel).

Matching rule (the key design decision):
  - FORMAT differences ARE accepted: upper/lower case, accents, spacing,
    punctuation and equivalent abbreviations
    ("S.L." = "SL", "C/" = "Calle", "Avda." = "Avenida", "Ltd" = "Limited").
  - NO CONTENT difference is accepted: one different letter or digit in the
    company name, VAT ID, street/number, postcode, town or country is a
    mismatch. "Itutorre" != "Itutorro", "48" != "84", "39035" != "39635".

That is why there is no fuzzy matching and no AI model deciding whether two
values "look alike": the very errors to be caught (one changed digit, two
swapped digits) are the ones a similarity-based comparison would let through.

How the customer is found in the master file:
  1st by normalised VAT ID; 2nd, if not found, by normalised company name.
  So if the VAT ID on the invoice is wrong, the customer is still found by
  name and the alert says "VAT ID differs" instead of "customer not found".
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .extraction import InvoiceData

# --------------------------------------------------------------------------
# Normalisation: removes differences of form only, never of content.
# --------------------------------------------------------------------------

# Equivalent abbreviations (applied to text already upper-cased, without dots).
EQUIVALENCES = [
    (r"\bAVDA\b", "AVENIDA"),
    (r"\bAV\b", "AVENIDA"),
    (r"\bPº\b", "PASEO"),
    (r"\bPLZA?\b", "PLAZA"),
    (r"\bSOCIEDAD LIMITADA\b", "SL"),
    (r"\bSOCIEDAD ANONIMA\b", "SA"),
    (r"\bS\s+L\b", "SL"),               # "S. L." -> "S L" -> "SL"
    (r"\bS\s+A\b", "SA"),
    (r"\bST\b", "STREET"),
    (r"\bRD\b", "ROAD"),
    (r"\bAVE\b", "AVENUE"),
    (r"\bLIMITED\b", "LTD"),
]


def strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def normalise_text(value) -> str:
    """'C/ Alcalá, 22' and 'Calle Alcala 22' -> 'CALLE ALCALA 22'."""
    if value is None:
        return ""
    t = strip_accents(str(value)).upper()
    t = re.sub(r"\bC\s*/\s*", "CALLE ", t)   # "C/ Alcala" -> "CALLE Alcala"
    t = re.sub(r"[.,;:]", " ", t)            # drop punctuation ("S.L." -> "S L")
    t = re.sub(r"\s+", " ", t).strip()
    for pattern, replacement in EQUIVALENCES:
        t = re.sub(pattern, replacement, t)
    return re.sub(r"\s+", " ", t).strip()


def normalise_code(value) -> str:
    """VAT ID and postcode: letters and digits only, upper case. 'B-78.310.018' -> 'B78310018'."""
    if value is None:
        return ""
    return re.sub(r"[^0-9A-Z]", "", strip_accents(str(value)).upper())


def normalise_vat_id(value, country: str | None = None) -> str:
    """Also, the 'ES' prefix of an EU VAT number is equivalent to the Spanish tax ID."""
    n = normalise_code(value)
    if (country or "").upper() in ("ESP", "ES") and n.startswith("ES") and len(n) == 11:
        n = n[2:]
    return n


# --------------------------------------------------------------------------
# Customer master file
# --------------------------------------------------------------------------

MASTER_COLUMNS = {
    "company_name": "Company name",
    "street": "Street",
    "postcode": "Postcode",
    "town": "Town",
    "country": "Country",
    "vat_id": "VAT ID",
}


@dataclass
class Customer:
    row: int
    company_name: str
    street: str
    postcode: str
    town: str
    country: str
    vat_id: str | None


def load_master(path: str | Path) -> list[Customer]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]
    missing = [c for c in MASTER_COLUMNS.values() if c not in header]
    if missing:
        raise ValueError(f"The customer master file is missing these columns: {missing}")
    idx = {k: header.index(v) for k, v in MASTER_COLUMNS.items()}

    customers = []
    for n, row in enumerate(rows, start=2):
        if not row or row[idx["company_name"]] in (None, ""):
            continue
        # Postcodes are read as text so leading zeros are kept ("01013").
        customers.append(Customer(
            row=n,
            **{k: (str(row[i]).strip() if row[i] is not None else None) for k, i in idx.items()},
        ))
    wb.close()
    return customers


# --------------------------------------------------------------------------
# Validation
# --------------------------------------------------------------------------

@dataclass
class Mismatch:
    field: str
    invoice_value: str
    master_value: str


@dataclass
class ValidationResult:
    invoice: InvoiceData
    customer: Customer | None
    mismatches: list[Mismatch]
    matched_by: str  # "VAT ID", "company name" or "" if not found

    @property
    def is_valid(self) -> bool:
        return self.customer is not None and not self.mismatches


# (display name, attribute, normalisation function)
FIELDS_TO_CHECK = [
    ("Company name", "company_name", normalise_text),
    ("VAT ID", "vat_id", None),  # uses normalise_vat_id with the country
    ("Street", "street", normalise_text),
    ("Postcode", "postcode", normalise_code),
    ("Town", "town", normalise_text),
    ("Country", "country", normalise_text),
]


def find_customer(invoice: InvoiceData, master: list[Customer]) -> tuple[Customer | None, str]:
    vat = normalise_vat_id(invoice.vat_id, invoice.country)
    by_vat = [c for c in master if c.vat_id and normalise_vat_id(c.vat_id, c.country) == vat]
    if len(by_vat) == 1:
        return by_vat[0], "VAT ID"

    name = normalise_text(invoice.company_name)
    by_name = [c for c in master if normalise_text(c.company_name) == name]
    if len(by_name) == 1:
        return by_name[0], "company name"
    return None, ""


def validate(invoice: InvoiceData, master: list[Customer]) -> ValidationResult:
    customer, how = find_customer(invoice, master)
    if customer is None:
        return ValidationResult(
            invoice, None,
            [Mismatch("Customer", f"{invoice.company_name} / {invoice.vat_id}",
                      "Not in the master file (neither by VAT ID nor by name)")],
            "",
        )

    mismatches = []
    for name, attr, norm in FIELDS_TO_CHECK:
        v_inv = getattr(invoice, attr)
        v_mas = getattr(customer, attr)
        if attr == "vat_id":
            equal = bool(v_mas) and normalise_vat_id(v_inv, invoice.country) == normalise_vat_id(v_mas, customer.country)
        else:
            equal = norm(v_inv) == norm(v_mas)
        if not equal:
            mismatches.append(Mismatch(name, v_inv or "(empty)", v_mas or "(empty in the master file)"))

    return ValidationResult(invoice, customer, mismatches, how)
