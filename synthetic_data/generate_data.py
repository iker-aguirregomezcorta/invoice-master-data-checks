"""
Generate the demo data: a customer master file (Excel) and a batch of PDF sales
invoices, with errors planted in a controlled way.

Everything is made up: the issuing company, the customers, the tax IDs, the
addresses and the amounts. Spanish tax IDs have a valid check digit so they look
real, but they are generated at random.

Besides the files, it writes data/expected.csv: for each PDF, which error was
planted and which outcome the control should give. The automated tests compare
the actual outcome with this file.

Usage (from the project folder):
    python synthetic_data/generate_data.py

The same seed always produces the same data.
"""
from __future__ import annotations

import csv
import random
import shutil
from dataclasses import dataclass, replace
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

SEED = 2026
BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
INVOICES_FOLDER = DATA / "invoices"

ISSUER = {
    "name": "ALDAKO TECHNICAL SUPPLIES, S.L.",
    "address": "Calle de la Demostración, 1 - 48001 Bilbao (Spain)",
    "vat_id": "ESB00000000",
    "email": "invoicing@example.com",
}

# ------------------------------------------------------------------ customers

PREFIXES = ["Talleres", "Mecanizados", "Suministros", "Distribuciones", "Montajes",
            "Electro", "Logística", "Ferretería", "Calderería", "Automatismos"]
SYLLABLES_START = ["Ar", "Be", "Zu", "Ola", "Itu", "Mar", "Gal", "Er", "Lan", "Uri", "Ibar", "Amez"]
SYLLABLES_END = ["luce", "beti", "goi", "zabal", "torre", "mendi", "aga", "eta", "berri", "arte", "ondo", "lde"]

SPANISH_TOWNS = [
    ("Bilbao", "480"), ("Barakaldo", "489"), ("Getxo", "489"), ("Vitoria-Gasteiz", "010"),
    ("Donostia-San Sebastián", "200"), ("Pamplona", "310"), ("Logroño", "260"),
    ("Burgos", "090"), ("Santander", "390"), ("Zaragoza", "500"), ("Madrid", "280"),
    ("Valencia", "460"), ("Barcelona", "080"), ("Valladolid", "470"),
]
SPANISH_STREETS = ["Calle Mayor", "Avenida de la Libertad", "Calle Nueva", "Calle San Juan",
                   "Avenida del Puerto", "Calle de la Industria", "Calle Real", "Paseo del Río",
                   "Calle del Comercio", "Avenida de Europa", "Calle Gran Vía", "Calle del Carmen"]

FOREIGN = [
    # country, town, postcode, street, legal form, VAT prefix, number of digits
    ("PRT", "Porto", "4100-001", "Rua da Indústria", "Lda.", "PT", 9),
    ("FRA", "Bordeaux", "33000", "Rue des Ateliers", "SAS", "FR", 11),
    ("DEU", "Stuttgart", "70173", "Industriestrasse", "GmbH", "DE", 9),
    ("ITA", "Torino", "10121", "Via dell'Industria", "S.r.l.", "IT", 11),
    ("NLD", "Eindhoven", "5611", "Fabrieksstraat", "B.V.", "NL", 12),
    ("POL", "Poznan", "60-101", "Ulica Fabryczna", "Sp. z o.o.", "PL", 10),
    ("GBR", "Leeds", "LS1 4AP", "Mill Lane", "Ltd", "GB", 9),
]


def valid_cif(rng: random.Random) -> str:
    """Spanish company tax ID (B + 7 digits + check digit) with a correct check digit."""
    digits = [rng.randint(0, 9) for _ in range(7)]
    even = sum(digits[1::2])
    odd = sum(sum(divmod(2 * d, 10)) for d in digits[0::2])
    check = (10 - (even + odd) % 10) % 10
    return "B" + "".join(map(str, digits)) + str(check)


@dataclass
class Customer:
    company_name: str
    street: str
    postcode: str
    town: str
    country: str
    vat_id: str | None
    phone: str | None
    customer_id: str


def made_up_name(rng: random.Random) -> str:
    return rng.choice(SYLLABLES_START) + rng.choice(SYLLABLES_END)


def generate_customers(rng: random.Random) -> list[Customer]:
    customers: list[Customer] = []
    used: set[str] = set()
    next_id = 100100

    def new_name(prefix: str, suffix: str) -> str:
        while True:
            n = f"{prefix} {made_up_name(rng)}{suffix}"
            if n not in used:
                used.add(n)
                return n

    for _ in range(64):
        town, pc3 = rng.choice(SPANISH_TOWNS)
        next_id += rng.randint(1, 9)
        customers.append(Customer(
            company_name=new_name(rng.choice(PREFIXES), ", S.L."),
            street=f"{rng.choice(SPANISH_STREETS)}, {rng.randint(1, 140)}",
            postcode=f"{pc3}{rng.randint(1, 99):02d}",
            town=town,
            country="ESP",
            vat_id=valid_cif(rng),
            phone=f"+34 9{rng.randint(10, 99)} {rng.randint(100, 999)} {rng.randint(100, 999)}" if rng.random() < 0.7 else None,
            customer_id=str(next_id),
        ))

    for country, town, pc, street, form, prefix, nd in FOREIGN:
        for _ in range(2):
            next_id += rng.randint(1, 9)
            customers.append(Customer(
                company_name=new_name(made_up_name(rng).capitalize(), f" Industrial {form}"),
                street=f"{street}, {rng.randint(1, 140)}",
                postcode=pc,
                town=town,
                country=country,
                vat_id=prefix + "".join(str(rng.randint(0, 9)) for _ in range(nd)),
                phone=None,
                customer_id=str(next_id),
            ))

    # Near-identical names that really exist in the master file: this is why
    # "looks similar" cannot be accepted (they are two different companies).
    base = customers[3]
    word = base.company_name.split()[1].rstrip(",")
    similar = word[:-1] + ("z" if word[-1] != "z" else "s")
    twin = replace(base, company_name=base.company_name.replace(word, similar, 1),
                   vat_id=valid_cif(rng), street=f"{rng.choice(SPANISH_STREETS)}, {rng.randint(1, 140)}",
                   customer_id=str(next_id + 7))
    customers.append(twin)

    # Two customers whose VAT ID is still missing in the master file.
    for i in (10, 40):
        customers[i] = replace(customers[i], vat_id=None)

    rng.shuffle(customers)
    return customers


# ------------------------------------------------------------------ invoices

PRODUCTS = [
    ("Deep groove ball bearing 6205", 4.8, 9.5), ("V-belt SPA 1250", 11.0, 19.0),
    ("Pneumatic cylinder ISO 15552", 85.0, 160.0), ("Inductive sensor M18", 22.0, 41.0),
    ("3-pole contactor 25 A", 31.0, 58.0), ("Frequency inverter 2.2 kW", 240.0, 410.0),
    ("Hex key set", 9.0, 18.0), ("Cut-resistant gloves level D", 2.1, 4.3),
    ("Synthetic lubricant 5 L", 38.0, 64.0), ("Solenoid valve 5/2 1/4\"", 46.0, 79.0),
    ("Flexible coupling", 17.0, 35.0), ("Emergency stop button", 12.0, 24.0),
]
PAYMENT = [("BANK TRANSFER", "30 DAYS FROM INVOICE DATE", 30), ("CONFIRMING", "60 DAYS FROM INVOICE DATE", 60),
           ("DIRECT DEBIT", "CASH", 0)]


@dataclass
class Invoice:
    file: str
    number: str
    date: date
    customer: Customer        # what is printed on the invoice (may contain errors)
    lines: list[tuple]
    error: str                # description of the planted error
    expected: str             # expected status
    fields: str               # expected mismatching fields, separated by "|"


def generate_lines(rng: random.Random, n: int) -> list[tuple]:
    lines = []
    for _ in range(n):
        name, pmin, pmax = rng.choice(PRODUCTS)
        qty = rng.choice([1, 2, 4, 5, 10, 12, 20, 25, 50, 100])
        price = round(rng.uniform(pmin, pmax), 2)
        lines.append((name, qty, price, round(qty * price, 2)))
    return lines


def change_digit(txt: str, rng: random.Random) -> str:
    pos = [i for i, c in enumerate(txt) if c.isdigit()]
    i = rng.choice(pos[-4:])
    new = str((int(txt[i]) + rng.randint(1, 8)) % 10)
    return txt[:i] + new + txt[i + 1:]


def swap_number(street: str) -> str:
    """'Calle Mayor, 18' -> 'Calle Mayor, 81'."""
    name, num = street.rsplit(", ", 1)
    return f"{name}, {num[::-1]}"


def different_format(c: Customer) -> Customer:
    """Same information, written differently: must be accepted."""
    street = c.street.replace("Calle ", "C/ ").replace("Avenida ", "Avda. ").replace(",", "")
    return replace(c, company_name=c.company_name.upper().replace(", S.L.", " SL"),
                   street=street, town=c.town.upper(),
                   vat_id=f"ES{c.vat_id}" if c.country == "ESP" else c.vat_id)


SCENARIOS = (
    [("ok", "No errors", "validated", "")] * 9
    + [("format", "Formatting only (upper case, C/, SL, ES prefix on the VAT ID)", "validated", "")] * 4
    + [("vat", "One wrong digit in the VAT ID", "mismatch", "VAT ID")] * 2
    + [("street", "Street number with swapped digits", "mismatch", "Street")] * 2
    + [("postcode", "One wrong digit in the postcode", "mismatch", "Postcode")] * 2
    + [("name", "One wrong letter in the company name", "mismatch", "Company name")]
    + [("town", "Town differs from the master file", "mismatch", "Town")]
    + [("several", "Outdated postcode and town", "mismatch", "Postcode|Town")]
    + [("unknown", "Customer not in the master file", "mismatch", "Customer")]
    + [("no_vat", "Customer has no VAT ID in the master file", "mismatch", "VAT ID")]
    + [("long", "3-page invoice, no errors", "validated", "")]
)


def generate_invoices(rng: random.Random, customers: list[Customer]) -> list[Invoice]:
    with_vat = [c for c in customers if c.vat_id]
    without_vat = [c for c in customers if not c.vat_id]
    numeric_pc = [c for c in with_vat if c.postcode.isdigit()]
    spanish = [c for c in with_vat if c.country == "ESP"]
    invoices = []
    number = 2600417
    day = date(2026, 10, 1)
    scenarios = list(SCENARIOS)
    rng.shuffle(scenarios)

    for key, error, expected, fields in scenarios:
        number += rng.randint(1, 6)
        day += timedelta(days=rng.choice([0, 1, 1, 2, 3]))
        if key == "postcode":
            c = rng.choice(numeric_pc)
        elif key == "street":   # number with 2+ digits that changes when reversed
            c = rng.choice([x for x in with_vat if (n := x.street.rsplit(", ", 1)[1]) != n[::-1]])
        elif key in ("format", "town", "several"):
            c = rng.choice(spanish)
        elif key == "no_vat":
            c = rng.choice(without_vat)
        else:
            c = rng.choice(with_vat)

        printed = c
        if key == "format":
            printed = different_format(c)
        elif key == "vat":
            printed = replace(c, vat_id=change_digit(c.vat_id, rng))
        elif key == "street":
            printed = replace(c, street=swap_number(c.street))
        elif key == "postcode":
            printed = replace(c, postcode=change_digit(c.postcode, rng))
        elif key == "name":
            p = c.company_name.split()
            p[1] = p[1][:-2] + ("o" if p[1][-2] != "o" else "a") + p[1][-1]
            printed = replace(c, company_name=" ".join(p))
        elif key == "town":
            other = rng.choice([t for t, _ in SPANISH_TOWNS if t != c.town])
            printed = replace(c, town=other)
        elif key == "several":
            other, pc3 = rng.choice([x for x in SPANISH_TOWNS if x[0] != c.town])
            printed = replace(c, postcode=f"{pc3}{rng.randint(1, 99):02d}", town=other)
        elif key == "unknown":
            town, pc3 = rng.choice(SPANISH_TOWNS)
            printed = Customer(f"Montajes {made_up_name(rng)}tegi, S.L.", f"Calle Nueva, {rng.randint(1, 99)}",
                               f"{pc3}{rng.randint(1, 99):02d}", town, "ESP", valid_cif(rng), None, "100999")
        elif key == "no_vat":
            printed = replace(c, vat_id=valid_cif(rng))

        n_lines = 38 if key == "long" else rng.randint(1, 7)
        invoices.append(Invoice(
            file="", number=f"INV-{number}", date=day, customer=printed,
            lines=generate_lines(rng, n_lines), error=error, expected=expected, fields=fields,
        ))

    invoices.sort(key=lambda f: f.number)
    for i, f in enumerate(invoices, start=1):
        f.file = f"Invoice_{i:02d}.pdf"
    return invoices


# ------------------------------------------------------------------ PDF

def money(x: float) -> str:
    return f"{x:,.2f}"


def draw_invoice(f: Invoice, path: Path) -> None:
    domestic = f.customer.country == "ESP"
    c = canvas.Canvas(str(path), pagesize=A4, invariant=1)
    c.setTitle(f"Invoice {f.number}")
    c.setAuthor("Synthetic data")
    c.setCreator("generate_data.py")
    width, height = A4
    cust = f.customer
    per_page = 16
    chunks = [f.lines[i:i + per_page] for i in range(0, len(f.lines), per_page)]
    base = round(sum(l[3] for l in f.lines), 2)
    rate = 21.0 if domestic else 0.0
    vat = round(base * rate / 100, 2)
    total = round(base + vat, 2)
    method, terms, days = PAYMENT[int(f.number[-2:]) % len(PAYMENT)]

    for page_no, chunk in enumerate(chunks, start=1):
        # Issuer header
        c.setFillColorRGB(0.12, 0.22, 0.39)
        c.rect(0, height - 95, width, 95, fill=1, stroke=0)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(40, height - 40, ISSUER["name"])
        c.setFont("Helvetica", 8.5)
        c.drawString(40, height - 55, ISSUER["address"])
        c.drawString(40, height - 67, f"Tax ID: {ISSUER['vat_id']} | {ISSUER['email']}")
        c.drawString(40, height - 79, "Synthetic data generated for demonstration")
        c.setFont("Helvetica-Bold", 18)
        c.drawRightString(width - 40, height - 45, "INVOICE")
        c.setFont("Helvetica", 8.5)
        c.drawRightString(width - 40, height - 60, f"Page {page_no} / {len(chunks)}")
        c.setFillColorRGB(0, 0, 0)

        y = height - 130
        if page_no == 1:
            c.setFont("Helvetica-Bold", 9)
            c.drawString(40, y, "Invoice-To")
            c.drawString(300, y, "Postal address")
            c.setFont("Helvetica", 9)
            block = [cust.company_name, cust.street, f"{cust.postcode} {cust.town}", cust.country]
            for i, txt in enumerate(block, start=1):
                c.drawString(40, y - 14 * i, txt)
                c.drawString(300, y - 14 * i, txt)
            y -= 90

        c.setFont("Helvetica-Bold", 9)
        c.drawString(40, y, "Ship to")
        c.setFont("Helvetica", 9)
        details = [f"Invoice No: {f.number}", f"Invoice date: {f.date:%d.%m.%Y}",
                   f"VAT ID: {cust.vat_id}", f"Customer ID: {cust.customer_id}"]
        ship = [cust.company_name, cust.street, f"{cust.postcode} {cust.town}", cust.country]
        for i in range(4):
            c.drawString(40, y - 14 * (i + 1), ship[i])
            c.drawString(330, y - 14 * (i + 1), details[i])

        # Line items
        y -= 90
        c.setFont("Helvetica-Bold", 8.5)
        for x, t in zip((40, 100, 360, 430, 510), ["Ref.", "Description", "Quantity", "Unit price", "Amount"]):
            c.drawString(x, y, t)
        c.line(40, y - 4, width - 40, y - 4)
        c.setFont("Helvetica", 8.5)
        for j, (name, qty, price, amount) in enumerate(chunk):
            y -= 16
            ref = 700100 + (page_no - 1) * per_page + j + 1
            c.drawString(40, y, str(ref))
            c.drawString(100, y, name)
            c.drawRightString(400, y, str(qty))
            c.drawRightString(480, y, money(price))
            c.drawRightString(width - 40, y, money(amount))

        if page_no < len(chunks):
            carried = sum(l[3] for ch in chunks[:page_no] for l in ch)
            y -= 22
            c.drawString(100, y, "Carried forward: " + money(carried).replace(",", " "))
            c.showPage()
            continue

        # Totals (last page only)
        y -= 40
        c.line(40, y + 14, width - 40, y + 14)
        c.setFont("Helvetica-Bold", 8.5)
        headers = ["Freight", "Net amount", "Taxable base", "VAT rate", "VAT", "Total (EUR)", "Rate"]
        xs = (40, 105, 185, 270, 330, 400, 480)
        for x, t in zip(xs, headers):
            c.drawString(x, y, t)
        y -= 14
        c.setFont("Helvetica", 9)
        values = ["0.00", money(base), money(base), f"{rate:.2f} %" if domestic else "", money(vat), money(total), "1.000000"]
        for x, t in zip(xs, values):
            if t:
                c.drawString(x, y, t)
        y -= 30
        due = f.date + timedelta(days=days)
        c.setFont("Helvetica", 8.5)
        c.drawString(40, y, "Payment: " + method)
        c.drawString(40, y - 12, "Terms: " + terms)
        c.drawString(40, y - 24, f"Due date: {due:%d.%m.%Y}")
        if not domestic:
            c.drawString(40, y - 36, "Export outside the EU: VAT exempt" if cust.country == "GBR"
                         else "Intra-community supply: VAT exempt")
        c.setFont("Helvetica-Oblique", 7.5)
        c.drawString(40, 30, "Demo document. Companies, addresses and tax IDs are made up.")
        c.showPage()
    c.save()


def draw_scanned(path: Path) -> None:
    """A PDF with no text (like a scan): the control must never pass it as valid."""
    c = canvas.Canvas(str(path), pagesize=A4, invariant=1)
    c.setFillColorRGB(0.85, 0.85, 0.85)
    for i in range(40):
        c.rect(40, 780 - i * 18, 300 + (i * 37) % 200, 8, fill=1, stroke=0)
    c.showPage()
    c.save()


# ------------------------------------------------------------------ Excel

def write_master(customers: list[Customer], path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Customers"
    ws.append(["Company name", "Street", "Postcode", "Town", "Country", "Phone", "VAT ID"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F3864")
    for c in customers:
        ws.append([c.company_name, c.street, c.postcode, c.town, c.country, c.phone, c.vat_id])
    for col, w in zip("ABCDEFG", (42, 34, 11, 24, 9, 17, 16)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = "@"   # postcode as text: keep leading zeros
    ws.freeze_panes = "A2"
    wb.properties.creator = "generate_data.py"
    wb.save(path)


def main() -> None:
    rng = random.Random(SEED)
    if INVOICES_FOLDER.exists():
        shutil.rmtree(INVOICES_FOLDER)
    INVOICES_FOLDER.mkdir(parents=True)

    customers = generate_customers(rng)
    write_master(customers, DATA / "customer_master.xlsx")

    invoices = generate_invoices(rng, customers)
    for f in invoices:
        draw_invoice(f, INVOICES_FOLDER / f.file)

    rows = [(f.file, f.number, f.error, f.expected, f.fields) for f in invoices]

    # Process cases: an invoice submitted twice and a PDF with no text.
    original = next(f for f in invoices if f.expected == "validated")
    n = len(invoices)
    copy_name = f"Invoice_{n + 1:02d}.pdf"
    shutil.copy(INVOICES_FOLDER / original.file, INVOICES_FOLDER / copy_name)
    rows.append((copy_name, original.number, f"Copy of {original.file} (same invoice submitted twice)",
                 "duplicate", ""))
    scanned = f"Invoice_{n + 2:02d}.pdf"
    draw_scanned(INVOICES_FOLDER / scanned)
    rows.append((scanned, "", "Scanned PDF with no text", "unreadable", ""))

    with open(DATA / "expected.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["file", "number", "planted_error", "expected_status", "expected_fields"])
        w.writerows(rows)

    print(f"{len(customers)} customers in the master file, {len(rows)} PDFs in {INVOICES_FOLDER}")


if __name__ == "__main__":
    main()
