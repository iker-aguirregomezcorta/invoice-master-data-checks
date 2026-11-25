"""
Automated tests. Run from the project folder:
    python -m unittest discover tests -v
"""
import csv
import shutil
import tempfile
import unittest
from dataclasses import replace
from datetime import date, datetime
from pathlib import Path

import openpyxl

from app.extraction import ExtractionError, extract_invoice
from app.report import build_report
from app.validation import load_master, normalise_text, normalise_vat_id, validate

BASE = Path(__file__).resolve().parent.parent
INVOICES = BASE / "data" / "invoices"
MASTER = load_master(BASE / "data" / "customer_master.xlsx")
with open(BASE / "data" / "expected.csv", encoding="utf-8") as fh:
    EXPECTED = list(csv.DictReader(fh))


def first(error: str) -> str:
    return next(e["file"] for e in EXPECTED if e["planted_error"] == error)


class TestNormalisation(unittest.TestCase):
    """FORMAT differences -> accepted."""

    def test_legal_form(self):
        self.assertEqual(normalise_text("Talleres Itutorre, S.L."), normalise_text("TALLERES ITUTORRE SL"))
        self.assertEqual(normalise_text("Itutorre S. L."), normalise_text("Itutorre SL"))
        self.assertEqual(normalise_text("Mill Industrial Limited"), normalise_text("Mill Industrial Ltd."))

    def test_accents_case_spacing(self):
        self.assertEqual(normalise_text("Logística  Zubeti"), normalise_text("logistica zubeti"))
        self.assertEqual(normalise_text("Logroño"), normalise_text("LOGRONO"))

    def test_street_abbreviations(self):
        self.assertEqual(normalise_text("C/ Mayor, 22"), normalise_text("Calle Mayor 22"))
        self.assertEqual(normalise_text("Avda. de Europa, 18"), normalise_text("Avenida de Europa, 18"))
        self.assertEqual(normalise_text("Mill St, 4"), normalise_text("Mill Street 4"))

    def test_vat_id_format(self):
        self.assertEqual(normalise_vat_id("b-53.640.560"), normalise_vat_id("B53640560"))
        self.assertEqual(normalise_vat_id("ESB53640560", "ESP"), normalise_vat_id("B53640560", "ESP"))


class TestDifferentContent(unittest.TestCase):
    """CONTENT differences -> mismatches."""

    def test_similar_names_are_not_equal(self):
        self.assertNotEqual(normalise_text("Talleres Itutorre, S.L."), normalise_text("Talleres Itutorro, S.L."))

    def test_swapped_street_number(self):
        self.assertNotEqual(normalise_text("Calle del Comercio, 48"), normalise_text("Calle del Comercio, 84"))

    def test_vat_id_one_digit(self):
        self.assertNotEqual(normalise_vat_id("IT76113878921"), normalise_vat_id("IT76113875921"))


class TestPlantedErrors(unittest.TestCase):
    """
    Each demo invoice carries a known error (data/expected.csv).
    The control must find exactly that error: no more, no less.
    """

    def test_every_invoice_gives_the_expected_outcome(self):
        seen = set()
        for e in EXPECTED:
            with self.subTest(file=e["file"], error=e["planted_error"]):
                if e["expected_status"] == "unreadable":
                    with self.assertRaises(ExtractionError):
                        extract_invoice(INVOICES / e["file"])
                    continue
                data = extract_invoice(INVOICES / e["file"])
                self.assertEqual(data.number, e["number"])
                if e["expected_status"] == "duplicate":
                    self.assertIn(data.number, seen)
                    continue
                seen.add(data.number)
                r = validate(data, MASTER)
                self.assertEqual(r.is_valid, e["expected_status"] == "validated")
                self.assertEqual("|".join(m.field for m in r.mismatches), e["expected_fields"])

    def test_wrong_vat_id_still_finds_the_customer(self):
        r = validate(extract_invoice(INVOICES / first("One wrong digit in the VAT ID")), MASTER)
        self.assertEqual(r.matched_by, "company name")

    def test_issuer_is_not_read_as_customer(self):
        f = extract_invoice(INVOICES / first("No errors"))
        self.assertNotIn("ALDAKO", f.company_name.upper())
        self.assertNotEqual(f.vat_id, "ESB00000000")

    def test_multi_page_invoice(self):
        f = extract_invoice(INVOICES / first("3-page invoice, no errors"))
        self.assertIsNotNone(f.total)
        self.assertAlmostEqual(f.total, round(f.taxable_base + f.vat_amount, 2), places=2)


class TestEdgeCases(unittest.TestCase):
    def setUp(self):
        self.base = extract_invoice(INVOICES / first("No errors"))

    def test_different_format_is_still_valid(self):
        f = replace(self.base, company_name=self.base.company_name.upper().replace(", S.L.", " SL"),
                    street=self.base.street.replace("Calle ", "C/ ").replace(",", ""),
                    vat_id=f"ES{self.base.vat_id}" if self.base.country == "ESP" else self.base.vat_id)
        self.assertTrue(validate(f, MASTER).is_valid)

    def test_similar_name_is_a_mismatch(self):
        f = replace(self.base, company_name=self.base.company_name.replace(",", "x,", 1))
        self.assertEqual([m.field for m in validate(f, MASTER).mismatches], ["Company name"])

    def test_several_mismatches_at_once(self):
        f = replace(self.base, postcode="99999", town="Another Town")
        self.assertEqual([m.field for m in validate(f, MASTER).mismatches], ["Postcode", "Town"])

    def test_unknown_customer(self):
        f = replace(self.base, company_name="Made Up Company, S.L.", vat_id="B00000001")
        r = validate(f, MASTER)
        self.assertIsNone(r.customer)
        self.assertEqual(r.mismatches[0].field, "Customer")

    def test_customer_without_vat_id_in_master(self):
        # A customer with no VAT ID in the master file can never be validated.
        c = next(c for c in MASTER if not c.vat_id)
        f = replace(self.base, company_name=c.company_name, street=c.street, postcode=c.postcode,
                    town=c.town, country=c.country, vat_id="B12345678")
        self.assertEqual([m.field for m in validate(f, MASTER).mismatches], ["VAT ID"])


class TestReport(unittest.TestCase):
    def test_date_filter(self):
        tmp = Path(tempfile.mkdtemp())
        try:
            ledger = tmp / "ledger.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Invoice no."] + [None] * 13)
            ws.append(["A", date(2026, 10, 1)] + [None] * 7 + [100, 21, 121, "a.pdf", datetime(2026, 10, 19, 9)])
            ws.append(["B", date(2026, 10, 20)] + [None] * 7 + [200, 42, 242, "b.pdf", datetime(2026, 10, 20, 9)])
            wb.save(ledger)
            _, n = build_report(ledger, tmp, date(2026, 10, 19), date(2026, 10, 19), "validation")
            self.assertEqual(n, 1)
            _, n = build_report(ledger, tmp, date(2026, 10, 1), date(2026, 10, 31), "invoice")
            self.assertEqual(n, 2)
            _, n = build_report(ledger, tmp, date(2026, 11, 1), date(2026, 11, 30), "invoice")
            self.assertEqual(n, 0)
        finally:
            shutil.rmtree(tmp)


if __name__ == "__main__":
    unittest.main()
