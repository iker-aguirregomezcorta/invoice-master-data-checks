"""Web interface (API) tests. They use a temporary output folder."""
import csv
import shutil
import tempfile
import unittest
from collections import Counter
from datetime import date
from pathlib import Path
from unittest import mock

from fastapi.testclient import TestClient

import web
from app import service

BASE = Path(__file__).resolve().parent.parent
INVOICES = BASE / "data" / "invoices"
with open(BASE / "data" / "expected.csv", encoding="utf-8") as fh:
    EXPECTED = list(csv.DictReader(fh))
OK = [e for e in EXPECTED if e["planted_error"] == "No errors"]
POSTCODE = next(e for e in EXPECTED if e["expected_fields"] == "Postcode")
VAT = next(e for e in EXPECTED if e["expected_fields"] == "VAT ID")


class TestWeb(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        cfg = service.load_config()
        cfg["paths"]["output_folder"] = str(self.tmp)
        cfg["email"] = {"enabled": False}
        self.p = mock.patch.object(web, "current_cfg", lambda: cfg)
        self.p.start()
        self.c = TestClient(web.app)

    def tearDown(self):
        self.p.stop()
        shutil.rmtree(self.tmp)

    def upload(self, *names):
        files = [("files", (n, (INVOICES / n).read_bytes(), "application/pdf")) for n in names]
        return self.c.post("/api/validate", files=files)

    def test_page(self):
        r = self.c.get("/")
        self.assertEqual(r.status_code, 200)
        self.assertIn("Invoice checks", r.text)

    def test_upload_and_validate(self):
        r = self.upload(OK[0]["file"], POSTCODE["file"])
        statuses = {x["number"]: x["status"] for x in r.json()["results"]}
        self.assertEqual(statuses, {OK[0]["number"]: "validated", POSTCODE["number"]: "mismatch"})
        m = r.json()["results"][1]["mismatches"][0]
        self.assertEqual(m["field"], "Postcode")
        self.assertNotEqual(m["invoice"], m["master"])

    def test_duplicate(self):
        self.upload(OK[0]["file"])
        r = self.upload(OK[0]["file"])
        self.assertEqual(r.json()["results"][0]["status"], "duplicate")
        self.assertEqual(len(self.c.get("/api/ledger").json()), 1)

    def test_exception_not_repeated(self):
        self.upload(POSTCODE["file"])
        r = self.upload(POSTCODE["file"]).json()["results"][0]
        self.assertEqual(r["status"], "mismatch")
        self.assertIsNone(r["alert_file"])
        self.assertIn("Already alerted", r["alert"])
        self.assertEqual(len(self.c.get("/api/exceptions").json()), 1)

    def test_validate_folder(self):
        r = self.c.post("/api/validate-folder")
        self.assertEqual(Counter(x["status"] for x in r.json()["results"]),
                         Counter(e["expected_status"] for e in EXPECTED))

    def test_not_a_pdf(self):
        r = self.c.post("/api/validate", files=[("files", ("notes.txt", b"hello", "text/plain"))])
        self.assertEqual(r.status_code, 400)

    def test_damaged_pdf(self):
        r = self.c.post("/api/validate", files=[("files", ("broken.pdf", b"not a pdf", "application/pdf"))])
        self.assertEqual(r.json()["results"][0]["status"], "unreadable")

    def test_report_and_download(self):
        res = self.upload(OK[0]["file"], OK[1]["file"]).json()["results"]
        today = date.today().isoformat()
        r = self.c.get(f"/api/report?start={today}&end={today}&by=validation").json()
        self.assertEqual(len(r["rows"]), 2)
        self.assertAlmostEqual(r["total"], round(res[0]["total"] + res[1]["total"], 2))
        d1 = res[0]["date"]
        r = self.c.get(f"/api/report?start={d1}&end={d1}&by=invoice").json()
        self.assertIn(OK[0]["number"], [row["Invoice no."] for row in r["rows"]])
        d = self.c.get(f"/api/report/download?start={today}&end={today}")
        self.assertEqual(d.status_code, 200)
        self.assertTrue(d.content.startswith(b"PK"))  # it is an .xlsx

    def test_alert_download_and_path_protection(self):
        res = self.upload(VAT["file"]).json()["results"][0]
        self.assertEqual(self.c.get(f"/api/alerts/{res['alert_file']}").status_code, 200)
        self.assertEqual(self.c.get("/api/alerts/..%2F..%2Fconfig.json").status_code, 404)

    def test_excel_locked(self):
        with mock.patch("app.service.check_writable",
                        side_effect=service.FileLocked("exceptions.xlsx is open")):
            r = self.upload(OK[0]["file"])
        self.assertEqual(r.status_code, 409)
        self.assertIn("open", r.json()["detail"])


if __name__ == "__main__":
    unittest.main()
