import tempfile
import unittest
from pathlib import Path

import openpyxl

from app import app
from property_builder import (
    PropertyIncreaseForm,
    PropertyItem,
    build_property_increase_xlsx,
    summarize_property_form,
)


class TestPropertyBuilder(unittest.TestCase):
    def test_auto_classification_and_dual_sheet_xlsx(self):
        form = PropertyIncreaseForm(
            school_name="臺南市立測試國小",
            doc_no="PI-1151101",
            acquire_date="115/11/22",
            department="資訊室",
            items=[
                PropertyItem(name="桌上型電腦", spec="i5", unit="台", quantity=2, unit_price=22000),
                PropertyItem(name="無線滑鼠", spec="M331", unit="個", quantity=5, unit_price=600),
            ],
        )
        computed, summary = summarize_property_form(form)
        self.assertEqual(computed[0].category, "財產")
        self.assertEqual(computed[0].useful_life, 5)
        self.assertEqual(computed[1].category, "非消耗品")
        self.assertEqual(computed[1].useful_life, 2)
        self.assertEqual(summary["property_amount"], 44000)
        self.assertEqual(summary["non_consumable_amount"], 3000)

        with tempfile.TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "prop_test.xlsx"
            build_property_increase_xlsx(form, out_path)
            self.assertTrue(out_path.exists())
            wb = openpyxl.load_workbook(out_path, data_only=False)
            self.assertEqual(wb.sheetnames, ["財產增加單", "非消耗品登記清冊"])
            ws1 = wb["財產增加單"]
            self.assertEqual(ws1.page_setup.fitToWidth, 1)
            self.assertEqual(ws1["F4"].value, 2)
            self.assertEqual(ws1["G4"].value, 22000)
            self.assertEqual(ws1["H4"].value, "=F4*G4")

    def test_property_api_endpoints_and_redownload(self):
        client = app.test_client()
        payload = {
            "school_name": "臺南市立測試國小",
            "doc_no": "PI-001",
            "department": "總務處",
            "items": [
                {"name": "雷射印表機", "unit": "台", "quantity": 1, "unit_price": 12500},
                {"name": "白板", "unit": "面", "quantity": 2, "unit_price": 1800},
            ],
        }
        res_prev = client.post("/api/property/preview", json=payload)
        self.assertEqual(res_prev.status_code, 200)
        d_prev = res_prev.get_json()
        self.assertTrue(d_prev["ok"])
        self.assertEqual(d_prev["summary"]["grand_total"], 16100)

        res_gen = client.post("/api/property/generate", json=payload)
        self.assertEqual(res_gen.status_code, 200)
        d_gen = res_gen.get_json()
        self.assertTrue(d_gen["ok"])
        entry = d_gen["history_entry"]
        self.assertEqual(entry["property_count"], 1)
        self.assertEqual(entry["non_consumable_count"], 1)

        res_pp = client.post("/api/property/print-preview", json=payload)
        self.assertEqual(res_pp.status_code, 200)
        self.assertTrue(res_pp.get_json()["ok"])

        res_re = client.post("/api/property/history/redownload", json={"record": entry})
        self.assertEqual(res_re.status_code, 200)
        self.assertGreater(len(res_re.data), 1000)
        res_re.close()

        res_del = client.delete(f"/api/property/history/{entry['id']}")
        self.assertEqual(res_del.status_code, 200)


if __name__ == "__main__":
    unittest.main()
