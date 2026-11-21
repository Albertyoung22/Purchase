import tempfile
import unittest
from pathlib import Path

import openpyxl

from app import app
from substitute_builder import (
    SubstituteEntry,
    SubstituteRoster,
    build_substitute_roster_xlsx,
    compute_substitute_entry,
    lookup_insurance_brackets,
)


class TestSubstituteBuilder(unittest.TestCase):
    def test_school_level_hourly_rates_and_parttime_2nd_gen_nhi(self):
        # Elementary 405 * 20 = 8,100 -> part-time LI bracket 11,100, 2nd-gen NHI 2.11% = 171
        e = SubstituteEntry(
            teacher_name="王小明",
            school_level="elementary",
            periods=20,
            enroll_mode="part_time",
            insured_days=30,
        )
        c = compute_substitute_entry(e, roc_year=115)
        self.assertEqual(c.hourly_rate, 405)
        self.assertEqual(c.gross_pay, 8100)
        self.assertEqual(c.li_bracket, 11100)
        self.assertEqual(c.nhi_bracket, 0)
        self.assertEqual(c.pension_bracket, 8700)
        self.assertEqual(c.li_employee, round(11100 * 0.125 * 0.20))
        self.assertEqual(c.li_employer, round(11100 * 0.125 * 0.70))
        self.assertEqual(c.pension_employer, round(8700 * 0.06))
        self.assertEqual(c.nhi_supp, round(8100 * 0.0211))
        self.assertEqual(c.net_pay, 8100 - c.li_employee - c.nhi_supp)

    def test_115_vs_114_full_month_floor_and_prorated_days(self):
        b115 = lookup_insurance_brackets(25000, year_roc=115, enroll_mode="full_month")
        b114 = lookup_insurance_brackets(25000, year_roc=114, enroll_mode="full_month")
        self.assertEqual(b115["li"], 29500)
        self.assertEqual(b115["nhi"], 29500)
        self.assertEqual(b114["li"], 28590)
        self.assertEqual(b114["nhi"], 28590)

        # Partial month 15 days -> half LI & Pension, NHI full month if end-of-month active
        e = SubstituteEntry(
            teacher_name="陳老師",
            school_level="junior",
            periods=40,
            enroll_mode="full_month",
            insured_days=15,
            dependents=1,
            voluntary_pension_pct=6.0,
            nhi_end_of_month=True,
        )
        c = compute_substitute_entry(e, roc_year=115)
        self.assertEqual(c.hourly_rate, 455)
        self.assertEqual(c.gross_pay, 18200)
        self.assertEqual(c.li_employee, round(29500 * 0.125 * 0.20 * 0.5))
        self.assertEqual(c.nhi_employee, round(29500 * 0.0517 * 0.30 * 2))
        self.assertEqual(c.pension_employee, round(19047 * 0.06 * 0.5))
        self.assertEqual(c.nhi_supp, 0)

    def test_senior_rate_and_build_xlsx(self):
        roster = SubstituteRoster(
            school_name="臺南市立範例高級中學",
            roc_year=115,
            month=11,
            plan_name="115學年度代理代課鐘點費",
            entries=[
                SubstituteEntry(teacher_name="林老師", school_level="senior", periods=10, enroll_mode="supp_only"),
                SubstituteEntry(teacher_name="黃老師", school_level="custom", hourly_rate=600, periods=5, enroll_mode="part_time"),
            ],
        )
        with tempfile.TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "sub_test.xlsx"
            build_substitute_roster_xlsx(roster, out_path)
            self.assertTrue(out_path.exists())
            wb = openpyxl.load_workbook(out_path, data_only=False)
            self.assertEqual(wb.sheetnames, ["鐘點費印領清冊", "勞健保勞退機關負擔明細"])
            ws1 = wb["鐘點費印領清冊"]
            self.assertIn("臺南市立範例高級中學", ws1["A1"].value)
            self.assertEqual(ws1["F4"].value, 505)
            self.assertEqual(ws1["G4"].value, "=E4*F4")
            self.assertEqual(ws1["M4"].value, "=G4-SUM(H4:L4)")

    def test_substitute_api_endpoints(self):
        client = app.test_client()
        payload = {
            "school_name": "臺南市立測試國小",
            "roc_year": 115,
            "month": 11,
            "plan_name": "短期代課",
            "entries": [
                {
                    "teacher_name": "張老師",
                    "school_level": "elementary",
                    "periods": 16,
                    "enroll_mode": "part_time",
                    "insured_days": 30,
                }
            ],
        }
        res_prev = client.post("/api/substitute/preview", json=payload)
        self.assertEqual(res_prev.status_code, 200)
        d_prev = res_prev.get_json()
        self.assertTrue(d_prev["ok"])
        self.assertEqual(d_prev["totals"]["gross_pay"], 6480)

        res_gen = client.post("/api/substitute/generate", json=payload)
        self.assertEqual(res_gen.status_code, 200)
        d_gen = res_gen.get_json()
        self.assertTrue(d_gen["ok"])
        self.assertTrue(d_gen["filename"].endswith(".xlsx"))

        res_pp = client.post("/api/substitute/print-preview", json=payload)
        self.assertEqual(res_pp.status_code, 200)
        d_pp = res_pp.get_json()
        self.assertTrue(d_pp["ok"])


if __name__ == "__main__":
    unittest.main()
