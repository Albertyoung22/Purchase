"""Unit tests for the standalone Tainan City Schedule 2 travel report builder."""

import tempfile
import unittest
from pathlib import Path

import openpyxl

from travel_form_builder import (
    RouteSegment,
    TravelDayEntry,
    TravelReport,
    amount_to_chinese_upper,
    build_travel_report_xlsx,
    compute_day_amounts,
)


class TestTravelFormBuilder(unittest.TestCase):
    def test_self_car_multi_stop_loop_and_misc_tiers(self) -> None:
        report = TravelReport(
            traveler_name="王大明",
            job_title="組長",
            purpose="參加教育局會議及校際觀摩",
            destination="永華市政中心、南寧高中",
            start_date="115/11/20",
            end_date="115/11/20",
            days=[
                TravelDayEntry(
                    date="11/20",
                    summary="參加教育局會議並順道觀摩",
                    transport_mode="self_car",
                    segments=[
                        RouteSegment("土城高中", "教育局", 14.0),
                        RouteSegment("教育局", "南寧高中", 6.0),
                        RouteSegment("南寧高中", "土城高中", 18.0),
                    ],
                )
            ],
        )
        cd = compute_day_amounts(report.days[0], rank_level=report.rank_level, rates=report.rates)
        # (14 + 6 + 18) = 38 km * 3.25 = 123.5 -> round(123.5) = 124
        self.assertEqual(cd.bus_mrt_fare, 124)
        self.assertEqual(cd.route_text, "土城高中→教育局→南寧高中→土城高中")
        # One-way = 38 / 2 = 19 km (5~60 km) -> misc 120
        self.assertEqual(cd.misc_fee, 120)

    def test_lodging_cap_and_hsr_note_and_xlsx_output(self) -> None:
        report = TravelReport(
            traveler_name="李小華",
            job_title="教師",
            rank_level="簡任級以下",
            purpose="全國科展指導",
            destination="台北",
            start_date="115/11/18",
            end_date="115/11/19",
            total_days_text="共計 2 日 0 時",
            budget_code="321",
            plan_name="各校經常門分支計畫",
            days=[
                TravelDayEntry(
                    date="11/18",
                    route_text="台南→台北",
                    summary="搭高鐵北上佈展",
                    transport_mode="flight_hsr",
                    flight_hsr_fare=1350,
                    one_way_km=300.0,
                    lodging_actual=2600,  # exceeds 2000 limit
                    receipt_count=2,
                ),
                TravelDayEntry(
                    date="11/19",
                    route_text="台北→台南",
                    summary="參展及返程",
                    transport_mode="flight_hsr",
                    flight_hsr_fare=1350,
                    one_way_km=300.0,
                    lodging_actual=0,
                    receipt_count=1,
                ),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "travel_report.xlsx"
            build_travel_report_xlsx(report, out_path)
            self.assertTrue(out_path.exists())

            wb = openpyxl.load_workbook(out_path, data_only=False)
            ws = wb["國內出差旅費報告表"]
            self.assertEqual(ws["C3"].value, "李小華")
            # Day 1: HSR 1350, Lodging capped at 2000, Misc 300 (>=60km)
            self.assertEqual(ws["E8"].value, 1350)
            self.assertEqual(ws["H8"].value, 2000)
            self.assertEqual(ws["I8"].value, 300)
            self.assertEqual(ws["K8"].value, "=SUM(D8:I8)")
            self.assertIn("2000元", ws["L8"].value)
            # Numeric cells have shrinkToFit enabled to avoid ###
            self.assertTrue(ws["E8"].alignment.shrinkToFit)
            self.assertFalse(ws["E8"].alignment.wrapText)
            # Uppercase grand total: (1350+2000+300) + (1350+300) = 5300
            self.assertIn("新台幣伍仟參佰元整", ws["A15"].value)

    def test_taxi_requires_approval_and_under_5km_zero_misc(self) -> None:
        report = TravelReport()
        with self.assertRaises(ValueError):
            compute_day_amounts(
                TravelDayEntry(date="11/20", transport_mode="taxi", taxi_approved=False),
                rank_level=report.rank_level,
                rates=report.rates,
            )

        cd = compute_day_amounts(
            TravelDayEntry(date="11/20", transport_mode="official_free", one_way_km=3.5),
            rank_level=report.rank_level,
            rates=report.rates,
        )
        self.assertEqual(cd.bus_mrt_fare, 0)
        self.assertEqual(cd.misc_fee, 0)
        self.assertEqual(amount_to_chinese_upper(10205), "新台幣壹萬零貳佰零伍元整")


if __name__ == "__main__":
    unittest.main()
