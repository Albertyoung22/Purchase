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
    def test_self_car_and_scooter_multi_stop_loop_and_misc_tiers(self) -> None:
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
        cd = compute_day_amounts(report.days[0], rates=report.rates)
        # 38 km * 3.0 = 114
        self.assertEqual(cd.self_car_fare, 114)
        self.assertEqual(cd.route_text, "土城高中→教育局→南寧高中→土城高中")
        # One-way = 38 / 2 = 19 km (5~60 km) -> misc 120
        self.assertEqual(cd.misc_fee, 120)

        # Self-scooter: 38 km * 2.0 = 76
        scooter_entry = TravelDayEntry(
            date="11/20",
            transport_mode="self_scooter",
            segments=report.days[0].segments,
        )
        cd_sc = compute_day_amounts(scooter_entry, rates=report.rates)
        self.assertEqual(cd_sc.self_scooter_fare, 76)

    def test_lodging_weekday_holiday_caps_and_xlsx_output(self) -> None:
        report = TravelReport(
            traveler_name="李小華",
            job_title="教師",
            rank_level="薦任",
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
                    is_holiday_lodging=False,
                    lodging_actual=4000,  # exceeds weekday 3500 limit
                    receipt_count=2,
                ),
                TravelDayEntry(
                    date="11/19",
                    route_text="台北→台南",
                    summary="參展及返程",
                    transport_mode="flight_hsr",
                    flight_hsr_fare=1350,
                    one_way_km=300.0,
                    is_holiday_lodging=True,
                    lodging_actual=4800,  # exceeds holiday 4500 limit
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
            self.assertEqual(ws["A2"].value, "採購（費用動支）申請單暨粘貼憑證用紙")
            self.assertEqual(ws["C9"].value, "李小華")
            # Day 1 (col C): HSR 1350 (row 15), Lodging capped at 3500 (row 19), Misc 300 (row 21)
            self.assertEqual(ws["C15"].value, 1350)
            self.assertEqual(ws["C19"].value, 3500)
            self.assertEqual(ws["C21"].value, 300)
            self.assertEqual(ws["C22"].value, "=SUM(C15:C21)")
            self.assertIn("3500元", ws["C23"].value)
            # Day 2 (col D): Holiday lodging capped at 4500
            self.assertEqual(ws["D19"].value, 4500)
            self.assertIn("4500元", ws["D23"].value)
            # Row total formulas in Col I
            self.assertEqual(ws["I15"].value, "=SUM(C15:H15)")
            self.assertEqual(ws["I22"].value, "=SUM(C22:H22)")
            # Numeric cells have shrinkToFit enabled to avoid ###
            self.assertTrue(ws["C15"].alignment.shrinkToFit)
            self.assertFalse(ws["C15"].alignment.wrapText)

    def test_taxi_requires_approval_and_under_5km_zero_misc(self) -> None:
        report = TravelReport()
        with self.assertRaises(ValueError):
            compute_day_amounts(
                TravelDayEntry(date="11/20", transport_mode="taxi", taxi_approved=False),
                rates=report.rates,
            )

        cd = compute_day_amounts(
            TravelDayEntry(date="11/20", transport_mode="official_free", one_way_km=3.5),
            rates=report.rates,
        )
        self.assertEqual(cd.self_car_fare, 0)
        self.assertEqual(cd.misc_fee, 0)
        self.assertEqual(amount_to_chinese_upper(10205), "新台幣壹萬零貳佰零伍元整")

        # Lodging under 60km should add approval reminder note (要點第9點)
        cd_near_lodging = compute_day_amounts(
            TravelDayEntry(date="11/20", transport_mode="self_car", one_way_km=25.0, lodging_actual=1500),
            rates=report.rates,
        )
        self.assertIn("須事前經機關核准", cd_near_lodging.note)


if __name__ == "__main__":
    unittest.main()
