"""Standalone builder for Tainan City Schedule 2 (臺南市政府附表二) Domestic Business Travel Expense Report."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


@dataclass
class TravelRateConfig:
    """Latest Tainan City domestic business travel reimbursement rates (附表一 & 要點第5、11點)."""

    car_rate_per_km: float = 3.0  # 駕駛自用或租賃汽車每公里 3 元
    scooter_rate_per_km: float = 2.0  # 駕駛自用或租賃機車每公里 2 元
    lodging_limit_weekday: int = 3500  # 平日住宿費每日上限 3,500 元
    lodging_limit_holiday: int = 4500  # 假日住宿費每日上限 4,500 元
    misc_far_km: float = 60.0  # 單程超過 60 公里
    misc_near_km: float = 5.0  # 單程 5 公里以上、60 公里以下
    misc_far_amount: int = 300
    misc_near_amount: int = 120
    misc_under_5km_amount: int = 0


@dataclass
class RouteSegment:
    """Single leg of a multi-stop self-driven business trip (e.g. A -> B)."""

    origin: str
    destination: str
    distance_km: float = 0.0


@dataclass
class TravelDayEntry:
    """One day's travel itinerary and expenses."""

    date: str
    route_text: str = ""
    summary: str = ""
    # self_car | self_scooter | train | flight_hsr | bus_mrt | ship | ubike | taxi | official_free
    transport_mode: str = "self_car"
    segments: list[RouteSegment] = field(default_factory=list)
    one_way_km: float | None = None
    bus_equivalent_fare: int | float | None = None
    train_fare: int | float = 0
    flight_hsr_fare: int | float = 0
    bus_mrt_fare: int | float = 0
    ship_fare: int | float = 0
    ubike_fare: int | float = 0
    self_car_fare: int | float = 0
    self_scooter_fare: int | float = 0
    taxi_approved: bool = False
    is_holiday_lodging: bool = False  # True = 假日(4500), False = 平日(3500)
    lodging_actual: int | float = 0
    misc_fee: int | float | None = None
    receipt_count: int = 0
    receipt_no: str = ""
    note: str = ""


@dataclass
class TravelReport:
    """Full domestic business travel voucher + report (臺南市政府附表二)."""

    school_name: str = "臺南市立土城高級中學"
    traveler_name: str = ""
    job_title: str = ""
    rank_level: str = ""
    purpose: str = ""
    destination: str = ""
    start_date: str = ""
    end_date: str = ""
    total_days_text: str = ""
    budget_code: str = ""
    plan_name: str = ""
    work_plan: str = ""
    level1_3: str = ""
    days: list[TravelDayEntry] = field(default_factory=list)
    rates: TravelRateConfig = field(default_factory=TravelRateConfig)


@dataclass
class ComputedDay:
    """Resolved amounts and notes for a single TravelDayEntry."""

    date: str
    route_text: str
    summary: str
    train_fare: int
    flight_hsr_fare: int
    bus_mrt_fare: int
    ship_fare: int
    ubike_fare: int
    self_car_fare: int
    self_scooter_fare: int
    lodging_fee: int
    misc_fee: int
    receipt_count: int
    receipt_no: str
    note: str

    @property
    def subtotal(self) -> int:
        return (
            self.train_fare
            + self.flight_hsr_fare
            + self.bus_mrt_fare
            + self.ship_fare
            + self.ubike_fare
            + self.self_car_fare
            + self.self_scooter_fare
            + self.lodging_fee
            + self.misc_fee
        )


def amount_to_chinese_upper(amount: int) -> str:
    """Convert a non-negative integer NTD amount to Traditional Chinese financial uppercase."""
    if amount <= 0:
        return "新台幣零元整"

    digits = "零壹貳參肆伍陸柒捌玖"
    small_units = ["", "拾", "佰", "仟"]
    big_units = ["", "萬", "億"]

    s = str(int(amount))
    groups: list[str] = []
    while s:
        groups.append(s[-4:])
        s = s[:-4]

    result_parts: list[str] = []
    for g_idx, grp in enumerate(groups):
        val = int(grp)
        if val == 0:
            if result_parts and not result_parts[-1].startswith("零"):
                result_parts.append("零")
            continue
        grp_str = ""
        leading_zero = g_idx < len(groups) - 1 and grp[0] == "0"
        zero_pending = False
        for i, ch in enumerate(grp):
            d = int(ch)
            pos = len(grp) - 1 - i
            if d == 0:
                zero_pending = True
            else:
                if (zero_pending and grp_str) or (leading_zero and not grp_str):
                    grp_str += "零"
                grp_str += digits[d] + small_units[pos]
                zero_pending = False
        result_parts.append(grp_str + big_units[g_idx])

    joined = "".join(reversed(result_parts)).strip("零")
    return f"新台幣{joined}元整"


def _format_route_from_segments(segments: list[RouteSegment]) -> str:
    if not segments:
        return ""
    stops = [segments[0].origin]
    for seg in segments:
        if seg.destination:
            stops.append(seg.destination)
    return "→".join(s for s in stops if s)


def _infer_one_way_km(entry: TravelDayEntry) -> float:
    if entry.one_way_km is not None:
        return max(0.0, float(entry.one_way_km))
    if not entry.segments:
        return 0.0
    total_km = sum(max(0.0, float(s.distance_km)) for s in entry.segments)
    if len(entry.segments) >= 2 and entry.segments[0].origin == entry.segments[-1].destination:
        return total_km / 2.0
    return total_km


def compute_day_amounts(
    entry: TravelDayEntry,
    *,
    rates: TravelRateConfig,
    is_same_day_round_trip: bool = False,
) -> ComputedDay:
    """Apply Tainan City travel rules to compute transport, lodging, and misc fees for one day."""
    notes: list[str] = [entry.note.strip()] if entry.note.strip() else []
    route_text = entry.route_text.strip() or _format_route_from_segments(entry.segments)

    train = int(round(entry.train_fare))
    flight_hsr = int(round(entry.flight_hsr_fare))
    bus_mrt = int(round(entry.bus_mrt_fare))
    ship = int(round(entry.ship_fare))
    ubike = int(round(entry.ubike_fare))
    self_car = int(round(entry.self_car_fare))
    self_scooter = int(round(entry.self_scooter_fare))

    if entry.transport_mode == "official_free":
        train = flight_hsr = bus_mrt = ship = ubike = self_car = self_scooter = 0
        notes.append("機關專備交通工具/免費票，交通費0元")
    elif entry.transport_mode == "taxi":
        if not entry.taxi_approved:
            raise ValueError("依報支要點第5點規定，搭乘計程車須因業務需要經機關核准方可報支")
    elif entry.transport_mode == "self_car":
        if entry.bus_equivalent_fare is not None:
            self_car = int(round(entry.bus_equivalent_fare))
            notes.append("自用汽車按比照票價報支")
        elif entry.segments:
            total_km = sum(max(0.0, float(s.distance_km)) for s in entry.segments)
            self_car = int(round(total_km * rates.car_rate_per_km))
            notes.append(f"自用汽車({total_km:g}km×{rates.car_rate_per_km}元)")
    elif entry.transport_mode == "self_scooter":
        if entry.bus_equivalent_fare is not None:
            self_scooter = int(round(entry.bus_equivalent_fare))
            notes.append("自用機車按比照票價報支")
        elif entry.segments:
            total_km = sum(max(0.0, float(s.distance_km)) for s in entry.segments)
            self_scooter = int(round(total_km * rates.scooter_rate_per_km))
            notes.append(f"自用機車({total_km:g}km×{rates.scooter_rate_per_km}元)")
    elif entry.transport_mode == "flight_hsr" and flight_hsr > 0:
        if is_same_day_round_trip:
            notes.append("當日往返無須檢附高鐵票根")
        else:
            notes.append("需檢附票根或購票證明")

    # Lodging cap (附表一: 平日 3,500 / 假日 4,500)
    lodging_limit = rates.lodging_limit_holiday if entry.is_holiday_lodging else rates.lodging_limit_weekday
    raw_lodging = max(0, int(round(entry.lodging_actual)))
    if raw_lodging > lodging_limit:
        lodging_fee = lodging_limit
        day_kind = "假日" if entry.is_holiday_lodging else "平日"
        notes.append(f"住宿費依{day_kind}上限{lodging_limit}元報支")
    else:
        lodging_fee = raw_lodging

    # Misc fee tiered by one-way distance (要點第11點: >60km: 300, 5~60km: 120, <5km: 0)
    if entry.misc_fee is not None:
        misc_fee = max(0, int(round(entry.misc_fee)))
    else:
        one_way = _infer_one_way_km(entry)
        if one_way > rates.misc_far_km:
            misc_fee = rates.misc_far_amount
        elif one_way >= rates.misc_near_km:
            misc_fee = rates.misc_near_amount
        else:
            misc_fee = rates.misc_under_5km_amount

    return ComputedDay(
        date=entry.date,
        route_text=route_text,
        summary=entry.summary,
        train_fare=train,
        flight_hsr_fare=flight_hsr,
        bus_mrt_fare=bus_mrt,
        ship_fare=ship,
        ubike_fare=ubike,
        self_car_fare=self_car,
        self_scooter_fare=self_scooter,
        lodging_fee=lodging_fee,
        misc_fee=misc_fee,
        receipt_count=entry.receipt_count,
        receipt_no=entry.receipt_no,
        note="；".join(n for n in notes if n),
    )


_THIN = Side(style="thin", color="000000")
_BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FONT_TITLE = Font(name="標楷體", size=14, bold=True)
_FONT_SUB = Font(name="標楷體", size=12, bold=True)
_FONT_BOLD = Font(name="標楷體", size=9, bold=True)
_FONT_BODY = Font(name="標楷體", size=9)
_FONT_SMALL = Font(name="標楷體", size=8)

_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
_ALIGN_NUM = Alignment(horizontal="right", vertical="center", wrap_text=False, shrinkToFit=True)


def _split_amount_digits_5(amount: int) -> list[str]:
    """Split integer NTD into [萬, 千, 百, 十, 元] cells matching the school .doc voucher."""
    if amount <= 0:
        return ["", "", "", "", "0"]
    s = str(int(amount)).rjust(5, " ")
    if len(s) > 5:
        s = s[-5:]
    return list(s)


def _format_roc_date_phrase(roc_date: str) -> str:
    """Convert '115/11/21' to '115 年 11 月 21 日', or return raw string."""
    parts = [p.strip() for p in roc_date.replace("-", "/").split("/") if p.strip()]
    if len(parts) == 3:
        return f"{parts[0]} 年 {parts[1]} 月 {parts[2]} 日"
    return roc_date or "   年   月   日"


def build_travel_report_xlsx(report: TravelReport, output_path: Path) -> Path:
    """Create a print-ready .xlsx matching the school's '國內出差旅費報告表.doc' A4 portrait format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "國內出差旅費報告表"

    is_same_day = bool(report.start_date and report.start_date == report.end_date) or len(report.days) == 1
    computed_days = [
        compute_day_amounts(d, rates=report.rates, is_same_day_round_trip=is_same_day)
        for d in report.days
    ]
    grand_total = sum(cd.subtotal for cd in computed_days)
    total_receipts = sum(cd.receipt_count for cd in computed_days)

    num_day_cols = max(6, len(computed_days))
    total_col_idx = 3 + num_day_cols  # A,B = labels; C.. = days; last = 合計
    last_col_letter = get_column_letter(total_col_idx)

    # Column widths tuned for A4 portrait
    ws.column_dimensions["A"].width = 6.5
    ws.column_dimensions["B"].width = 13.5
    for idx in range(num_day_cols):
        ws.column_dimensions[get_column_letter(3 + idx)].width = 11.5
    ws.column_dimensions[last_col_letter].width = 10.5

    # --- Top Voucher Section: 採購（費用動支）申請單暨粘貼憑證用紙 ---
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = report.school_name
    ws["A1"].font = _FONT_TITLE
    ws["A1"].alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 22

    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = "採購（費用動支）申請單暨粘貼憑證用紙"
    ws["A2"].font = _FONT_TITLE
    ws["A2"].alignment = _ALIGN_CENTER
    ws.row_dimensions[2].height = 22

    ws.merge_cells("A3:E3")
    ws["A3"] = "傳票(付款憑單)編號："
    ws["A3"].font = _FONT_BODY
    ws["A3"].alignment = _ALIGN_LEFT
    ws.merge_cells(f"F3:{last_col_letter}3")
    ws["F3"] = "簽證編號："
    ws["F3"].font = _FONT_BODY
    ws["F3"].alignment = _ALIGN_LEFT
    ws.row_dimensions[3].height = 18

    # Voucher table (Rows 4-6)
    amt_start_idx = total_col_idx - 4  # 5 columns for 萬 千 百 十 元
    amt_start_l = get_column_letter(amt_start_idx)
    budget_end_l = get_column_letter(amt_start_idx - 1)

    ws.merge_cells("A4:A5")
    ws["A4"] = "憑證編號"
    ws.merge_cells(f"B4:{budget_end_l}5")
    ws["B4"] = "預   算   科   目"
    ws.merge_cells(f"{amt_start_l}4:{last_col_letter}4")
    ws[f"{amt_start_l}4"] = "金        額"
    for idx, lbl in enumerate(["萬", "千", "百", "十", "元"]):
        ws.cell(5, amt_start_idx + idx, lbl)

    ws["A6"] = report.budget_code
    ws.merge_cells(f"B6:{budget_end_l}6")
    is_agency = bool(report.budget_code and report.budget_code[0].upper() in ("L", "R", "T", "E"))
    chk_reg = "□" if is_agency else "☑"
    chk_ag = "☑" if is_agency else "□"
    plan_extra = f"（{report.plan_name}）" if report.plan_name else ""
    ws["B6"] = (
        f"計劃名稱：{chk_reg}各校經常門分支計畫{chk_ag}應付代收款{plan_extra}\n"
        f"一級用途別：  服務費用  \n"
        f"用  途  別：  旅運費--國內旅費"
    )
    for idx, dv in enumerate(_split_amount_digits_5(grand_total)):
        ws.cell(6, amt_start_idx + idx, dv)

    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 48
    for r in range(4, 7):
        for c in range(1, total_col_idx + 1):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_BOLD if r in (4, 5) else _FONT_BODY
            cell.alignment = _ALIGN_LEFT if (r == 6 and c == 2) else _ALIGN_CENTER

    # --- Main Report Title (Rows 7-8) ---
    ws.merge_cells(f"A7:{last_col_letter}7")
    ws["A7"] = f"{report.school_name}國內出差旅費報告表"
    ws["A7"].font = _FONT_TITLE
    ws["A7"].alignment = _ALIGN_CENTER
    ws.row_dimensions[7].height = 24

    ws.merge_cells(f"A8:{last_col_letter}8")
    ws["A8"] = "第  1  頁共  1  頁"
    ws["A8"].font = _FONT_BODY
    ws["A8"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[8].height = 18

    # --- Header Info Rows (Rows 9-11) ---
    ws.merge_cells("A9:B9")
    ws["A9"] = "姓      名"
    ws.merge_cells("C9:D9")
    ws["C9"] = report.traveler_name
    ws["E9"] = "職稱"
    ws["F9"] = report.job_title
    ws["G9"] = "職等"
    ws.merge_cells(f"H9:{last_col_letter}9")
    ws["H9"] = report.rank_level

    ws.merge_cells("A10:B10")
    ws["A10"] = "出  差  事  由"
    ws.merge_cells(f"C10:{last_col_letter}10")
    ws["C10"] = report.purpose

    ws.merge_cells(f"A11:{last_col_letter}11")
    s_phrase = _format_roc_date_phrase(report.start_date)
    e_phrase = _format_roc_date_phrase(report.end_date)
    days_phrase = report.total_days_text or f"共計 {len(computed_days)} 日"
    ws["A11"] = f"中華民國 {s_phrase} 起至 {e_phrase} 止  {days_phrase}  附單據  {total_receipts}  張"

    for r in range(9, 12):
        ws.row_dimensions[r].height = 22
        for c in range(1, total_col_idx + 1):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_BOLD if (r in (9, 10) and c in (1, 5, 7)) else _FONT_BODY
            cell.alignment = _ALIGN_LEFT if (r == 10 and c >= 3) else _ALIGN_CENTER

    # --- Transposed Expense Table (Rows 12-23) ---
    # Row 12: 月日 | Day 1..N | 合計
    ws.merge_cells("A12:B12")
    ws["A12"] = "月  /  日"
    for idx, cd in enumerate(computed_days):
        ws.cell(12, 3 + idx, cd.date)
    ws.cell(12, total_col_idx, "合計")
    ws.row_dimensions[12].height = 22

    # Row 13: 起訖地點
    ws.merge_cells("A13:B13")
    ws["A13"] = "起訖地點"
    for idx, cd in enumerate(computed_days):
        ws.cell(13, 3 + idx, cd.route_text)
    ws.row_dimensions[13].height = 30

    # Row 14: 工作記要
    ws.merge_cells("A14:B14")
    ws["A14"] = "工作記要"
    for idx, cd in enumerate(computed_days):
        ws.cell(14, 3 + idx, cd.summary)
    ws.row_dimensions[14].height = 30

    # Rows 15-18: 交通費 (4 rows)
    ws.merge_cells("A15:A18")
    ws["A15"] = "交\n通\n費"
    has_hsr = any(cd.flight_hsr_fare > 0 for cd in computed_days)
    has_car_bus = any((cd.bus_mrt_fare + cd.ubike_fare + cd.self_car_fare + cd.self_scooter_fare) > 0 for cd in computed_days)
    ws["B15"] = f"□飛機{'☑' if has_hsr else '□'}高鐵"
    ws["B16"] = f"{'☑' if has_car_bus else '□'}汽車□捷運"
    ws["B17"] = "火車"
    ws["B18"] = "輪船"

    for idx, cd in enumerate(computed_days):
        col_c = 3 + idx
        car_bus_sum = cd.bus_mrt_fare + cd.ubike_fare + cd.self_car_fare + cd.self_scooter_fare
        ws.cell(15, col_c, cd.flight_hsr_fare or None)
        ws.cell(16, col_c, car_bus_sum or None)
        ws.cell(17, col_c, cd.train_fare or None)
        ws.cell(18, col_c, cd.ship_fare or None)

    # Row 19: 住宿費
    ws.merge_cells("A19:B19")
    ws["A19"] = "住宿費"
    for idx, cd in enumerate(computed_days):
        ws.cell(19, 3 + idx, cd.lodging_fee or None)

    # Row 20: 住宿費加計交通費(旅行業代收轉付)
    ws.merge_cells("A20:B20")
    ws["A20"] = "住宿費加計交通費\n(旅行業代收轉付)"

    # Row 21: 雜費
    ws.merge_cells("A21:B21")
    ws["A21"] = "雜費"
    for idx, cd in enumerate(computed_days):
        ws.cell(21, 3 + idx, cd.misc_fee or None)

    # Row 22: 總計
    ws.merge_cells("A22:B22")
    ws["A22"] = "總  計"
    for idx in range(num_day_cols):
        col_l = get_column_letter(3 + idx)
        ws.cell(22, 3 + idx, f"=SUM({col_l}15:{col_l}21)")

    # 合計 column (Col I) formulas for rows 15..22
    last_day_l = get_column_letter(2 + num_day_cols)
    for r in range(15, 23):
        ws.cell(r, total_col_idx, f"=SUM(C{r}:{last_day_l}{r})")

    # Row 23: 備註
    ws.merge_cells("A23:B23")
    ws["A23"] = "備  註"
    for idx, cd in enumerate(computed_days):
        ws.cell(23, 3 + idx, cd.note)

    row_heights = {15: 22, 16: 22, 17: 22, 18: 22, 19: 22, 20: 28, 21: 22, 22: 24, 23: 34}
    for r, h in row_heights.items():
        ws.row_dimensions[r].height = h

    for r in range(12, 24):
        for c in range(1, total_col_idx + 1):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            is_num = (15 <= r <= 22) and (c >= 3)
            if is_num:
                cell.font = _FONT_BOLD if (r == 22 or c == total_col_idx) else _FONT_BODY
                cell.alignment = _ALIGN_NUM
                cell.number_format = "#,##0"
            elif r in (13, 14, 23) and c >= 3:
                cell.font = _FONT_SMALL
                cell.alignment = _ALIGN_LEFT
            else:
                cell.font = _FONT_BOLD if (r == 12 or c <= 2) else _FONT_BODY
                cell.alignment = _ALIGN_CENTER

    # --- Bottom Signature Row (Row 24) ---
    ws.row_dimensions[24].height = 30
    ws.merge_cells(f"A24:{last_col_letter}24")
    ws["A24"] = "出差人              單位主管              人事單位              會計單位              基金主持人"
    ws["A24"].font = _FONT_BOLD
    ws["A24"].alignment = _ALIGN_LEFT

    # Single-page A4 portrait print setup matching the .doc template
    ws.print_area = f"A1:{last_col_letter}24"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
