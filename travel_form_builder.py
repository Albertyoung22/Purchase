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


def _split_amount_digits(amount: int) -> list[str]:
    """Split integer NTD into [十萬, 萬, 千, 百, 十, 元, 角, 分] cells."""
    if amount <= 0:
        return ["", "", "", "", "", "0", "0", "0"]
    s = str(int(amount)).rjust(6, " ")
    if len(s) > 6:
        s = s[-6:]
    return list(s) + ["0", "0"]


def build_travel_report_xlsx(report: TravelReport, output_path: Path) -> Path:
    """Create a print-ready Tainan City Schedule 2 (粘貼憑證用紙 + 出差旅費報告表) .xlsx file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "出差旅費報告表"

    # 15 columns (A..O) matching official 附表二
    col_widths = {
        "A": 7.5,  # 年月日
        "B": 15.0,  # 起訖地點
        "C": 17.0,  # 工作摘要
        "D": 7.5,  # 火車費
        "E": 8.5,  # 飛機及高鐵
        "F": 8.5,  # 汽車及捷運
        "G": 6.5,  # 船舶
        "H": 7.5,  # 公共自行車
        "I": 9.5,  # 駕駛自用或租賃汽車
        "J": 9.5,  # 駕駛自用或租賃機車
        "K": 8.5,  # 住宿費
        "L": 7.5,  # 雜費
        "M": 7.5,  # 單據號數
        "N": 9.5,  # 總計
        "O": 16.0,  # 備註
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    is_same_day = bool(report.start_date and report.start_date == report.end_date) or len(report.days) == 1
    computed_days = [
        compute_day_amounts(d, rates=report.rates, is_same_day_round_trip=is_same_day)
        for d in report.days
    ]
    grand_total = sum(cd.subtotal for cd in computed_days)
    total_receipts = sum(cd.receipt_count for cd in computed_days)

    # --- Top Voucher Section: 粘貼憑證用紙 (Rows 1-7) ---
    ws.merge_cells("A1:O1")
    ws["A1"] = report.school_name
    ws["A1"].font = _FONT_TITLE
    ws["A1"].alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 22

    ws.merge_cells("A2:O2")
    ws["A2"] = "粘貼憑證用紙"
    ws["A2"].font = _FONT_SUB
    ws["A2"].alignment = _ALIGN_CENTER
    ws.row_dimensions[2].height = 20

    ws.merge_cells("A3:C4")
    ws["A3"] = "憑單編號：旅字第      號"
    ws.merge_cells("D3:E4")
    ws["D3"] = "預  算  科  目"
    ws.merge_cells("F3:M3")
    ws["F3"] = "金       額"
    ws.merge_cells("N3:O3")
    ws["N3"] = "（此欄由主計室填用）"

    digit_labels = ["十萬", "萬", "千", "百", "十", "元", "角", "分"]
    for idx, lbl in enumerate(digit_labels):
        ws.cell(4, 6 + idx, lbl)

    ws.merge_cells("N4:O5")
    ws["N4"] = "預借旅費：$"

    ws.merge_cells("A5:C5")
    ws["A5"] = "傳  票  號  碼："
    ws.merge_cells("D5:E5")
    ws["D5"] = f"業務計畫：{report.plan_name}"
    ws.merge_cells("A6:C6")
    ws["A6"] = "  月  日  支      號"
    ws.merge_cells("D6:E6")
    ws["D6"] = f"工作計畫／用途別：{report.budget_code} {report.work_plan or report.level1_3}".strip()
    ws.merge_cells("N6:O6")
    ws["N6"] = "預算控制登記："

    digits_vals = _split_amount_digits(grand_total)
    for idx, dv in enumerate(digits_vals):
        ws.merge_cells(start_row=5, start_column=6 + idx, end_row=6, end_column=6 + idx)
        ws.cell(5, 6 + idx, dv)

    for r in range(3, 7):
        ws.row_dimensions[r].height = 19
        for c in range(1, 16):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_BOLD if (r == 3 or (r == 4 and 6 <= c <= 13)) else _FONT_BODY
            cell.alignment = _ALIGN_LEFT if c in (1, 4, 14) and r >= 5 else _ALIGN_CENTER

    ws.merge_cells("A7:O7")
    ws["A7"] = "（檢附單據粘貼處）"
    ws["A7"].font = _FONT_SMALL
    ws["A7"].alignment = _ALIGN_CENTER
    ws.row_dimensions[7].height = 16

    # --- Main Report Title & Info (Rows 8-12) ---
    ws.merge_cells("A8:O8")
    ws["A8"] = "出差旅費報告表"
    ws["A8"].font = _FONT_SUB
    ws["A8"].alignment = _ALIGN_CENTER
    ws.row_dimensions[8].height = 22

    ws["A9"] = "姓名"
    ws.merge_cells("B9:D9")
    ws["B9"] = report.traveler_name
    ws["E9"] = "職稱"
    ws.merge_cells("F9:H9")
    ws["F9"] = report.job_title
    ws["I9"] = "職等"
    ws.merge_cells("J9:O9")
    ws["J9"] = report.rank_level

    ws["A10"] = "出差事由"
    ws.merge_cells("B10:O10")
    ws["B10"] = report.purpose

    ws["A11"] = "出差地點"
    ws.merge_cells("B11:O11")
    ws["B11"] = report.destination

    ws["A12"] = "起訖日期"
    ws.merge_cells("B12:K12")
    date_span = f"{report.start_date} 起至 {report.end_date} 止".strip()
    if report.total_days_text:
        date_span = f"{date_span}（{report.total_days_text}）"
    ws["B12"] = date_span
    ws.merge_cells("L12:O12")
    ws["L12"] = f"附單據  {total_receipts}  張"

    for r in range(9, 13):
        ws.row_dimensions[r].height = 20
        for c in range(1, 16):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_BOLD if c in (1, 5, 9) and r == 9 or c == 1 else _FONT_BODY
            cell.alignment = _ALIGN_LEFT if c > 1 and r >= 10 else _ALIGN_CENTER

    # --- Table Header (Rows 13-14) ---
    ws.merge_cells("A13:A14")
    ws["A13"] = "年月日"
    ws.merge_cells("B13:B14")
    ws["B13"] = "起訖地點"
    ws.merge_cells("C13:C14")
    ws["C13"] = "工作摘要"
    ws.merge_cells("D13:J13")
    ws["D13"] = "交        通        費"
    sub_headers = ["火車費", "飛機及高鐵", "汽車及捷運", "船舶", "公共自行車", "駕駛自用或租賃汽車", "駕駛自用或租賃機車"]
    for idx, sh in enumerate(sub_headers):
        ws.cell(14, 4 + idx, sh)
    ws.merge_cells("K13:K14")
    ws["K13"] = "住宿費"
    ws.merge_cells("L13:L14")
    ws["L13"] = "雜費"
    ws.merge_cells("M13:M14")
    ws["M13"] = "單據號數"
    ws.merge_cells("N13:N14")
    ws["N13"] = "總計"
    ws.merge_cells("O13:O14")
    ws["O13"] = "備註"

    ws.row_dimensions[13].height = 18
    ws.row_dimensions[14].height = 30
    for r in (13, 14):
        for c in range(1, 16):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_BOLD
            cell.alignment = _ALIGN_CENTER

    # --- Item Rows (Rows 15+) ---
    start_row = 15
    min_rows = max(4, len(computed_days))
    for idx in range(min_rows):
        r = start_row + idx
        ws.row_dimensions[r].height = 24
        if idx < len(computed_days):
            cd = computed_days[idx]
            ws[f"A{r}"] = cd.date
            ws[f"B{r}"] = cd.route_text
            ws[f"C{r}"] = cd.summary
            ws[f"D{r}"] = cd.train_fare or None
            ws[f"E{r}"] = cd.flight_hsr_fare or None
            ws[f"F{r}"] = cd.bus_mrt_fare or None
            ws[f"G{r}"] = cd.ship_fare or None
            ws[f"H{r}"] = cd.ubike_fare or None
            ws[f"I{r}"] = cd.self_car_fare or None
            ws[f"J{r}"] = cd.self_scooter_fare or None
            ws[f"K{r}"] = cd.lodging_fee or None
            ws[f"L{r}"] = cd.misc_fee or None
            ws[f"M{r}"] = cd.receipt_no or (cd.receipt_count or None)
            ws[f"N{r}"] = f"=SUM(D{r}:L{r})"
            ws[f"O{r}"] = cd.note
        else:
            ws[f"N{r}"] = f"=SUM(D{r}:L{r})"

        for c in range(1, 16):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_SMALL if c in (2, 3, 15) else _FONT_BODY
            if (4 <= c <= 12) or c == 14:
                cell.alignment = _ALIGN_NUM
                cell.number_format = "#,##0"
            elif c in (2, 3, 15):
                cell.alignment = _ALIGN_LEFT
            else:
                cell.alignment = _ALIGN_CENTER

    end_row = start_row + min_rows - 1
    total_row = end_row + 1
    ws.row_dimensions[total_row].height = 22

    ws.merge_cells(f"A{total_row}:C{total_row}")
    ws[f"A{total_row}"] = "合        計"
    ws[f"A{total_row}"].font = _FONT_BOLD
    ws[f"A{total_row}"].alignment = _ALIGN_CENTER

    for col_idx in list(range(4, 13)) + [14]:
        col_l = get_column_letter(col_idx)
        cell = ws[f"{col_l}{total_row}"]
        cell.value = f"=SUM({col_l}{start_row}:{col_l}{end_row})"
        cell.font = _FONT_BOLD
        cell.alignment = _ALIGN_NUM
        cell.number_format = "#,##0"

    for c in range(1, 16):
        ws.cell(total_row, c).border = _BORDER_ALL

    # Grand total uppercase row
    upper_row = total_row + 1
    ws.row_dimensions[upper_row].height = 20
    ws.merge_cells(f"A{upper_row}:O{upper_row}")
    ws[f"A{upper_row}"] = f"總計（大寫）：{amount_to_chinese_upper(grand_total)}（NT$ {grand_total:,}）"
    ws[f"A{upper_row}"].font = _FONT_BOLD
    ws[f"A{upper_row}"].alignment = _ALIGN_LEFT
    for c in range(1, 16):
        ws.cell(upper_row, c).border = _BORDER_ALL

    # --- Approval Signature Block (2 rows matching official 附表二) ---
    sig_r1 = upper_row + 1
    sig_r2 = upper_row + 2
    ws.row_dimensions[sig_r1].height = 28
    ws.row_dimensions[sig_r2].height = 28

    ws.merge_cells(f"A{sig_r1}:C{sig_r1}")
    ws[f"A{sig_r1}"] = "出差人："
    ws.merge_cells(f"D{sig_r1}:F{sig_r1}")
    ws[f"D{sig_r1}"] = "科長："
    ws.merge_cells(f"G{sig_r1}:K{sig_r1}")
    ws[f"G{sig_r1}"] = "單位主管："
    ws.merge_cells(f"L{sig_r1}:O{sig_r2}")
    ws[f"L{sig_r1}"] = "機關首長（校長）："

    ws.merge_cells(f"A{sig_r2}:C{sig_r2}")
    ws[f"A{sig_r2}"] = "人事單位（出差日期及職級核符）："
    ws.merge_cells(f"D{sig_r2}:F{sig_r2}")
    ws[f"D{sig_r2}"] = "主計承辦人："
    ws.merge_cells(f"G{sig_r2}:K{sig_r2}")
    ws[f"G{sig_r2}"] = "主計單位主管："

    for r in (sig_r1, sig_r2):
        for c in range(1, 16):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_SMALL
            cell.alignment = _ALIGN_LEFT

    # Single-page A4 landscape print setup
    ws.print_area = f"A1:O{sig_r2}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
