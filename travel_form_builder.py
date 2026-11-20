"""Standalone builder for Tainan City Schedule 2 (臺南市政府附表二) Domestic Business Travel Expense Report."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


@dataclass
class TravelRateConfig:
    """Configurable Tainan City domestic business travel reimbursement rates."""

    car_rate_per_km: float = 3.25  # 104.10.16 教育局公告第 78184 號 (興南客運基本票價)
    lodging_limit_standard: int = 2000  # 簡任級以下每日上限
    lodging_limit_special: int = 2400  # 特任級每日上限
    misc_far_km: float = 60.0  # 單程 60 公里以上
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
    transport_mode: str = "self_car"  # self_car | train | flight_hsr | bus_mrt | ship | taxi | official_free
    segments: list[RouteSegment] = field(default_factory=list)
    one_way_km: float | None = None
    bus_equivalent_fare: int | float | None = None
    train_fare: int | float = 0
    flight_hsr_fare: int | float = 0
    bus_mrt_fare: int | float = 0
    ship_fare: int | float = 0
    taxi_approved: bool = False
    lodging_actual: int | float = 0
    misc_fee: int | float | None = None
    receipt_count: int = 0
    note: str = ""


@dataclass
class TravelReport:
    """Full domestic business travel expense report (臺南市政府附表二)."""

    school_name: str = "臺南市立土城高級中學"
    traveler_name: str = ""
    job_title: str = ""
    rank_level: str = "簡任級以下"
    purpose: str = ""
    destination: str = ""
    start_date: str = ""
    end_date: str = ""
    total_days_text: str = ""
    budget_code: str = ""
    plan_name: str = ""
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
    lodging_fee: int
    misc_fee: int
    receipt_count: int
    note: str

    @property
    def subtotal(self) -> int:
        return (
            self.train_fare
            + self.flight_hsr_fare
            + self.bus_mrt_fare
            + self.ship_fare
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


def compute_day_amounts(entry: TravelDayEntry, *, rank_level: str, rates: TravelRateConfig) -> ComputedDay:
    """Apply Tainan City travel rules to compute transport, lodging, and misc fees for one day."""
    notes: list[str] = [entry.note.strip()] if entry.note.strip() else []
    route_text = entry.route_text.strip() or _format_route_from_segments(entry.segments)

    train = int(round(entry.train_fare))
    flight_hsr = int(round(entry.flight_hsr_fare))
    bus_mrt = int(round(entry.bus_mrt_fare))
    ship = int(round(entry.ship_fare))

    if entry.transport_mode == "official_free":
        train = flight_hsr = bus_mrt = ship = 0
        notes.append("機關專備交通工具/免費票，交通費0元")
    elif entry.transport_mode == "taxi":
        if not entry.taxi_approved:
            raise ValueError("依報支要點第5點規定，搭乘計程車須因業務需要經機關核准方可報支")
    elif entry.transport_mode == "self_car":
        if entry.bus_equivalent_fare is not None:
            bus_mrt = int(round(entry.bus_equivalent_fare))
            notes.append("自用汽機車按同路段公民營客運票價報支")
        elif entry.segments:
            total_km = sum(max(0.0, float(s.distance_km)) for s in entry.segments)
            bus_mrt = int(round(total_km * rates.car_rate_per_km))
            notes.append(f"自用車({total_km:g}km×{rates.car_rate_per_km}元)")
    elif entry.transport_mode == "flight_hsr" and flight_hsr > 0:
        notes.append("需檢附票根或購票證明")

    # Lodging cap
    lodging_limit = rates.lodging_limit_special if rank_level == "特任級" else rates.lodging_limit_standard
    raw_lodging = max(0, int(round(entry.lodging_actual)))
    if raw_lodging > lodging_limit:
        lodging_fee = lodging_limit
        notes.append(f"住宿費依每日上限{lodging_limit}元報支")
    else:
        lodging_fee = raw_lodging

    # Misc fee tiered by one-way distance
    if entry.misc_fee is not None:
        misc_fee = max(0, int(round(entry.misc_fee)))
    else:
        one_way = _infer_one_way_km(entry)
        if one_way >= rates.misc_far_km:
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
        lodging_fee=lodging_fee,
        misc_fee=misc_fee,
        receipt_count=entry.receipt_count,
        note="；".join(n for n in notes if n),
    )


_THIN = Side(style="thin", color="000000")
_BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FONT_TITLE = Font(name="標楷體", size=14, bold=True)
_FONT_SUB = Font(name="標楷體", size=12, bold=True)
_FONT_BOLD = Font(name="標楷體", size=10, bold=True)
_FONT_BODY = Font(name="標楷體", size=10)
_FONT_SMALL = Font(name="標楷體", size=9)

_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
_ALIGN_NUM = Alignment(horizontal="right", vertical="center", wrap_text=False, shrinkToFit=True)


def build_travel_report_xlsx(report: TravelReport, output_path: Path) -> Path:
    """Create a print-ready Tainan City Schedule 2 (國內出差旅費報告表) .xlsx file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "國內出差旅費報告表"

    # 12 columns (A..L) tuned for A4 landscape
    col_widths = {
        "A": 9,   # 日期
        "B": 18,  # 起訖地點
        "C": 22,  # 工作摘要
        "D": 9,   # 火車
        "E": 11,  # 飛機及高鐵
        "F": 11,  # 汽車及捷運
        "G": 9,   # 船舶
        "H": 10,  # 住宿費
        "I": 9,   # 雜費
        "J": 9,   # 單據張數
        "K": 11,  # 小計
        "L": 22,  # 備註
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    # Title rows
    ws.merge_cells("A1:L1")
    ws["A1"] = report.school_name
    ws["A1"].font = _FONT_TITLE
    ws["A1"].alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:L2")
    ws["A2"] = "國內出差旅費報告表（臺南市政府附表二）"
    ws["A2"].font = _FONT_SUB
    ws["A2"].alignment = _ALIGN_CENTER
    ws.row_dimensions[2].height = 22

    # Header info block (Rows 3-5)
    ws.merge_cells("A3:B3")
    ws["A3"] = "出差人姓名"
    ws.merge_cells("C3:D3")
    ws["C3"] = report.traveler_name
    ws.merge_cells("E3:F3")
    ws["E3"] = "職稱 / 職等"
    ws.merge_cells("G3:H3")
    ws["G3"] = f"{report.job_title}（{report.rank_level}）".strip("（）")
    ws.merge_cells("I3:J3")
    ws["I3"] = "預算科目 / 計畫"
    ws.merge_cells("K3:L3")
    ws["K3"] = f"{report.budget_code} {report.plan_name}".strip()

    ws.merge_cells("A4:B4")
    ws["A4"] = "出差事由"
    ws.merge_cells("C4:H4")
    ws["C4"] = report.purpose
    ws.merge_cells("I4:J4")
    ws["I4"] = "出差地點"
    ws.merge_cells("K4:L4")
    ws["K4"] = report.destination

    ws.merge_cells("A5:B5")
    ws["A5"] = "出差起訖日期"
    ws.merge_cells("C5:H5")
    date_span = f"{report.start_date} 至 {report.end_date}".strip(" 至")
    if report.total_days_text:
        date_span = f"{date_span}（{report.total_days_text}）" if date_span else report.total_days_text
    ws["C5"] = date_span
    ws.merge_cells("I5:J5")
    ws["I5"] = "附單據總張數"

    computed_days = [
        compute_day_amounts(d, rank_level=report.rank_level, rates=report.rates)
        for d in report.days
    ]

    # Table header (Rows 6-7)
    ws.merge_cells("A6:A7")
    ws["A6"] = "月日"
    ws.merge_cells("B6:B7")
    ws["B6"] = "起訖地點"
    ws.merge_cells("C6:C7")
    ws["C6"] = "工作摘要"
    ws.merge_cells("D6:G6")
    ws["D6"] = "交   通   費"
    ws["D7"] = "火車"
    ws["E7"] = "飛機及高鐵"
    ws["F7"] = "汽車及捷運"
    ws["G7"] = "船舶"
    ws.merge_cells("H6:H7")
    ws["H6"] = "住宿費"
    ws.merge_cells("I6:I7")
    ws["I6"] = "雜費"
    ws.merge_cells("J6:J7")
    ws["J6"] = "單據張數"
    ws.merge_cells("K6:K7")
    ws["K6"] = "小計"
    ws.merge_cells("L6:L7")
    ws["L6"] = "備註"

    for r in range(3, 8):
        ws.row_dimensions[r].height = 21
        for c in range(1, 13):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_BOLD if (r >= 6 or c in (1, 5, 9)) else _FONT_BODY
            cell.alignment = _ALIGN_CENTER

    # Item rows
    start_row = 8
    min_rows = max(6, len(computed_days))
    for idx in range(min_rows):
        r = start_row + idx
        ws.row_dimensions[r].height = 24
        if idx < len(computed_days):
            cd = computed_days[idx]
            ws[f"A{r}"] = cd.date
            ws[f"B{r}"] = cd.route_text
            ws[f"C{r}"] = cd.summary
            ws[f"D{r}"] = cd.train_fare if cd.train_fare else None
            ws[f"E{r}"] = cd.flight_hsr_fare if cd.flight_hsr_fare else None
            ws[f"F{r}"] = cd.bus_mrt_fare if cd.bus_mrt_fare else None
            ws[f"G{r}"] = cd.ship_fare if cd.ship_fare else None
            ws[f"H{r}"] = cd.lodging_fee if cd.lodging_fee else None
            ws[f"I{r}"] = cd.misc_fee if cd.misc_fee else None
            ws[f"J{r}"] = cd.receipt_count if cd.receipt_count else None
            ws[f"K{r}"] = f"=SUM(D{r}:I{r})"
            ws[f"L{r}"] = cd.note
        else:
            ws[f"K{r}"] = f"=SUM(D{r}:I{r})"

        for c in range(1, 13):
            cell = ws.cell(r, c)
            cell.border = _BORDER_ALL
            cell.font = _FONT_SMALL if c in (2, 3, 12) else _FONT_BODY
            if 4 <= c <= 11:
                cell.alignment = _ALIGN_NUM
                cell.number_format = "#,##0"
            elif c in (2, 3, 12):
                cell.alignment = _ALIGN_LEFT
            else:
                cell.alignment = _ALIGN_CENTER

    end_row = start_row + min_rows - 1
    total_row = end_row + 1
    ws.row_dimensions[total_row].height = 24

    ws.merge_cells(f"A{total_row}:C{total_row}")
    ws[f"A{total_row}"] = "合        計"
    ws[f"A{total_row}"].font = _FONT_BOLD
    ws[f"A{total_row}"].alignment = _ALIGN_CENTER

    for col_idx in range(4, 12):
        col_l = get_column_letter(col_idx)
        cell = ws[f"{col_l}{total_row}"]
        cell.value = f"=SUM({col_l}{start_row}:{col_l}{end_row})"
        cell.font = _FONT_BOLD
        cell.alignment = _ALIGN_NUM
        cell.number_format = "#,##0"

    ws[f"L{total_row}"] = ""
    for c in range(1, 13):
        ws.cell(total_row, c).border = _BORDER_ALL

    # Link total receipt count in header K5:L5
    ws.merge_cells("K5:L5")
    ws["K5"] = f"=J{total_row}"
    ws["K5"].alignment = _ALIGN_NUM
    ws["K5"].number_format = "#,##0"

    # Grand total in Chinese uppercase
    grand_total = sum(cd.subtotal for cd in computed_days)
    upper_row = total_row + 1
    ws.row_dimensions[upper_row].height = 24
    ws.merge_cells(f"A{upper_row}:L{upper_row}")
    ws[f"A{upper_row}"] = f"總計（大寫）：{amount_to_chinese_upper(grand_total)}（NT$ {grand_total:,}）"
    ws[f"A{upper_row}"].font = _FONT_BOLD
    ws[f"A{upper_row}"].alignment = _ALIGN_LEFT
    for c in range(1, 13):
        ws.cell(upper_row, c).border = _BORDER_ALL

    # Approval signature block
    sig_row = upper_row + 1
    ws.row_dimensions[sig_row].height = 36
    sig_spans = [("A", "B", "出差人："), ("C", "D", "單位主管："), ("E", "F", "人事室："), ("G", "I", "主計室："), ("J", "L", "校長：")]
    for c1, c2, label in sig_spans:
        ws.merge_cells(f"{c1}{sig_row}:{c2}{sig_row}")
        ws[f"{c1}{sig_row}"] = label
        ws[f"{c1}{sig_row}"].font = _FONT_BOLD
        ws[f"{c1}{sig_row}"].alignment = _ALIGN_LEFT
    for c in range(1, 13):
        ws.cell(sig_row, c).border = _BORDER_ALL

    # Single-page A4 landscape print setup
    ws.print_area = f"A1:L{sig_row}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
