"""Substitute Teacher Labor/Health Insurance & Hourly Pay Roster Builder.

Implements Tainan City / Taiwan public school regulations for:
- Hourly pay rates (114/09/01+): Elementary 405, Junior High 455, Senior High 505.
- 115-year (2026) & 114-year (2025) Labor Insurance, Occupational Accident,
  Health Insurance (NHI), Labor Pension (6% + voluntary), and 2nd-Gen NHI
  supplemental premium (2.11% when >= 5,000 and not enrolled at the school).
- Pro-rated 30-day calculation for partial-month enrollment and dual-sheet Excel
  output (印領清冊 + 勞健保勞退機關負擔明細).
"""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from travel_form_builder import amount_to_chinese_upper

SCHOOL_LEVEL_RATES: dict[str, int] = {
    "elementary": 405,
    "junior": 455,
    "senior": 505,
}

SCHOOL_LEVEL_LABELS: dict[str, str] = {
    "elementary": "國小 (405元/節)",
    "junior": "國中 (455元/節)",
    "senior": "高中職 (505元/節)",
    "custom": "自訂單價",
}

# Labor Insurance (incl. Employment Insurance 1% = 12.5%) brackets
# Part-time brackets start at 11,100; full-time floor is 29,500 (115) or 28,590 (114).
_LI_PARTTIME_BRACKETS = [
    11100, 12540, 13500, 15840, 16500, 17280, 17880, 19047, 20008,
    21009, 22000, 23100, 24000, 25250, 26400, 27470, 28590,
]
_LI_UPPER_BRACKETS = [
    29500, 30300, 31800, 33300, 34800, 36300, 38200, 40100, 42000, 43900, 45800,
]
_OCC_EXTRA_BRACKETS = [
    48200, 50600, 53000, 55400, 57800, 60800, 63800, 66800, 69800, 72800,
]
_PENSION_BRACKETS = [
    1500, 3000, 4500, 6000, 7500, 8700, 9900, 11100, 12540, 13500,
    15840, 16500, 17280, 17880, 19047, 20008, 21009, 22000, 23100,
    24000, 25250, 26400, 27470, 28590, 29500, 30300, 31800, 33300,
    34800, 36300, 38200, 40100, 42000, 43900, 45800, 48200, 50600,
    53000, 55400, 57800, 60800, 63800, 66800, 69800, 72800, 76500,
    80200, 83900, 87600, 92100, 96600, 101100, 105600, 110100, 115500,
    120900, 126300, 131700, 137100, 142500, 147900, 150000,
]

LI_RATE = 0.125
LI_EMPLOYEE_SHARE = 0.20
LI_EMPLOYER_SHARE = 0.70

NHI_RATE = 0.0517
NHI_EMPLOYEE_SHARE = 0.30
NHI_EMPLOYER_SHARE = 0.60
NHI_AVG_FAMILY_FACTOR = 1.56  # 1 + 0.56 average dependents
NHI_SUPP_RATE = 0.0211
NHI_SUPP_THRESHOLD = 5000

DEFAULT_OCC_RATE = 0.0013  # Education sector typical occupational accident rate


def _lookup_bracket(amount: int, brackets: list[int], floor: int) -> int:
    target = max(amount, floor)
    for b in brackets:
        if target <= b:
            return b
    return brackets[-1]


def lookup_insurance_brackets(gross_pay: int, year_roc: int = 115, enroll_mode: str = "part_time") -> dict[str, int]:
    """Return insured salary brackets for LI, Occupational, NHI, and Labor Pension."""
    min_wage = 29500 if year_roc >= 115 else 28590
    if enroll_mode == "supp_only":
        return {"li": 0, "occ": 0, "nhi": 0, "pension": 0}

    li_all = [b for b in (_LI_PARTTIME_BRACKETS + _LI_UPPER_BRACKETS) if year_roc >= 115 or b != 29500]
    occ_all = li_all + _OCC_EXTRA_BRACKETS
    li_floor = min_wage if enroll_mode == "full_month" else 11100
    nhi_floor = min_wage

    li_bracket = _lookup_bracket(gross_pay, li_all, li_floor)
    occ_bracket = _lookup_bracket(gross_pay, occ_all, li_floor)
    nhi_bracket = _lookup_bracket(gross_pay, occ_all, nhi_floor) if enroll_mode == "full_month" else 0
    pension_all = [b for b in _PENSION_BRACKETS if year_roc >= 115 or b != 29500]
    pension_bracket = _lookup_bracket(gross_pay, pension_all, 1500)

    return {
        "li": li_bracket,
        "occ": occ_bracket,
        "nhi": nhi_bracket,
        "pension": pension_bracket,
    }


@dataclass
class SubstituteEntry:
    teacher_name: str
    id_number: str = ""
    subject_note: str = ""
    school_level: str = "elementary"  # elementary | junior | senior | custom
    hourly_rate: int = 0
    periods: float = 0.0
    enroll_mode: str = "part_time"  # full_month | part_time | supp_only
    insured_days: int = 30
    dependents: int = 0
    voluntary_pension_pct: float = 0.0
    nhi_end_of_month: bool = True
    nhi_enrolled_here: bool = False
    income_tax: int = 0
    note: str = ""


@dataclass
class SubstituteComputed:
    teacher_name: str
    id_number: str
    subject_note: str
    school_level: str
    hourly_rate: int
    periods: float
    gross_pay: int
    enroll_mode: str
    insured_days: int
    li_bracket: int
    nhi_bracket: int
    pension_bracket: int
    li_employee: int
    li_employer: int
    occ_employer: int
    nhi_employee: int
    nhi_employer: int
    pension_employee: int
    pension_employer: int
    nhi_supp: int
    income_tax: int
    total_deduction: int
    net_pay: int
    employer_total: int
    note: str


@dataclass
class SubstituteRoster:
    school_name: str = "臺南市立○○國民小學"
    roc_year: int = 115
    month: int = 11
    plan_name: str = "教師請假短期代課鐘點費"
    budget_code: str = ""
    occ_rate: float = DEFAULT_OCC_RATE
    entries: list[SubstituteEntry] = field(default_factory=list)


def compute_substitute_entry(entry: SubstituteEntry, roc_year: int = 115, occ_rate: float = DEFAULT_OCC_RATE) -> SubstituteComputed:
    """Compute gross pay, employee deductions, employer contributions, and net pay."""
    rate = entry.hourly_rate if entry.hourly_rate > 0 else SCHOOL_LEVEL_RATES.get(entry.school_level, 405)
    periods = max(0.0, float(entry.periods or 0.0))
    gross_pay = round(rate * periods)

    mode = entry.enroll_mode if entry.enroll_mode in {"full_month", "part_time", "supp_only"} else "part_time"
    days = max(0, min(30, int(entry.insured_days if entry.insured_days is not None else 30)))
    day_ratio = days / 30.0 if mode != "supp_only" else 0.0

    brackets = lookup_insurance_brackets(gross_pay, year_roc=roc_year, enroll_mode=mode)
    li_b = brackets["li"]
    occ_b = brackets["occ"]
    nhi_b = brackets["nhi"] if (mode == "full_month" or entry.nhi_enrolled_here) else 0
    if entry.nhi_enrolled_here and nhi_b == 0 and mode != "supp_only":
        min_wage = 29500 if roc_year >= 115 else 28590
        nhi_b = max(gross_pay, min_wage)
    pen_b = brackets["pension"]

    li_emp = round(li_b * LI_RATE * LI_EMPLOYEE_SHARE * day_ratio) if li_b else 0
    li_org = round(li_b * LI_RATE * LI_EMPLOYER_SHARE * day_ratio) if li_b else 0
    occ_org = round(occ_b * occ_rate * day_ratio) if occ_b else 0

    if nhi_b > 0 and entry.nhi_end_of_month:
        capped_deps = max(0, min(3, int(entry.dependents or 0)))
        nhi_emp = round(nhi_b * NHI_RATE * NHI_EMPLOYEE_SHARE * (1 + capped_deps))
        nhi_org = round(nhi_b * NHI_RATE * NHI_EMPLOYER_SHARE * NHI_AVG_FAMILY_FACTOR)
    else:
        nhi_emp = 0
        nhi_org = 0

    vol_pct = max(0.0, min(6.0, float(entry.voluntary_pension_pct or 0.0))) / 100.0
    pen_emp = round(pen_b * vol_pct * day_ratio) if pen_b else 0
    pen_org = round(pen_b * 0.06 * day_ratio) if pen_b else 0

    # 2nd-gen NHI supplemental premium when not enrolled in NHI at this school and gross >= 5,000
    nhi_supp = 0
    if nhi_emp == 0 and not (mode == "full_month" and entry.nhi_end_of_month) and not entry.nhi_enrolled_here:
        if gross_pay >= NHI_SUPP_THRESHOLD:
            nhi_supp = round(gross_pay * NHI_SUPP_RATE)

    tax = max(0, int(entry.income_tax or 0))
    total_deduction = li_emp + nhi_emp + pen_emp + nhi_supp + tax
    net_pay = gross_pay - total_deduction
    employer_total = li_org + occ_org + nhi_org + pen_org

    notes: list[str] = []
    if entry.note:
        notes.append(entry.note)
    if nhi_supp > 0:
        notes.append(f"達{NHI_SUPP_THRESHOLD}元代扣二代健保2.11%")
    if mode != "supp_only" and days < 30:
        notes.append(f"破月投保{days}天")

    return SubstituteComputed(
        teacher_name=entry.teacher_name.strip(),
        id_number=entry.id_number.strip(),
        subject_note=entry.subject_note.strip(),
        school_level=entry.school_level,
        hourly_rate=rate,
        periods=periods,
        gross_pay=gross_pay,
        enroll_mode=mode,
        insured_days=days,
        li_bracket=li_b,
        nhi_bracket=nhi_b,
        pension_bracket=pen_b,
        li_employee=li_emp,
        li_employer=li_org,
        occ_employer=occ_org,
        nhi_employee=nhi_emp,
        nhi_employer=nhi_org,
        pension_employee=pen_emp,
        pension_employer=pen_org,
        nhi_supp=nhi_supp,
        income_tax=tax,
        total_deduction=total_deduction,
        net_pay=net_pay,
        employer_total=employer_total,
        note="；".join(notes),
    )


def summarize_roster(roster: SubstituteRoster) -> tuple[list[SubstituteComputed], dict[str, int]]:
    items = [compute_substitute_entry(e, roc_year=roster.roc_year, occ_rate=roster.occ_rate) for e in roster.entries]
    totals = {
        "gross_pay": sum(i.gross_pay for i in items),
        "li_employee": sum(i.li_employee for i in items),
        "nhi_employee": sum(i.nhi_employee for i in items),
        "pension_employee": sum(i.pension_employee for i in items),
        "nhi_supp": sum(i.nhi_supp for i in items),
        "income_tax": sum(i.income_tax for i in items),
        "total_deduction": sum(i.total_deduction for i in items),
        "net_pay": sum(i.net_pay for i in items),
        "li_employer": sum(i.li_employer for i in items),
        "occ_employer": sum(i.occ_employer for i in items),
        "nhi_employer": sum(i.nhi_employer for i in items),
        "pension_employer": sum(i.pension_employer for i in items),
        "employer_total": sum(i.employer_total for i in items),
        "school_grand_total": sum(i.gross_pay + i.employer_total for i in items),
    }
    return items, totals


def build_substitute_roster_xlsx(roster: SubstituteRoster, output_path: Path) -> Path:
    """Build a 2-sheet Tainan school substitute teacher payroll & insurance workbook."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    items, totals = summarize_roster(roster)

    wb = Workbook()
    thin = Side(style="thin", color="334155")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="E2E8F0")
    total_fill = PatternFill("solid", fgColor="F1F5F9")
    font_title = Font(name="標楷體", size=15, bold=True)
    font_sub = Font(name="標楷體", size=10)
    font_hdr = Font(name="標楷體", size=10, bold=True)
    font_cell = Font(name="標楷體", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Sheet 1: 鐘點費印領清冊
    ws1 = wb.active
    ws1.title = "鐘點費印領清冊"
    ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.page_margins.left = 0.25
    ws1.page_margins.right = 0.25
    ws1.page_margins.top = 0.4
    ws1.page_margins.bottom = 0.4
    ws1.page_margins.header = 0.2
    ws1.page_margins.footer = 0.2

    ws1.merge_cells("A1:O1")
    ws1["A1"] = f"{roster.school_name} {roster.roc_year}年{roster.month}月 代課教師鐘點費印領清冊"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = center
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:O2")
    budget_txt = f"　預算科目：{roster.budget_code}" if roster.budget_code else ""
    ws1["A2"] = f"案別／計畫：{roster.plan_name}{budget_txt}　（金額單位：新臺幣元）"
    ws1["A2"].font = font_sub
    ws1["A2"].alignment = left
    ws1.row_dimensions[2].height = 20

    headers1 = [
        "序號", "教師姓名", "身分證字號", "代課科目／說明", "節數", "單價",
        "應領金額", "勞保自付", "健保自付", "勞退自提", "二代健保", "所得稅",
        "實領金額", "領款人簽章", "備註",
    ]
    ws1.row_dimensions[3].height = 26
    for col_idx, h in enumerate(headers1, start=1):
        c = ws1.cell(row=3, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = header_fill
        c.alignment = center
        c.border = border

    row_num = 4
    for idx, item in enumerate(items, start=1):
        ws1.row_dimensions[row_num].height = 24
        vals = [
            idx, item.teacher_name, item.id_number, item.subject_note,
            item.periods, item.hourly_rate,
            f"=E{row_num}*F{row_num}",
            item.li_employee, item.nhi_employee, item.pension_employee,
            item.nhi_supp, item.income_tax,
            f"=G{row_num}-SUM(H{row_num}:L{row_num})",
            "", item.note,
        ]
        for col_idx, val in enumerate(vals, start=1):
            c = ws1.cell(row=row_num, column=col_idx, value=val)
            c.font = font_cell
            c.border = border
            if col_idx in (1, 3):
                c.alignment = center
            elif col_idx in (2, 4, 15):
                c.alignment = left
            else:
                c.alignment = right
                if col_idx >= 6 and col_idx != 14:
                    c.number_format = "#,##0"
        row_num += 1

    # Total row
    ws1.row_dimensions[row_num].height = 26
    ws1.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=6)
    ws1.cell(row=row_num, column=1, value="合　計").alignment = center
    for col_idx in range(1, 16):
        c = ws1.cell(row=row_num, column=col_idx)
        c.font = font_hdr
        c.fill = total_fill
        c.border = border
        if 7 <= col_idx <= 13:
            col_ltr = get_column_letter(col_idx)
            c.value = f"=SUM({col_ltr}4:{col_ltr}{row_num - 1})" if items else 0
            c.alignment = right
            c.number_format = "#,##0"

    # Chinese uppercase total
    upper_row = row_num + 1
    ws1.row_dimensions[upper_row].height = 24
    ws1.merge_cells(start_row=upper_row, start_column=1, end_row=upper_row, end_column=15)
    c_up = ws1.cell(
        row=upper_row, column=1,
        value=f"應領總計：{amount_to_chinese_upper(totals['gross_pay'])}　｜　實領總計：{amount_to_chinese_upper(totals['net_pay'])}　｜　學校機關負擔總計：新臺幣 {totals['employer_total']:,} 元",
    )
    c_up.font = font_hdr
    c_up.alignment = left
    for col_idx in range(1, 16):
        ws1.cell(row=upper_row, column=col_idx).border = border

    # Approval sign-off row
    sign_row = upper_row + 2
    ws1.row_dimensions[sign_row].height = 24
    ws1.merge_cells(start_row=sign_row, start_column=1, end_row=sign_row, end_column=15)
    ws1.cell(
        row=sign_row, column=1,
        value="經手人：　　　　　　教務處：　　　　　　人事室：　　　　　　出納組：　　　　　　主計室：　　　　　　校長：",
    ).font = font_hdr

    widths1 = [6, 12, 13, 20, 7, 8, 11, 10, 10, 10, 10, 9, 11, 13, 18]
    for i, w in enumerate(widths1, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # Sheet 2: 勞健保勞退機關負擔明細
    ws2 = wb.create_sheet("勞健保勞退機關負擔明細")
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0
    ws2.page_margins.left = 0.25
    ws2.page_margins.right = 0.25
    ws2.page_margins.top = 0.4
    ws2.page_margins.bottom = 0.4
    ws2.page_margins.header = 0.2
    ws2.page_margins.footer = 0.2

    ws2.merge_cells("A1:M1")
    ws2["A1"] = f"{roster.school_name} {roster.roc_year}年{roster.month}月 代課教師勞健保及勞退機關負擔明細表"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = center
    ws2.row_dimensions[1].height = 28

    headers2 = [
        "序號", "教師姓名", "投保模式", "投保天數", "勞保級距", "健保級距", "勞退級距",
        "勞保機關(70%)", "職災機關", "健保機關", "勞退機關(6%)", "機關負擔合計", "人事費總成本(應領+機關)",
    ]
    ws2.row_dimensions[3].height = 26
    for col_idx, h in enumerate(headers2, start=1):
        c = ws2.cell(row=3, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = header_fill
        c.alignment = center
        c.border = border

    mode_labels = {"full_month": "全月加保", "part_time": "部分工時", "supp_only": "僅二代健保"}
    r2 = 4
    for idx, item in enumerate(items, start=1):
        ws2.row_dimensions[r2].height = 22
        vals = [
            idx, item.teacher_name, mode_labels.get(item.enroll_mode, item.enroll_mode),
            item.insured_days, item.li_bracket, item.nhi_bracket, item.pension_bracket,
            item.li_employer, item.occ_employer, item.nhi_employer, item.pension_employer,
            f"=SUM(H{r2}:K{r2})",
            f"='鐘點費印領清冊'!G{r2}+L{r2}",
        ]
        for col_idx, val in enumerate(vals, start=1):
            c = ws2.cell(row=r2, column=col_idx, value=val)
            c.font = font_cell
            c.border = border
            if col_idx in (1, 3, 4):
                c.alignment = center
            elif col_idx == 2:
                c.alignment = left
            else:
                c.alignment = right
                c.number_format = "#,##0"
        r2 += 1

    ws2.row_dimensions[r2].height = 24
    ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=7)
    ws2.cell(row=r2, column=1, value="合　計").alignment = center
    for col_idx in range(1, 14):
        c = ws2.cell(row=r2, column=col_idx)
        c.font = font_hdr
        c.fill = total_fill
        c.border = border
        if col_idx >= 8:
            col_ltr = get_column_letter(col_idx)
            c.value = f"=SUM({col_ltr}4:{col_ltr}{r2 - 1})" if items else 0
            c.alignment = right
            c.number_format = "#,##0"

    widths2 = [6, 12, 12, 10, 11, 11, 11, 14, 11, 12, 14, 14, 22]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_path)
    return output_path


_SUB_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "teacher_name": ("教師姓名", "姓名", "代課教師", "老師姓名", "name", "teacher_name"),
    "id_number": ("身分證", "身分證字號", "id", "id_number"),
    "subject_note": ("代課科目", "科目", "說明", "班級科目", "授課科目", "subject_note"),
    "school_level": ("學制", "教育階段", "階段", "school_level"),
    "hourly_rate": ("單價", "鐘點費", "鐘點單價", "時薪", "hourly_rate"),
    "periods": ("節數", "授課節數", "時數", "總節數", "periods"),
    "enroll_mode": ("投保模式", "投保別", "加保別", "加保方式", "enroll_mode"),
    "insured_days": ("天數", "投保天數", "加保天數", "insured_days"),
    "dependents": ("眷口", "眷口數", "dependents"),
    "voluntary_pension_pct": ("自提", "勞退自提", "自提%", "voluntary_pension_pct"),
    "nhi_end_of_month": ("月底健保", "健保", "nhi_end_of_month"),
    "income_tax": ("所得稅", "代扣所得稅", "income_tax"),
    "note": ("備註", "note"),
}


def _norm_level(val: str, rate: int = 0) -> str:
    v = val.strip().lower()
    if "國中" in v or v == "junior":
        return "junior"
    if "高中" in v or "高職" in v or v == "senior":
        return "senior"
    if "國小" in v or v == "elementary":
        return "elementary"
    if rate == 455:
        return "junior"
    if rate == 505:
        return "senior"
    if rate > 0 and rate != 405:
        return "custom"
    return "elementary"


def _norm_enroll_mode(val: str) -> str:
    v = val.strip().lower()
    if "全月" in v or "長期" in v or v == "full_month":
        return "full_month"
    if "二代" in v or "外聘" in v or "免投保" in v or v == "supp_only":
        return "supp_only"
    return "part_time"


def _to_float(val: Any, default: float = 0.0) -> float:
    if val is None or val == "":
        return default
    m = re.search(r"-?\d+(?:\.\d+)?", str(val).replace(",", ""))
    return float(m.group(0)) if m else default


def _map_sub_row(row_dict: dict[str, Any]) -> dict[str, Any] | None:
    name = str(row_dict.get("teacher_name") or "").strip()
    if not name or name in {"合計", "總計", "序號"}:
        return None
    rate = int(round(_to_float(row_dict.get("hourly_rate"), 0)))
    lvl = _norm_level(str(row_dict.get("school_level") or ""), rate=rate)
    if lvl != "custom" and rate in (0, 405, 455, 505):
        rate_out = "" if rate in (0, SCHOOL_LEVEL_RATES.get(lvl)) else rate
    else:
        rate_out = rate if rate > 0 else ""
    eom_raw = str(row_dict.get("nhi_end_of_month") or "").strip().lower()
    eom = eom_raw not in {"0", "false", "no", "否", "無", "n"}
    return {
        "teacher_name": name,
        "id_number": str(row_dict.get("id_number") or "").strip(),
        "subject_note": str(row_dict.get("subject_note") or "").strip(),
        "school_level": lvl,
        "hourly_rate": rate_out,
        "periods": _to_float(row_dict.get("periods"), 1),
        "enroll_mode": _norm_enroll_mode(str(row_dict.get("enroll_mode") or "")),
        "insured_days": int(round(_to_float(row_dict.get("insured_days"), 30))),
        "dependents": int(round(_to_float(row_dict.get("dependents"), 0))),
        "voluntary_pension_pct": _to_float(row_dict.get("voluntary_pension_pct"), 0),
        "nhi_end_of_month": eom,
        "income_tax": int(round(_to_float(row_dict.get("income_tax"), 0))),
        "note": str(row_dict.get("note") or "").strip(),
    }


def _match_header_col(cell_text: str) -> str | None:
    txt = re.sub(r"\s+", "", str(cell_text or "")).lower()
    if not txt:
        return None
    for field_name, aliases in _SUB_HEADER_ALIASES.items():
        if any(a.lower() in txt for a in aliases):
            return field_name
    return None


def parse_substitute_csv_or_xlsx(file_bytes: bytes, filename: str) -> list[dict[str, Any]]:
    """Parse substitute teacher roster from .xlsx or .csv with flexible Chinese headers."""
    ext = Path(filename).suffix.lower()
    rows: list[list[Any]] = []
    if ext in {".xlsx", ".xlsm"}:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        for r in ws.iter_rows(values_only=True):
            if any(c is not None and str(c).strip() for c in r):
                rows.append(list(r))
    else:
        text = None
        for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
            try:
                text = file_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = file_bytes.decode("utf-8", errors="ignore")
        reader = csv.reader(io.StringIO(text))
        for r in reader:
            if any(str(c).strip() for c in r):
                rows.append(r)

    if not rows:
        return []

    # Locate header row within first 10 rows
    header_idx = 0
    col_map: dict[int, str] = {}
    for r_idx, r in enumerate(rows[:10]):
        candidate: dict[int, str] = {}
        for c_idx, val in enumerate(r):
            matched = _match_header_col(str(val or ""))
            if matched and matched not in candidate.values():
                candidate[c_idx] = matched
        if "teacher_name" in candidate.values():
            header_idx = r_idx
            col_map = candidate
            break

    if not col_map:
        # Fallback positional order: name, subject, level, periods, mode, days
        col_map = {0: "teacher_name", 1: "subject_note", 2: "school_level", 3: "periods", 4: "enroll_mode", 5: "insured_days"}
        data_rows = rows
    else:
        data_rows = rows[header_idx + 1:]

    entries: list[dict[str, Any]] = []
    for r in data_rows:
        raw_dict: dict[str, Any] = {}
        for c_idx, field_name in col_map.items():
            if c_idx < len(r):
                raw_dict[field_name] = r[c_idx]
        mapped = _map_sub_row(raw_dict)
        if mapped:
            entries.append(mapped)
    return entries


def parse_substitute_text_smart(raw_text: str) -> list[dict[str, Any]]:
    """Smart parser for pasted Excel/LINE text or OCR/PDF lines of substitute teachers."""
    lines = [ln.strip() for ln in (raw_text or "").splitlines() if ln.strip()]
    if not lines:
        return []

    # Check if Tab/comma separated with header
    if any("\t" in ln or "," in ln for ln in lines[:3]):
        sep = "\t" if any("\t" in ln for ln in lines[:3]) else ","
        buf = io.StringIO()
        writer = csv.writer(buf)
        for ln in lines:
            writer.writerow([p.strip() for p in ln.split(sep)])
        parsed = parse_substitute_csv_or_xlsx(buf.getvalue().encode("utf-8"), "paste.csv")
        if parsed:
            return parsed

    entries: list[dict[str, Any]] = []
    for ln in lines:
        if any(skip in ln for skip in ("教師姓名", "序號", "合計", "總計")):
            continue
        # Extract explicit patterns: e.g. "1. 王小明 國小 16節 五年甲班國語 部分工時 30天"
        cleaned = re.sub(r"^\d+[\.、\)\s]+", "", ln).strip()
        if not cleaned:
            continue

        lvl = _norm_level(cleaned)
        mode = _norm_enroll_mode(cleaned)
        days_m = re.search(r"(\d{1,2})\s*[天日]", cleaned)
        days = int(days_m.group(1)) if days_m else 30
        per_m = re.search(r"(\d+(?:\.\d+)?)\s*(?:節|堂|小時|hr)", cleaned, re.IGNORECASE)
        rate_m = re.search(r"(\d{3,4})\s*元", cleaned)

        tokens = re.split(r"[\s,，、/]+", cleaned)
        name = ""
        subject_parts: list[str] = []
        periods = float(per_m.group(1)) if per_m else 0.0
        rate = int(rate_m.group(1)) if rate_m else 0

        for tok in tokens:
            if not tok:
                continue
            if not name and re.fullmatch(r"[\u4e00-\u9fffA-Za-z]{2,5}", tok) and not any(
                k in tok for k in ("國小", "國中", "高中", "全月", "部分工時", "二代健保")
            ):
                name = tok
                continue
            if periods == 0.0 and re.fullmatch(r"\d+(?:\.\d+)?", tok):
                num = float(tok)
                if num <= 200:
                    periods = num
                    continue
            if any(k in tok for k in ("國小", "國中", "高中", "高職", "全月", "部分工時", "二代健保")) or re.fullmatch(r"\d+(?:\.\d+)?[節堂天日元]", tok):
                continue
            subject_parts.append(tok)

        if name:
            entries.append({
                "teacher_name": name,
                "id_number": "",
                "subject_note": " ".join(subject_parts),
                "school_level": lvl,
                "hourly_rate": rate if (lvl == "custom" and rate > 0) else "",
                "periods": periods if periods > 0 else 1,
                "enroll_mode": mode,
                "insured_days": max(1, min(30, days)),
                "dependents": 0,
                "voluntary_pension_pct": 0,
                "nhi_end_of_month": True,
                "income_tax": 0,
                "note": "",
            })
    return entries


SUBSTITUTE_SAMPLE_CSV = (
    "教師姓名,身分證字號,代課科目,學制,單價,節數,投保模式,投保天數,眷口數,勞退自提,所得稅,備註\n"
    "王小明,A123456789,五年甲班國語,國小,405,16,部分工時,30,0,0,0,\n"
    "李美華,B234567890,七年乙班數學,國中,455,40,全月加保,30,1,6,0,\n"
    "陳大文,,社團外聘講師,自訂單價,600,10,僅二代健保,0,0,0,0,\n"
)
