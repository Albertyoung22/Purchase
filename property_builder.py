"""Property & Non-Consumable Asset Increase Form Builder.

Implements Tainan City / Taiwan public school asset registration forms:
- Automatically classifies unit_price >= 10,000 as 財產 (Property) and < 10,000 as 非消耗品 (Non-consumable).
- Generates A4 Landscape Excel (.xlsx) with separate sheets for 財產增加單 and 非消耗品登記清冊,
  formulas, Chinese uppercase totals, and standard school sign-off blocks.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from travel_form_builder import amount_to_chinese_upper

PROPERTY_PRICE_THRESHOLD = 10000


@dataclass
class PropertyItem:
    name: str
    spec: str = ""
    unit: str = "台"
    quantity: float = 1.0
    unit_price: int = 0
    category: str = ""  # 財產 | 非消耗品 (auto if blank)
    property_no: str = ""
    useful_life: int = 0  # years (auto 5 for 財產, 2 for 非消耗品 if <= 0)
    location: str = ""
    custodian: str = ""
    invoice_no: str = ""
    note: str = ""


@dataclass
class PropertyIncreaseForm:
    school_name: str = "臺南市立○○國民小學"
    doc_no: str = ""
    acquire_date: str = ""
    department: str = "總務處"
    budget_code: str = ""
    plan_name: str = ""
    items: list[PropertyItem] = field(default_factory=list)


@dataclass
class PropertyItemComputed:
    name: str
    spec: str
    unit: str
    quantity: float
    unit_price: int
    total_price: int
    category: str
    property_no: str
    useful_life: int
    location: str
    custodian: str
    invoice_no: str
    note: str


def compute_property_item(item: PropertyItem, default_location: str = "") -> PropertyItemComputed:
    qty = max(0.0, float(item.quantity or 0.0))
    price = max(0, int(round(float(item.unit_price or 0))))
    total_price = round(qty * price)
    cat = item.category.strip()
    if cat not in {"財產", "非消耗品"}:
        cat = "財產" if price >= PROPERTY_PRICE_THRESHOLD else "非消耗品"
    life = int(item.useful_life or 0)
    if life <= 0:
        life = 5 if cat == "財產" else 2
    return PropertyItemComputed(
        name=item.name.strip(),
        spec=item.spec.strip(),
        unit=item.unit.strip() or "個",
        quantity=qty,
        unit_price=price,
        total_price=total_price,
        category=cat,
        property_no=item.property_no.strip(),
        useful_life=life,
        location=item.location.strip() or default_location,
        custodian=item.custodian.strip(),
        invoice_no=item.invoice_no.strip(),
        note=item.note.strip(),
    )


def summarize_property_form(form: PropertyIncreaseForm) -> tuple[list[PropertyItemComputed], dict[str, int]]:
    computed = [compute_property_item(it, default_location=form.department) for it in form.items]
    prop_items = [c for c in computed if c.category == "財產"]
    non_items = [c for c in computed if c.category == "非消耗品"]
    summary = {
        "total_count": len(computed),
        "property_count": len(prop_items),
        "property_amount": sum(c.total_price for c in prop_items),
        "non_consumable_count": len(non_items),
        "non_consumable_amount": sum(c.total_price for c in non_items),
        "grand_total": sum(c.total_price for c in computed),
    }
    return computed, summary


def _write_asset_sheet(ws, title_label: str, form: PropertyIncreaseForm, items: list[PropertyItemComputed]) -> None:
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2

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

    ws.merge_cells("A1:L1")
    ws["A1"] = f"{form.school_name} {title_label}"
    ws["A1"].font = font_title
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:L2")
    meta_parts = [f"使用單位：{form.department or '-'}"]
    if form.doc_no:
        meta_parts.append(f"單號：{form.doc_no}")
    if form.acquire_date:
        meta_parts.append(f"購置／驗收日期：{form.acquire_date}")
    if form.budget_code:
        meta_parts.append(f"預算科目：{form.budget_code}")
    if form.plan_name:
        meta_parts.append(f"計畫：{form.plan_name}")
    ws["A2"] = "　｜　".join(meta_parts)
    ws["A2"].font = font_sub
    ws["A2"].alignment = left
    ws.row_dimensions[2].height = 20

    headers = [
        "序號", "財產／物品編號", "品名", "規格／廠牌型號", "單位", "數量",
        "單價", "總價", "使用年限(年)", "保管地點", "保管人", "發票號碼／備註",
    ]
    ws.row_dimensions[3].height = 26
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=col_idx, value=h)
        c.font = font_hdr
        c.fill = header_fill
        c.alignment = center
        c.border = border

    row_num = 4
    for idx, it in enumerate(items, start=1):
        ws.row_dimensions[row_num].height = 24
        note_txt = " ".join(p for p in [it.invoice_no, it.note] if p)
        vals = [
            idx, it.property_no, it.name, it.spec, it.unit,
            it.quantity, it.unit_price, f"=F{row_num}*G{row_num}",
            it.useful_life, it.location, it.custodian, note_txt,
        ]
        for col_idx, val in enumerate(vals, start=1):
            c = ws.cell(row=row_num, column=col_idx, value=val)
            c.font = font_cell
            c.border = border
            if col_idx in (1, 2, 5, 9):
                c.alignment = center
            elif col_idx in (6, 7, 8):
                c.alignment = right
                c.number_format = "#,##0"
            else:
                c.alignment = left
        row_num += 1

    ws.row_dimensions[row_num].height = 24
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=7)
    ws.cell(row=row_num, column=1, value="合　計").alignment = center
    for col_idx in range(1, 13):
        c = ws.cell(row=row_num, column=col_idx)
        c.font = font_hdr
        c.fill = total_fill
        c.border = border
        if col_idx == 8:
            c.value = f"=SUM(H4:H{row_num - 1})" if items else 0
            c.alignment = right
            c.number_format = "#,##0"

    total_val = sum(it.total_price for it in items)
    upper_row = row_num + 1
    ws.row_dimensions[upper_row].height = 24
    ws.merge_cells(start_row=upper_row, start_column=1, end_row=upper_row, end_column=12)
    c_up = ws.cell(
        row=upper_row, column=1,
        value=f"總計金額（大寫）：{amount_to_chinese_upper(total_val)}　（共 {len(items)} 筆）",
    )
    c_up.font = font_hdr
    c_up.alignment = left
    for col_idx in range(1, 13):
        ws.cell(row=upper_row, column=col_idx).border = border

    sign_row = upper_row + 2
    ws.row_dimensions[sign_row].height = 24
    ws.merge_cells(start_row=sign_row, start_column=1, end_row=sign_row, end_column=12)
    ws.cell(
        row=sign_row, column=1,
        value="經手人（採購）：　　　　　財產管理員：　　　　　保管單位主管：　　　　　主（會）計室：　　　　　校長：",
    ).font = font_hdr

    widths = [6, 15, 22, 18, 6, 7, 10, 11, 11, 13, 10, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_property_increase_xlsx(form: PropertyIncreaseForm, output_path: Path) -> Path:
    """Build Excel workbook containing 財產增加單 and/or 非消耗品登記清冊."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    computed, _ = summarize_property_form(form)
    prop_items = [c for c in computed if c.category == "財產"]
    non_items = [c for c in computed if c.category == "非消耗品"]

    wb = Workbook()
    first = True
    if prop_items:
        ws1 = wb.active
        ws1.title = "財產增加單"
        _write_asset_sheet(ws1, "財產增加單", form, prop_items)
        first = False
    if non_items:
        ws2 = wb.active if first else wb.create_sheet("非消耗品登記清冊")
        ws2.title = "非消耗品登記清冊"
        _write_asset_sheet(ws2, "非消耗品登記清冊", form, non_items)
        first = False
    if first:
        ws1 = wb.active
        ws1.title = "財產增加單"
        _write_asset_sheet(ws1, "財產增加單", form, [])

    wb.save(output_path)
    return output_path
