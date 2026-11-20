"""Automated batch filler for school purchase requisition Excel templates."""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from form_splitter import SplitItem, solve_form_splitting
from invoice_parser import ParsedInvoice, parse_einvoice_csv, parse_invoice_text


@dataclass
class LineItem:
    """Single item row on a purchase requisition form."""

    name: str
    spec: str = ""
    unit: str = ""
    quantity: int | float | None = None
    unit_price: int | float | None = None
    manual_subtotal: int | float | None = None
    invoice_no: str = ""
    note: str = ""

    @property
    def subtotal(self) -> int | float | None:
        if self.manual_subtotal is not None:
            return self.manual_subtotal
        if self.quantity is not None and self.unit_price is not None:
            total = self.quantity * self.unit_price
            return int(total) if total == int(total) else total
        return None

    @property
    def display_note(self) -> str:
        return self.note


@dataclass
class PurchaseRequest:
    """Grouped purchase requisition with header fields and line items."""

    request_id: str
    source_type: str = "明細"
    sheet_name: str = ""
    budget_code: str = ""
    plan_name: str = ""
    level1: str = ""
    level2: str = ""
    level3: str = ""
    department: str = ""
    request_date: str = ""
    payer_name: str = ""
    receipt_count: int | None = None
    purpose: str = ""
    items: list[LineItem] = field(default_factory=list)

    @property
    def total_amount(self) -> int | float:
        total = sum(item.subtotal for item in self.items if item.subtotal is not None)
        return int(total) if total == int(total) else total


def load_config(config_path: Path) -> dict[str, Any]:
    """Load cell mapping configuration from JSON."""
    with config_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def parse_number(raw: str | None) -> int | float | None:
    """Parse a numeric string (with optional commas) into int or float."""
    if raw is None:
        return None
    cleaned = str(raw).strip().replace(",", "")
    if not cleaned:
        return None
    val = float(cleaned)
    return int(val) if val.is_integer() else val


def normalize_budget_code(code: str) -> str:
    """Normalize budget code (e.g. keep '321' or 'L30018' clean; strip '.0' if from CSV float)."""
    code = code.strip()
    if re.fullmatch(r"\d+\.0", code):
        return code[:-2]
    return code


def resolve_top_left_coord(ws: Worksheet, coord: str) -> str:
    """Return the top-left coordinate if `coord` lies inside a merged cell range."""
    cell = ws[coord]
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row <= cell.row <= merged_range.max_row
            and merged_range.min_col <= cell.column <= merged_range.max_col
        ):
            return f"{get_column_letter(merged_range.min_col)}{merged_range.min_row}"
    return coord


def safe_set_cell(
    ws: Worksheet,
    coord: str,
    value: Any,
    *,
    overwrite_formula: bool = False,
) -> bool:
    """Safely write to a cell, resolving merged ranges and preserving template formulas."""
    if not coord:
        return False
    target_coord = resolve_top_left_coord(ws, coord)
    existing = ws[target_coord].value
    if (
        not overwrite_formula
        and isinstance(existing, str)
        and existing.startswith("=")
    ):
        return False
    ws[target_coord].value = value
    return True


def read_requests_csv(csv_path: Path) -> list[PurchaseRequest]:
    """Read purchase request line items from CSV and group them by 請購單號."""
    encodings = ("utf-8-sig", "utf-8", "cp950", "big5")
    rows: list[dict[str, str]] | None = None
    last_err: Exception | None = None

    for enc in encodings:
        try:
            with csv_path.open("r", encoding=enc, newline="") as f:
                reader = csv.DictReader(f)
                rows = [
                    {k.strip(): (v.strip() if v is not None else "") for k, v in row.items() if k}
                    for row in reader
                ]
            break
        except UnicodeDecodeError as err:
            last_err = err

    if rows is None:
        raise RuntimeError(f"無法讀取 CSV 檔案 {csv_path}，編碼錯誤：{last_err}")

    grouped: dict[str, PurchaseRequest] = {}
    order: list[str] = []

    for idx, row in enumerate(rows, start=2):
        req_id = row.get("請購單號", "").strip()
        if not req_id:
            raise ValueError(f"CSV 第 {idx} 列缺少「請購單號」")

        if req_id not in grouped:
            raw_source = row.get("資料來源", "明細") or "明細"
            if raw_source not in ("明細", "發票"):
                raise ValueError(f"CSV 第 {idx} 列「資料來源」必須為「明細」或「發票」，收到：{raw_source}")
            raw_rc = parse_number(row.get("單據張數"))
            grouped[req_id] = PurchaseRequest(
                request_id=req_id,
                source_type=raw_source,
                sheet_name=row.get("工作表名稱", ""),
                budget_code=normalize_budget_code(row.get("預算科目代碼", "")),
                department=row.get("請購單位", ""),
                request_date=row.get("請購日期", ""),
                payer_name=row.get("代墊人", ""),
                receipt_count=int(raw_rc) if raw_rc is not None else None,
                purpose=row.get("用途說明", ""),
            )
            order.append(req_id)

        item_name = row.get("品名", "")
        invoice_no = row.get("發票號碼", "")
        manual_sub = parse_number(row.get("小計") or row.get("金額"))
        if item_name or invoice_no:
            display_name = item_name or f"發票請款 ({invoice_no})"
            grouped[req_id].items.append(
                LineItem(
                    name=display_name,
                    spec=row.get("規格", ""),
                    unit=row.get("單位", ""),
                    quantity=parse_number(row.get("數量")),
                    unit_price=parse_number(row.get("單價")),
                    manual_subtotal=manual_sub,
                    invoice_no=invoice_no,
                    note=row.get("備註", ""),
                )
            )

    return [grouped[req_id] for req_id in order]


def select_target_sheet(wb: openpyxl.Workbook, req: PurchaseRequest, cfg: dict[str, Any]) -> Worksheet:
    """Choose the appropriate worksheet in the template for a given request."""
    if req.sheet_name:
        if req.sheet_name not in wb.sheetnames:
            raise KeyError(f"範本中找不到工作表「{req.sheet_name}」，可用工作表：{wb.sheetnames}")
        return wb[req.sheet_name]

    if cfg.get("auto_select_sheet_by_total", False):
        threshold = cfg.get("under_3000_threshold", 3000)
        candidate = (
            cfg.get("under_3000_sheet")
            if req.total_amount < threshold
            else cfg.get("default_sheet")
        )
        if candidate and candidate in wb.sheetnames:
            return wb[candidate]

    default_sheet = cfg.get("default_sheet")
    if default_sheet and default_sheet in wb.sheetnames:
        return wb[default_sheet]

    return wb.active


def _detect_sheet_layout(ws: Worksheet, cfg: dict[str, Any]) -> dict[str, Any]:
    """Auto-detect table start_row, max_rows, total_coord, and header coords from sheet labels."""
    unit_price_row: int | None = None
    total_row: int | None = None
    purpose_label_row: int | None = None
    date_coord: str | None = None
    dept_coord: str | None = None
    receipt_coord: str | None = None
    plan_coord: str | None = None
    lv1_coord: str | None = None
    lv2_coord: str | None = None
    lv3_coord: str | None = None

    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 45), max_col=12):
        for cell in row:
            val = cell.value
            if not isinstance(val, str):
                continue
            compact = val.replace(" ", "").replace("\u3000", "").strip()
            if unit_price_row is None and compact == "單價":
                unit_price_row = cell.row
            elif total_row is None and compact == "合計":
                total_row = cell.row
            elif purpose_label_row is None and compact == "用途說明":
                purpose_label_row = cell.row
            elif receipt_coord is None and "黏貼單據" in compact:
                receipt_coord = cell.coordinate
            elif plan_coord is None and compact == "計畫名稱":
                plan_coord = f"{get_column_letter(cell.column)}{cell.row + 1}"
            elif lv1_coord is None and compact == "一級用途別":
                lv1_coord = f"{get_column_letter(cell.column)}{cell.row + 1}"
            elif lv2_coord is None and "二、三級用途別" in compact:
                col_l = get_column_letter(cell.column)
                lv2_coord = f"{col_l}{cell.row + 1}"
                lv3_coord = f"{col_l}{cell.row + 2}"
            elif "請購日期" in compact:
                if date_coord is None:
                    date_coord = f"H{cell.row}"
                if dept_coord is None:
                    dept_coord = f"C{cell.row}"
            elif dept_coord is None and cell.row >= 12 and "請購單位" in compact:
                dept_coord = f"C{cell.row}"

    start_row = (unit_price_row + 1) if unit_price_row is not None else int(cfg["items_table"]["start_row"])
    if total_row is not None and total_row > start_row:
        max_rows = total_row - start_row
        total_coord = f"{cfg['items_table']['columns'].get('subtotal', 'H')}{total_row}"
    else:
        max_rows = int(cfg["items_table"]["max_rows"])
        total_coord = cfg["cells"].get("total_amount", "")

    purpose_coord = f"A{purpose_label_row + 1}" if purpose_label_row is not None else cfg["cells"].get("purpose", "")
    return {
        "start_row": start_row,
        "max_rows": max_rows,
        "total_coord": total_coord,
        "purpose_coord": purpose_coord,
        "date_coord": date_coord or cfg["cells"].get("request_date", ""),
        "dept_coord": dept_coord or cfg["cells"].get("department", ""),
        "receipt_coord": receipt_coord or cfg["cells"].get("receipt_count", ""),
        "plan_coord": plan_coord,
        "lv1_coord": lv1_coord,
        "lv2_coord": lv2_coord,
        "lv3_coord": lv3_coord,
    }


def _hide_instruction_row(ws: Worksheet) -> None:
    """Clear and hide the top '操作步驟：...' yellow banner row if present."""
    from openpyxl.styles import PatternFill

    val = ws["A1"].value
    if isinstance(val, str) and "操作步驟" in val:
        ws["A1"].value = None
        no_fill = PatternFill(fill_type=None)
        for col in range(1, ws.max_column + 1):
            ws.cell(1, col).fill = no_fill
        ws.row_dimensions[1].hidden = True


def _normalize_vertical_alignments(ws: Worksheet) -> None:
    """Ensure vertical CJK labels (e.g. 請購人, 處室主任, 代墊欄) are horizontally centered."""
    from copy import copy

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            if (
                isinstance(cell.value, str)
                and cell.alignment
                and cell.alignment.textRotation == 255
                and (cell.alignment.horizontal != "center" or cell.alignment.indent)
            ):
                new_align = copy(cell.alignment)
                new_align.horizontal = "center"
                new_align.indent = 0
                cell.alignment = new_align


def _normalize_budget_alignments(
    ws: Worksheet,
    layout: dict[str, Any],
    cells: dict[str, Any],
) -> None:
    """Normalize alignment & indent in budget header cells so text never clips or skews."""
    from copy import copy
    from openpyxl.styles import Alignment

    # Multi-row merged cells (計畫名稱, 一級用途別): centered with wrap
    for key in ("plan_coord", "lv1_coord"):
        coord = layout.get(key)
        if not coord:
            continue
        cell = ws[coord]
        al = copy(cell.alignment) if cell.alignment else Alignment()
        al.horizontal = "center"
        al.vertical = "center"
        al.indent = 0
        al.wrap_text = True
        cell.alignment = al

    # Single-row merged cells (二級用途別 F8, 三級用途別 F9): remove stray indent=3 & shrink-to-fit on 1 line
    for key in ("lv2_coord", "lv3_coord"):
        coord = layout.get(key)
        if not coord:
            continue
        cell = ws[coord]
        al = copy(cell.alignment) if cell.alignment else Alignment()
        al.horizontal = "center"
        al.vertical = "center"
        al.indent = 0
        al.wrap_text = False
        al.shrink_to_fit = True
        cell.alignment = al

    # Budget code cell (H6): center cleanly in the yellow box
    bc_coord = cells.get("budget_code")
    if bc_coord:
        cell = ws[bc_coord]
        al = copy(cell.alignment) if cell.alignment else Alignment()
        al.horizontal = "center"
        al.vertical = "center"
        al.indent = 0
        cell.alignment = al


def _configure_single_page_print(ws: Worksheet) -> None:
    """Ensure the filled form fits cleanly on 1 A4 portrait page in both Excel and LibreOffice."""
    last_row = 30
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(r, c).value is not None for c in range(1, 11)):
            last_row = r
            break

    start_row = 2 if ws.row_dimensions[1].hidden else 1
    ws.print_area = f"A{start_row}:J{last_row}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.15
    ws.page_margins.footer = 0.15


def _expand_item_rows(ws: Worksheet, start_row: int, total_row: int, extra_rows: int) -> None:
    """Safely insert `extra_rows` before `total_row`, shifting merges, styles, heights, and SUM formulas."""
    from copy import copy

    if extra_rows <= 0:
        return

    src_item_row = total_row - 1
    old_max_row = ws.max_row

    # 1. Capture existing row heights for rows at/below total_row before shifting
    saved_heights: dict[int, float | None] = {
        r: ws.row_dimensions[r].height for r in range(total_row, old_max_row + 1)
    }
    src_item_height = ws.row_dimensions[src_item_row].height

    # 2. Capture all merged ranges and identify item-row merges to clone
    old_merges = [
        (m.min_row, m.min_col, m.max_row, m.max_col) for m in list(ws.merged_cells.ranges)
    ]
    item_row_merges = [
        (c1, c2) for (r1, c1, r2, c2) in old_merges if r1 == src_item_row and r2 == src_item_row
    ]

    # Unmerge ranges that intersect or lie below total_row
    for r1, c1, r2, c2 in old_merges:
        if r2 >= total_row:
            ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    # 3. Insert rows right before total_row
    ws.insert_rows(total_row, amount=extra_rows)

    # 4. Re-apply shifted/extended merged ranges
    for r1, c1, r2, c2 in old_merges:
        if r2 < total_row:
            continue
        if r1 < total_row <= r2:
            # Vertical merge spanning across total_row (e.g. J8:J31)
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2 + extra_rows, end_column=c2)
        else:
            ws.merge_cells(
                start_row=r1 + extra_rows,
                start_column=c1,
                end_row=r2 + extra_rows,
                end_column=c2,
            )

    # 5. Restore shifted row heights
    for old_r, h in saved_heights.items():
        if h is not None:
            ws.row_dimensions[old_r + extra_rows].height = h

    # 6. Style newly inserted item rows & clone item-row merges + subtotal formula
    for r in range(total_row, total_row + extra_rows):
        if src_item_height is not None:
            ws.row_dimensions[r].height = src_item_height
        for col_idx in range(1, 10):  # A..I (J is the vertical merged payer column)
            src_cell = ws.cell(src_item_row, col_idx)
            dst_cell = ws.cell(r, col_idx)
            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.border = copy(src_cell.border)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.number_format = src_cell.number_format
                dst_cell.alignment = copy(src_cell.alignment)
        for c1, c2 in item_row_merges:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        ws[f"H{r}"] = f"=ROUND(F{r}*G{r},0)"

    # 7. Update shifted 合計 formula and header_amount (=H{total_row}) reference
    new_total_row = total_row + extra_rows
    ws[f"H{new_total_row}"].value = f"=SUM(H{start_row}:H{new_total_row - 1})"

    for coord in ("I8", "H7"):
        val = ws[coord].value
        if isinstance(val, str) and val.strip().upper() == f"=H{total_row}":
            ws[coord].value = f"=H{new_total_row}"


def fill_worksheet(ws: Worksheet, req: PurchaseRequest, cfg: dict[str, Any]) -> None:
    """Populate a single worksheet with a PurchaseRequest's data."""
    _hide_instruction_row(ws)
    cells = cfg["cells"]
    layout = _detect_sheet_layout(ws, cfg)

    table_cfg = cfg["items_table"]
    start_row = layout["start_row"]
    template_rows = layout["max_rows"]
    max_allowed_rows = int(table_cfg.get("max_rows", 12))

    # Dynamically expand item rows if needed (up to max_allowed_rows, e.g. 12)
    if len(req.items) > template_rows and len(req.items) <= max_allowed_rows and layout["total_coord"]:
        m_tot = re.search(r"(\d+)$", layout["total_coord"])
        if m_tot:
            old_total_row = int(m_tot.group(1))
            extra = len(req.items) - template_rows
            _expand_item_rows(ws, start_row, old_total_row, extra)
            layout = _detect_sheet_layout(ws, cfg)

    if req.budget_code and cells.get("budget_code"):
        code_val: Any = int(req.budget_code) if req.budget_code.isdigit() else req.budget_code
        safe_set_cell(ws, cells["budget_code"], code_val)

    if req.plan_name and layout["plan_coord"]:
        safe_set_cell(ws, layout["plan_coord"], req.plan_name, overwrite_formula=True)
    if req.level1 and layout["lv1_coord"]:
        safe_set_cell(ws, layout["lv1_coord"], req.level1, overwrite_formula=True)
    if req.level2 and layout["lv2_coord"]:
        safe_set_cell(ws, layout["lv2_coord"], req.level2, overwrite_formula=True)
    if req.level3 and layout["lv3_coord"]:
        safe_set_cell(ws, layout["lv3_coord"], req.level3, overwrite_formula=True)

    if req.payer_name and cells.get("payer_name"):
        template = cells.get("payer_template", "{name}")
        safe_set_cell(ws, cells["payer_name"], template.format(name=req.payer_name))

    if req.receipt_count is not None and layout["receipt_coord"]:
        rc_template = cells.get("receipt_count_template", "黏貼單據  {count}  張")
        safe_set_cell(ws, layout["receipt_coord"], rc_template.format(count=req.receipt_count))

    if req.department and layout["dept_coord"]:
        safe_set_cell(ws, layout["dept_coord"], req.department)

    if req.request_date and layout["date_coord"]:
        safe_set_cell(ws, layout["date_coord"], req.request_date)

    if req.purpose and layout["purpose_coord"]:
        safe_set_cell(ws, layout["purpose_coord"], req.purpose)

    start_row = layout["start_row"]
    max_rows = layout["max_rows"]
    cols = table_cfg["columns"]

    if len(req.items) > max_rows:
        raise ValueError(
            f"請購單 {req.request_id} 有 {len(req.items)} 筆品項，超過範本上限 {max_rows} 列"
        )

    # Clear any leftover sample values in item rows before writing
    for offset in range(max_rows):
        row_num = start_row + offset
        for col_key in ("name", "spec", "unit", "quantity", "unit_price", "subtotal", "note"):
            col_letter = cols.get(col_key)
            if col_letter:
                safe_set_cell(ws, f"{col_letter}{row_num}", None)

    for offset, item in enumerate(req.items):
        row_num = start_row + offset
        if cols.get("name"):
            safe_set_cell(ws, f"{cols['name']}{row_num}", item.name)
        if cols.get("spec") and item.spec:
            safe_set_cell(ws, f"{cols['spec']}{row_num}", item.spec)
        if cols.get("unit") and item.unit:
            safe_set_cell(ws, f"{cols['unit']}{row_num}", item.unit)
        if cols.get("quantity") and item.quantity is not None:
            safe_set_cell(ws, f"{cols['quantity']}{row_num}", item.quantity)
        if cols.get("unit_price") and item.unit_price is not None:
            safe_set_cell(ws, f"{cols['unit_price']}{row_num}", item.unit_price)
        if cols.get("subtotal") and item.subtotal is not None:
            # When unit_price is omitted (e.g. lump-sum invoice), write manual subtotal directly;
            # otherwise preserve template formula like =F21*G21.
            force_subtotal = item.unit_price is None and item.manual_subtotal is not None
            safe_set_cell(
                ws,
                f"{cols['subtotal']}{row_num}",
                item.subtotal,
                overwrite_formula=force_subtotal,
            )
        if cols.get("note") and item.display_note:
            safe_set_cell(ws, f"{cols['note']}{row_num}", item.display_note)

    # Preserves =SUM(...) / header link formulas if present in template
    if layout["total_coord"]:
        safe_set_cell(ws, layout["total_coord"], req.total_amount)
    if cells.get("header_amount"):
        safe_set_cell(ws, cells["header_amount"], req.total_amount)
    _normalize_vertical_alignments(ws)
    _normalize_budget_alignments(ws, layout, cells)
    _configure_single_page_print(ws)


def sanitize_filename(name: str) -> str:
    """Remove characters that are invalid in Windows filenames."""
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()


def invoices_to_requests(
    invoices: list[ParsedInvoice],
    *,
    budget_code: str = "",
    department: str = "",
    payer_name: str = "",
    purpose: str = "",
    max_rows: int = 7,
    summarize_over_max: bool = True,
) -> list[PurchaseRequest]:
    """Convert ParsedInvoice objects into PurchaseRequest objects ready for Excel filling."""
    requests: list[PurchaseRequest] = []
    for idx, inv in enumerate(invoices, start=1):
        req_id = inv.invoice_no or f"INV-{idx:03d}"
        req_purpose = purpose or (f"{inv.seller_name}採購" if inv.seller_name else "發票採購")

        if summarize_over_max and len(inv.items) > max_rows and inv.total_amount is not None:
            seller_prefix = inv.seller_name or "發票"
            items = [
                LineItem(
                    name=f"{seller_prefix}採購一批(共{len(inv.items)}項，詳見發票明細)",
                    unit="批",
                    quantity=1,
                    manual_subtotal=inv.total_amount,
                    invoice_no=inv.invoice_no,
                )
            ]
        elif inv.items:
            items = [
                LineItem(
                    name=item.name,
                    spec=item.spec,
                    unit=item.unit,
                    quantity=item.quantity,
                    unit_price=item.unit_price,
                    manual_subtotal=item.subtotal if item.unit_price is None else None,
                    invoice_no=inv.invoice_no,
                    note=item.note,
                )
                for item in inv.items
            ]
        else:
            items = [
                LineItem(
                    name=f"發票請款 ({inv.invoice_no})" if inv.invoice_no else "發票請款",
                    unit="批",
                    quantity=1,
                    manual_subtotal=inv.total_amount,
                    invoice_no=inv.invoice_no,
                )
            ]

        requests.append(
            PurchaseRequest(
                request_id=req_id,
                source_type="發票",
                budget_code=normalize_budget_code(budget_code),
                department=department,
                request_date=inv.roc_date,
                payer_name=payer_name,
                receipt_count=1,
                purpose=req_purpose,
                items=items,
            )
        )
    return requests


def export_requests_csv(requests: list[PurchaseRequest], csv_path: Path) -> None:
    """Export PurchaseRequest objects into the standard batch CSV format (UTF-8 BOM)."""
    headers = [
        "請購單號", "資料來源", "工作表名稱", "預算科目代碼", "請購單位", "請購日期",
        "代墊人", "單據張數", "發票號碼", "用途說明", "品名", "規格",
        "單位", "數量", "單價", "金額", "備註",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for req in requests:
            for item in req.items:
                writer.writerow([
                    req.request_id,
                    req.source_type,
                    req.sheet_name,
                    req.budget_code,
                    req.department,
                    req.request_date,
                    req.payer_name,
                    req.receipt_count if req.receipt_count is not None else "",
                    item.invoice_no,
                    req.purpose,
                    item.name,
                    item.spec,
                    item.unit,
                    item.quantity if item.quantity is not None else "",
                    item.unit_price if item.unit_price is not None else "",
                    item.manual_subtotal if item.manual_subtotal is not None else "",
                    item.note,
                ])


def split_requests_opt(
    requests: list[PurchaseRequest],
    *,
    max_amount: int | float = 3000,
    max_rows: int = 7,
    allow_oversized: bool = True,
) -> list[PurchaseRequest]:
    """Use OR-Tools CP-SAT to optimally re-pack items into minimum forms under limits."""
    split_items: list[SplitItem] = []
    idx = 1
    for req in requests:
        for item in req.items:
            amt = item.subtotal
            if amt is None:
                raise ValueError(f"品項「{item.name}」缺少金額，無法進行最佳化分單")
            split_items.append(
                SplitItem(
                    item_id=idx,
                    name=item.name,
                    amount=amt,
                    budget_code=req.budget_code,
                    plan_name=req.plan_name,
                    level1=req.level1,
                    level2=req.level2,
                    level3=req.level3,
                    department=req.department,
                    request_date=req.request_date,
                    payer_name=req.payer_name,
                    purpose=req.purpose,
                    spec=item.spec,
                    unit=item.unit,
                    quantity=item.quantity,
                    unit_price=item.unit_price,
                    manual_subtotal=item.manual_subtotal,
                    invoice_no=item.invoice_no,
                    note=item.note,
                )
            )
            idx += 1

    plans = solve_form_splitting(
        split_items,
        max_amount=max_amount,
        max_rows=max_rows,
        group_by_budget_code=True,
        allow_oversized=allow_oversized,
    )

    new_requests: list[PurchaseRequest] = []
    for plan in plans:
        first = plan.items[0]
        has_invoice = any(it.invoice_no for it in plan.items)
        new_requests.append(
            PurchaseRequest(
                request_id=f"SPLIT-{plan.form_index:03d}",
                source_type="發票" if has_invoice else "明細",
                budget_code=plan.budget_code,
                plan_name=first.plan_name,
                level1=first.level1,
                level2=first.level2,
                level3=first.level3,
                department=plan.department,
                request_date=first.request_date,
                payer_name=first.payer_name,
                receipt_count=len({it.invoice_no for it in plan.items if it.invoice_no}) or None,
                purpose=first.purpose,
                items=[
                    LineItem(
                        name=it.name,
                        spec=it.spec,
                        unit=it.unit,
                        quantity=it.quantity,
                        unit_price=it.unit_price,
                        manual_subtotal=it.manual_subtotal,
                        invoice_no=it.invoice_no,
                        note=it.note,
                    )
                    for it in plan.items
                ],
            )
        )
    return new_requests


_SIMPLE_REF_RE = re.compile(r"^=\s*'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)\s*$", re.IGNORECASE)
_VLOOKUP_RE = re.compile(
    r"VLOOKUP\(\s*(?:'[^']+'!)?([A-Z]+\d+)\s*,\s*'?([^'!]+)'?!\$?([A-Z]+)\$?\d*:\$?([A-Z]+)\$?\d*\s*,\s*(\d+)",
    re.IGNORECASE,
)


def _resolve_cross_sheet_formulas(wb: openpyxl.Workbook, ws: Worksheet) -> None:
    """Resolve cross-sheet references and VLOOKUPs into static values before removing other sheets."""
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if not (isinstance(val, str) and val.startswith("=") and "!" in val):
                continue

            m_ref = _SIMPLE_REF_RE.match(val)
            if m_ref:
                ref_sheet, col_l, row_str = m_ref.groups()
                if ref_sheet in wb.sheetnames:
                    cell.value = wb[ref_sheet][f"{col_l}{row_str}"].value
                continue

            for m in _VLOOKUP_RE.finditer(val):
                key_coord, ref_sheet_name, start_col, _end_col, col_offset_str = m.groups()
                if ref_sheet_name not in wb.sheetnames:
                    continue
                raw_key = ws[key_coord].value
                if raw_key is None:
                    continue
                norm_key = normalize_budget_code(str(raw_key))
                ref_ws = wb[ref_sheet_name]
                start_idx = column_index_from_string(start_col)
                target_idx = start_idx + int(col_offset_str) - 1
                matched = False
                for r in range(1, ref_ws.max_row + 1):
                    cand = ref_ws.cell(r, start_idx).value
                    if cand is not None and normalize_budget_code(str(cand)) == norm_key:
                        cell.value = ref_ws.cell(r, target_idx).value
                        matched = True
                        break
                if matched:
                    break


def keep_only_sheet(wb: openpyxl.Workbook, target_ws: Worksheet) -> None:
    """Resolve cross-sheet references on `target_ws` and remove all other worksheets."""
    _resolve_cross_sheet_formulas(wb, target_ws)
    for name in list(wb.sheetnames):
        if name != target_ws.title:
            wb.remove(wb[name])


def chunk_requests_by_max_rows(
    requests: list[PurchaseRequest],
    max_rows: int = 7,
) -> list[PurchaseRequest]:
    """Automatically split any request with > max_rows items into sequential sheets of max_rows."""
    result: list[PurchaseRequest] = []
    for req in requests:
        if len(req.items) <= max_rows:
            result.append(req)
            continue
        total_parts = (len(req.items) + max_rows - 1) // max_rows
        for part_idx in range(total_parts):
            chunk = req.items[part_idx * max_rows : (part_idx + 1) * max_rows]
            result.append(
                PurchaseRequest(
                    request_id=f"{req.request_id}-{part_idx + 1}",
                    source_type=req.source_type,
                    sheet_name=req.sheet_name,
                    budget_code=req.budget_code,
                    plan_name=req.plan_name,
                    level1=req.level1,
                    level2=req.level2,
                    level3=req.level3,
                    department=req.department,
                    request_date=req.request_date,
                    payer_name=req.payer_name,
                    receipt_count=req.receipt_count if part_idx == 0 else None,
                    purpose=f"{req.purpose} ({part_idx + 1}/{total_parts})" if req.purpose else f"第 {part_idx + 1}/{total_parts} 張",
                    items=chunk,
                )
            )
    return result


def generate_from_requests(
    template_path: Path | None,
    requests: list[PurchaseRequest],
    output_dir: Path,
    cfg: dict[str, Any],
) -> list[Path]:
    """Generate one filled .xlsx file per PurchaseRequest in `output_dir`."""
    from form_builder import build_purchase_form_xlsx

    output_dir.mkdir(parents=True, exist_ok=True)
    only_target = bool(cfg.get("keep_only_target_sheet", True))
    max_rows = int(cfg["items_table"]["max_rows"])
    chunked_requests = chunk_requests_by_max_rows(requests, max_rows=max_rows)
    generated_files: list[Path] = []
    used_stems: set[str] = set()
    for req in chunked_requests:
        safe_id = sanitize_filename(req.request_id) or "REQ"
        candidate_stem = safe_id
        seq = 2
        while candidate_stem in used_stems:
            candidate_stem = f"{safe_id}_{seq}"
            seq += 1
        used_stems.add(candidate_stem)
        out_path = output_dir / f"{candidate_stem}_採購申請單.xlsx"
        if template_path is None:
            build_purchase_form_xlsx(req, out_path)
        else:
            wb = openpyxl.load_workbook(template_path)
            ws = select_target_sheet(wb, req, cfg)
            fill_worksheet(ws, req, cfg)
            if only_target and len(wb.sheetnames) > 1:
                keep_only_sheet(wb, ws)
            wb.save(out_path)
        generated_files.append(out_path)
    return generated_files


def batch_generate(
    template_path: Path,
    csv_path: Path,
    output_dir: Path,
    config_path: Path,
) -> list[Path]:
    """Generate one filled .xlsx file per purchase request in `output_dir`."""
    cfg = load_config(config_path)
    requests = read_requests_csv(csv_path)
    return generate_from_requests(template_path, requests, output_dir, cfg)


def inspect_template(template_path: Path) -> None:
    """Print worksheet names and key anchor coordinates to help verify config.json."""
    wb = openpyxl.load_workbook(template_path, data_only=False)
    print(f"範本檔案：{template_path}")
    print(f"工作表清單 ({len(wb.sheetnames)})：{wb.sheetnames}\n")

    keywords = ("憑證", "預算", "代墊", "請 購 日 期", "請購日期", "名", "合", "用途")
    for name in wb.sheetnames:
        ws = wb[name]
        print(f"=== 工作表：{name} (rows={ws.max_row}, cols={ws.max_column}) ===")
        hits: list[str] = []
        formulas: list[str] = []

        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 50), max_col=min(ws.max_column, 15)):
            for cell in row:
                val = cell.value
                if isinstance(val, str):
                    if val.startswith("="):
                        formulas.append(f"  公式 {cell.coordinate}: {val}")
                    elif any(kw in val.replace(" ", "") for kw in [k.replace(" ", "") for k in keywords]):
                        preview = val.replace("\n", " ")[:30]
                        hits.append(f"  標籤 {cell.coordinate}: {preview}")

        for h in hits[:20]:
            print(h)
        for f in formulas[:15]:
            print(f)
        print()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="臺南市立土城高中採購申請單暨支出憑證黏存單 - Excel 批次自動填寫工具"
    )
    parser.add_argument("--template", "-t", type=Path, help="原始 Excel 範本檔路徑 (.xlsx)")
    parser.add_argument("--csv", "-c", type=Path, help="採購明細清單 CSV 檔案路徑")
    parser.add_argument("--invoice-text", type=Path, help="電子發票文字檔路徑（如小北百貨發票明細）")
    parser.add_argument("--einvoice-csv", type=Path, help="財政部電子發票匯出 CSV 檔案路徑")
    parser.add_argument("--budget-code", default="", help="發票模式指定預算科目代碼（例如 321）")
    parser.add_argument("--department", default="", help="發票模式指定請購單位（例如 總務處）")
    parser.add_argument("--payer", default="", help="發票模式指定代墊人姓名")
    parser.add_argument("--purpose", default="", help="發票模式指定用途說明")
    parser.add_argument("--export-csv", type=Path, help="將解析後的發票匯出為標準 CSV 檔案")
    parser.add_argument("--auto-split", action="store_true", help="使用 OR-Tools CP-SAT 最佳化分單（依金額與列數上限）")
    parser.add_argument("--max-amount", type=float, default=3000, help="最佳化分單的單張金額上限 (預設: 3000)")
    parser.add_argument("--output", "-o", type=Path, default=Path("output"), help="輸出資料夾 (預設: output)")
    parser.add_argument("--config", type=Path, default=Path("config.json"), help="欄位設定檔路徑 (預設: config.json)")
    parser.add_argument("--inspect", action="store_true", help="檢查範本工作表名稱、欄位座標與公式")

    args = parser.parse_args()

    if args.inspect:
        if not args.template or not args.template.exists():
            parser.error("使用 --inspect 時必須提供存在的 --template 檔案")
        inspect_template(args.template)
        return

    cfg = load_config(args.config)
    max_rows = int(cfg["items_table"]["max_rows"])

    if args.invoice_text or args.einvoice_csv:
        if args.invoice_text:
            text = args.invoice_text.read_text(encoding="utf-8-sig")
            invoices = [parse_invoice_text(text)]
        else:
            invoices = parse_einvoice_csv(args.einvoice_csv)

        requests = invoices_to_requests(
            invoices,
            budget_code=args.budget_code,
            department=args.department,
            payer_name=args.payer,
            purpose=args.purpose,
            max_rows=max_rows,
        )
        if args.export_csv:
            export_requests_csv(requests, args.export_csv)
            print(f"已將發票資料匯出至：{args.export_csv}")
            if not args.template:
                return

        if args.template and not args.template.exists():
            raise FileNotFoundError(f"找不到範本檔案：{args.template}")

        files = generate_from_requests(args.template, requests, args.output, cfg)
    elif args.csv:
        requests = read_requests_csv(args.csv)
        if args.auto_split:
            requests = split_requests_opt(requests, max_amount=args.max_amount, max_rows=max_rows)
        if args.template and not args.template.exists():
            raise FileNotFoundError(f"找不到範本檔案：{args.template}")
        files = generate_from_requests(args.template, requests, args.output, cfg)
    else:
        parser.error("請提供 --csv、--invoice-text 或 --einvoice-csv")

    print(f"已成功產出 {len(files)} 份採購申請單：")
    for f in files:
        print(f"  - {f}")


if __name__ == "__main__":
    main()
