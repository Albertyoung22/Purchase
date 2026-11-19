"""Standalone builder for the school's '採購申請單暨支出憑證黏存單' .xlsx form."""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from fill_purchase_forms import PurchaseRequest

# Common budget codes from the school's code table (fallback when no external template is used)
BUDGET_CODE_MAP: dict[str, tuple[str, str, str, str]] = {
    "212": ("各校經常門分支計畫", "服務費用", "水電費", "工作場所電費"),
    "214": ("各校經常門分支計畫", "服務費用", "水電費", "工作場所水費"),
    "222": ("各校經常門分支計畫", "服務費用", "郵電費", "電話費"),
    "241": ("各校經常門分支計畫", "服務費用", "印刷裝訂及廣告費", "印刷及裝訂費"),
    "252": ("各校經常門分支計畫", "服務費用", "修理保養及保固費", "一般房屋修護費"),
    "255": ("各校經常門分支計畫", "服務費用", "修理保養及保固費", "機械及設備修護費"),
    "257": ("各校經常門分支計畫", "服務費用", "修理保養及保固費", "雜項設備修護費"),
    "279": ("各校經常門分支計畫", "服務費用", "一般服務費", "外包費"),
    "27F": ("各校經常門分支計畫", "服務費用", "一般服務費", "體育活動費"),
    "285": ("各校經常門分支計畫", "服務費用", "專業服務費", "講課鐘點、稿費、出席審查及查詢費"),
    "28A": ("各校經常門分支計畫", "服務費用", "專業服務費", "電腦軟體服務費"),
    "314": ("各校經常門分支計畫", "材料及用品費", "使用材料費", "建築材料"),
    "315": ("各校經常門分支計畫", "材料及用品費", "使用材料費", "設備零件"),
    "321": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "辦公（事務）用品"),
    "322": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "報章雜誌"),
    "323": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "農業與園藝用品及環境美化費"),
    "324": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "化學藥劑與實驗用品"),
    "326": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "食品"),
    "328": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "醫療用品(非醫療院所使用)"),
    "32Y": ("各校經常門分支計畫", "材料及用品費", "用品消耗", "其他用品消耗"),
    "514": ("由學校編列執行之機械及設備", "購建固定資產、無形資產及長期投資", "購置固定資產", "購置機械及設備"),
    "516": ("由學校編列執行之其他設備", "購建固定資產、無形資產及長期投資", "購置固定資產", "購置雜項設備"),
    "L30018": ("城西教育相關經費", "", "", ""),
    "L36001": ("應付代收款-活化教學與多元學習計畫", "", "", ""),
}

_THIN = Side(style="thin", color="000000")
_BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FONT_TITLE = Font(name="標楷體", size=14, bold=True)
_FONT_SUB = Font(name="標楷體", size=12, bold=True)
_FONT_BODY = Font(name="標楷體", size=10)
_FONT_BOLD = Font(name="標楷體", size=10, bold=True)
_FONT_SMALL = Font(name="標楷體", size=9)

_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _apply_border(ws: openpyxl.worksheet.worksheet.Worksheet, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.border = _BORDER_ALL


def build_purchase_form_xlsx(req: PurchaseRequest, output_path: Path) -> Path:
    """Create a standalone, print-ready '採購申請單暨支出憑證黏存單' .xlsx file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "採購申請單暨支出憑證黏存單"

    # Column widths (A..J) tuned for A4 portrait
    widths = {"A": 14, "B": 14, "C": 10, "D": 10, "E": 8, "F": 8, "G": 10, "H": 12, "I": 12, "J": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Header titles
    ws.merge_cells("A1:J1")
    ws["A1"] = "臺南市立土城高級中學"
    ws["A1"].font = _FONT_TITLE
    ws["A1"].alignment = _ALIGN_CENTER
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:J2")
    ws["A2"] = "採購申請單暨支出憑證黏存單"
    ws["A2"].font = _FONT_SUB
    ws["A2"].alignment = _ALIGN_CENTER
    ws.row_dimensions[2].height = 22

    ws.merge_cells("A3:J3")
    ws["A3"] = "簽 證 編 號："
    ws["A3"].font = _FONT_BODY
    ws["A3"].alignment = _ALIGN_LEFT

    ws.merge_cells("A4:B4")
    ws["A4"] = "傳票(付款憑單)編號："
    ws["A4"].font = _FONT_BODY
    ws.merge_cells("C4:H4")
    ws["C4"] = "□1154預付費用  □1315暫付及待結轉帳項"
    ws["C4"].font = _FONT_BODY
    ws.merge_cells("I4:J4")
    rc_text = f"黏貼單據  {req.receipt_count}  張" if req.receipt_count is not None else "黏貼單據     張"
    ws["I4"] = rc_text
    ws["I4"].font = _FONT_BODY
    ws["I4"].alignment = _ALIGN_RIGHT

    # Budget & amount table (Rows 5-8)
    ws["A5"] = "憑證編號"
    ws.merge_cells("B5:G5")
    ws["B5"] = "預      算      科      目"
    ws["H5"] = req.budget_code
    ws.merge_cells("I5:J5")
    ws["I5"] = "金      額"
    for coord in ("A5", "B5", "H5", "I5"):
        ws[coord].font = _FONT_BOLD
        ws[coord].alignment = _ALIGN_CENTER

    ws.merge_cells("A6:A8")
    ws["B6"] = "計畫名稱"
    ws["C6"] = "一級用途別"
    ws.merge_cells("D6:G6")
    ws["D6"] = "二、三級用途別"
    ws["H6"] = "金額"
    ws.merge_cells("I6:J6")
    ws["I6"] = "用途 / 代墊說明"
    for coord in ("B6", "C6", "D6", "H6", "I6"):
        ws[coord].font = _FONT_SMALL
        ws[coord].alignment = _ALIGN_CENTER

    default_plan, default_lv1, default_lv2, default_lv3 = BUDGET_CODE_MAP.get(req.budget_code, ("", "", "", ""))
    plan_name = req.plan_name or default_plan
    lv1 = req.level1 or default_lv1
    lv2 = req.level2 or default_lv2
    lv3 = req.level3 or default_lv3
    ws.merge_cells("B7:B8")
    ws["B7"] = plan_name
    ws["B7"].font = _FONT_SMALL
    ws["B7"].alignment = _ALIGN_CENTER

    ws.merge_cells("C7:C8")
    ws["C7"] = lv1
    ws["C7"].font = _FONT_SMALL
    ws["C7"].alignment = _ALIGN_CENTER

    ws.merge_cells("D7:G7")
    ws["D7"] = lv2
    ws["D7"].font = _FONT_SMALL
    ws["D7"].alignment = _ALIGN_CENTER

    ws.merge_cells("D8:G8")
    ws["D8"] = lv3
    ws["D8"].font = _FONT_SMALL
    ws["D8"].alignment = _ALIGN_CENTER

    ws.merge_cells("H7:H8")
    ws["H7"] = "=H25"
    ws["H7"].font = _FONT_BOLD
    ws["H7"].alignment = _ALIGN_RIGHT
    ws["H7"].number_format = "#,##0"

    ws.merge_cells("I7:J8")
    payer_text = (
        f"因業務執行需要由(簽章) {req.payer_name} 代墊"
        if req.payer_name
        else "因業務執行需要由(簽章)      代墊"
    )
    ws["I7"] = payer_text
    ws["I7"].font = _FONT_SMALL
    ws["I7"].alignment = _ALIGN_LEFT

    _apply_border(ws, "A5:J8")

    # Approval signatures (Rows 9-12)
    ws.merge_cells("A9:B9")
    ws["A9"] = "請  購  單  位"
    ws.merge_cells("C9:E9")
    ws["C9"] = "採  購  單  位"
    ws.merge_cells("F9:G9")
    ws["F9"] = "會  計  室"
    ws.merge_cells("H9:J9")
    ws["H9"] = "基 金 主 持 人"
    for coord in ("A9", "C9", "F9", "H9"):
        ws[coord].font = _FONT_BOLD
        ws[coord].alignment = _ALIGN_CENTER

    ws["A10"] = "驗收或證明"
    ws.merge_cells("C10:E10")
    ws["C10"] = "事務組長"
    ws["A11"] = "保管人"
    ws["B11"] = "財物或軟體登記"
    ws.merge_cells("C11:E11")
    ws["C11"] = "□消耗品 □非消耗品 □財產 □軟體"
    ws["A12"] = "主  管"
    ws.merge_cells("C12:E12")
    ws["C12"] = "總務主任"
    ws.merge_cells("F10:G12")
    ws.merge_cells("H10:J12")
    for row_idx in range(10, 13):
        ws.row_dimensions[row_idx].height = 20
        for col_idx in range(1, 11):
            ws.cell(row_idx, col_idx).font = _FONT_SMALL
            ws.cell(row_idx, col_idx).alignment = _ALIGN_CENTER
    _apply_border(ws, "A9:J12")

    # Receipt pasting area (Rows 13-14)
    ws.merge_cells("A13:J14")
    ws["A13"] = "支出憑證(統一發票或普通收據)黏貼處(估價單等附件訂於背面)"
    ws["A13"].font = _FONT_SMALL
    ws["A13"].alignment = _ALIGN_CENTER
    ws.row_dimensions[13].height = 24
    ws.row_dimensions[14].height = 24
    _apply_border(ws, "A13:J14")

    # Request unit & date (Row 15)
    ws.merge_cells("A15:B15")
    ws["A15"] = "請  購  單  位"
    ws.merge_cells("C15:D15")
    ws["C15"] = f"{req.department} 處（室）" if req.department else "處（室）"
    ws.merge_cells("E15:F15")
    ws["E15"] = "請 購 日 期"
    ws.merge_cells("G15:J15")
    ws["G15"] = req.request_date or "中華民國     年     月     日"
    for coord in ("A15", "C15", "E15", "G15"):
        ws[coord].font = _FONT_BODY
        ws[coord].alignment = _ALIGN_CENTER
    ws.row_dimensions[15].height = 22
    _apply_border(ws, "A15:J15")

    # Items table header (Rows 16-17)
    ws.merge_cells("A16:B17")
    ws["A16"] = "名          稱"
    ws.merge_cells("C16:D17")
    ws["C16"] = "規    格"
    ws.merge_cells("E16:E17")
    ws["E16"] = "單位"
    ws.merge_cells("F16:F17")
    ws["F16"] = "數量"
    ws.merge_cells("G16:H16")
    ws["G16"] = "估    價"
    ws["G17"] = "單價"
    ws["H17"] = "小計"
    ws.merge_cells("I16:J17")
    ws["I16"] = "備    註"
    for coord in ("A16", "C16", "E16", "F16", "G16", "G17", "H17", "I16"):
        ws[coord].font = _FONT_BOLD
        ws[coord].alignment = _ALIGN_CENTER
    _apply_border(ws, "A16:J17")

    # 7 Item rows (Rows 18-24) -> Total at Row 25
    start_row = 18
    for idx in range(7):
        r = start_row + idx
        ws.row_dimensions[r].height = 22
        ws.merge_cells(f"A{r}:B{r}")
        ws.merge_cells(f"C{r}:D{r}")
        ws.merge_cells(f"I{r}:J{r}")

        if idx < len(req.items):
            item = req.items[idx]
            ws[f"A{r}"] = item.name
            ws[f"C{r}"] = item.spec
            ws[f"E{r}"] = item.unit
            ws[f"F{r}"] = item.quantity
            if item.unit_price is not None:
                ws[f"G{r}"] = item.unit_price
                ws[f"H{r}"] = f"=F{r}*G{r}"
            elif item.manual_subtotal is not None:
                ws[f"H{r}"] = item.manual_subtotal
            else:
                ws[f"H{r}"] = f"=F{r}*G{r}"
            ws[f"I{r}"] = item.display_note
        else:
            ws[f"H{r}"] = f"=F{r}*G{r}"

        for c in ("A", "C", "E", "F", "G", "H", "I"):
            ws[f"{c}{r}"].font = _FONT_BODY
        ws[f"A{r}"].alignment = _ALIGN_LEFT
        ws[f"C{r}"].alignment = _ALIGN_LEFT
        ws[f"E{r}"].alignment = _ALIGN_CENTER
        ws[f"F{r}"].alignment = _ALIGN_RIGHT
        ws[f"G{r}"].alignment = _ALIGN_RIGHT
        ws[f"G{r}"].number_format = "#,##0"
        ws[f"H{r}"].alignment = _ALIGN_RIGHT
        ws[f"H{r}"].number_format = "#,##0"
        ws[f"I{r}"].alignment = _ALIGN_LEFT

    _apply_border(ws, "A18:J24")

    # Total row (Row 25)
    ws.merge_cells("A25:G25")
    ws["A25"] = "合      計"
    ws["A25"].font = _FONT_BOLD
    ws["A25"].alignment = _ALIGN_CENTER
    ws["H25"] = "=SUM(H18:H24)"
    ws["H25"].font = _FONT_BOLD
    ws["H25"].alignment = _ALIGN_RIGHT
    ws["H25"].number_format = "#,##0"
    ws.merge_cells("I25:J25")
    ws.row_dimensions[25].height = 22
    _apply_border(ws, "A25:J25")

    # Purpose section (Rows 26-27)
    ws.merge_cells("A26:J26")
    ws["A26"] = "用     途     說     明"
    ws["A26"].font = _FONT_BOLD
    ws["A26"].alignment = _ALIGN_CENTER
    ws.merge_cells("A27:J27")
    ws["A27"] = req.purpose
    ws["A27"].font = _FONT_BODY
    ws["A27"].alignment = _ALIGN_LEFT
    ws.row_dimensions[27].height = 26
    _apply_border(ws, "A26:J27")

    # Footer checkboxes & signatures (Rows 28-32)
    ws.merge_cells("A28:J28")
    ws["A28"] = "綠色採購指定項目( 1-46項) □是  □否"
    ws["A28"].font = _FONT_SMALL
    ws.merge_cells("A29:B29")
    ws["A29"] = "請  購  單  位"
    ws.merge_cells("C29:E29")
    ws["C29"] = "採  購  單  位"
    ws.merge_cells("F29:G29")
    ws["F29"] = "會  計  室"
    ws.merge_cells("H29:J29")
    ws["H29"] = "基 金 主 持 人"
    ws["A30"] = "請購人"
    ws.merge_cells("C30:E30")
    ws["C30"] = "□共同供應契約 □無共同供應契約 □比共同供應契約價格低"
    ws["A31"] = "處室主任"
    ws.merge_cells("C31:E31")
    ws["C31"] = "庶務組長 / 總務主任"
    ws.merge_cells("F30:G31")
    ws.merge_cells("H30:J31")
    for row_idx in range(28, 32):
        ws.row_dimensions[row_idx].height = 20
        for col_idx in range(1, 11):
            ws.cell(row_idx, col_idx).font = _FONT_SMALL
            ws.cell(row_idx, col_idx).alignment = _ALIGN_CENTER
    _apply_border(ws, "A28:J31")

    ws.merge_cells("A32:J32")
    ws["A32"] = "備註：代墊款項 1萬元以上檢附專案簽准簽呈"
    ws["A32"].font = _FONT_SMALL

    # A4 Portrait print setup
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path

