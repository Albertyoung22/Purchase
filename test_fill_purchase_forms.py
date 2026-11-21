"""Unit tests for fill_purchase_forms.py."""

import tempfile
import unittest
from pathlib import Path

import openpyxl

from fill_purchase_forms import (
    LineItem,
    PurchaseRequest,
    batch_generate,
    generate_from_requests,
    inspect_template,
    invoices_to_requests,
    load_config,
    split_requests_opt,
)
from form_splitter import SplitItem, solve_form_splitting
from invoice_parser import parse_einvoice_csv, parse_invoice_text


def build_mock_template(path: Path) -> None:
    """Create a mock .xlsx template matching the school's purchase form layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "採購申請單(一般)"

    ws["A1"] = "操作步驟：1.請先至輸入區查詢子目代號2.將子目代號複製至下表黃色欄位(H6)即可"
    ws["A2"] = "臺南市立土城高級中學"
    ws["A3"] = "採購申請單暨支出憑證黏存單"
    ws["I5"] = "黏貼單據     張"
    ws["A6"] = "憑證  編號"
    ws["B6"] = "預      算     科     目"
    ws["I6"] = "  金      額"

    # Simulated VLOOKUP / SUM formulas in the template
    ws["B8"] = "=VLOOKUP(H6,代碼表!A:E,2,FALSE)"
    ws["I8"] = "=H26"
    ws["J8"] = "因業務執行需要由(簽章)      代墊"

    ws["A16"] = "請  購  單  位 "
    ws.merge_cells("C16:D16")
    ws["F16"] = "請 購 日 期 "
    ws.merge_cells("H16:I16")
    ws["G18"] = "單 價"

    for r in range(19, 26):
        ws.merge_cells(f"A{r}:B{r}")
        ws.merge_cells(f"C{r}:D{r}")
        ws[f"H{r}"] = f"=F{r}*G{r}"

    ws["A26"] = "合     計"
    ws["H26"] = "=SUM(H19:H25)"
    ws["A27"] = "用     途     說     明"
    ws.merge_cells("A28:I28")

    wb.save(path)


class TestFillPurchaseForms(unittest.TestCase):
    def test_batch_generate_preserves_formulas_and_merged_cells(self) -> None:
        base_dir = Path(__file__).parent
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            template_path = tmp_dir / "template.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(template_path)

            files = batch_generate(
                template_path=template_path,
                csv_path=base_dir / "requests_sample.csv",
                output_dir=out_dir,
                config_path=base_dir / "config.json",
            )

            self.assertEqual(len(files), 3)
            self.assertEqual(files[0].name, "REQ-001_採購申請單.xlsx")
            self.assertEqual(files[1].name, "REQ-002_採購申請單.xlsx")
            self.assertEqual(files[2].name, "REQ-003_採購申請單.xlsx")

            # Verify REQ-001
            wb1 = openpyxl.load_workbook(files[0], data_only=False)
            ws1 = wb1["採購申請單(一般)"]
            self.assertIsNone(ws1["A1"].value)
            self.assertTrue(ws1.row_dimensions[1].hidden)
            self.assertEqual(ws1["H6"].value, 321)
            self.assertEqual(ws1["I5"].value, "黏貼單據  1  張")
            self.assertEqual(ws1["J8"].value, "因業務執行需要由(簽章) 楊敦州 代墊")
            self.assertEqual(ws1["C16"].value, "總務處")
            self.assertEqual(ws1["H16"].value, "中華民國 115 年 11 月 16 日")
            self.assertEqual(ws1["A19"].value, "職務名牌")
            self.assertEqual(ws1["E19"].value, "個")
            self.assertEqual(ws1["F19"].value, 25)
            self.assertEqual(ws1["G19"].value, 68)
            self.assertEqual(ws1["A28"].value, "高中部辦公室老師座位名牌")
            # Formulas must remain intact
            self.assertEqual(ws1["H19"].value, "=F19*G19")
            self.assertEqual(ws1["H26"].value, "=SUM(H19:H25)")
            self.assertEqual(ws1["I8"].value, "=H26")
            self.assertEqual(ws1["B8"].value, "=VLOOKUP(H6,代碼表!A:E,2,FALSE)")

            # Verify REQ-002 (4 items)
            wb2 = openpyxl.load_workbook(files[1], data_only=False)
            ws2 = wb2["採購申請單(一般)"]
            self.assertEqual(ws2["H6"].value, "L30018")
            self.assertEqual(ws2["A19"].value, "資訊插座+電源孔+線路收納")
            self.assertEqual(ws2["A20"].value, "24埠交換器")
            self.assertEqual(ws2["A21"].value, "總務處線路整理")
            self.assertEqual(ws2["A22"].value, "工資+五金+線材")
            self.assertEqual(ws2["F22"].value, 1)
            self.assertEqual(ws2["G22"].value, 11500)
            self.assertEqual(ws2["A28"].value, "校內維修")

            # Verify REQ-003 (invoice source without unit_price)
            wb3 = openpyxl.load_workbook(files[2], data_only=False)
            ws3 = wb3["採購申請單(一般)"]
            self.assertEqual(ws3["I5"].value, "黏貼單據  2  張")
            self.assertEqual(ws3["A19"].value, "文具用品一批")
            self.assertIsNone(ws3["G19"].value)
            self.assertEqual(ws3["H19"].value, 1580)
            self.assertEqual(ws3["I19"].value, "已附發票")

    def test_inspect_template_runs_cleanly(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            template_path = Path(tmp) / "template.xlsx"
            build_mock_template(template_path)
            inspect_template(template_path)

    def test_parse_showba_invoice_text_and_fill(self) -> None:
        base_dir = Path(__file__).parent
        text = (base_dir / "invoice_showba_sample.txt").read_text(encoding="utf-8")
        inv = parse_invoice_text(text)

        self.assertEqual(inv.invoice_no, "AB-87654321")
        self.assertEqual(inv.roc_date, "中華民國 115 年 11 月 18 日")
        self.assertEqual(inv.total_amount, 298)
        self.assertEqual(len(inv.items), 3)
        self.assertEqual(inv.items[0].name, "3M菜瓜布")
        self.assertEqual(inv.items[0].spec, "5片裝")
        self.assertEqual(inv.items[0].quantity, 2)
        self.assertEqual(inv.items[0].unit_price, 49)

        cfg = load_config(base_dir / "config.json")
        reqs = invoices_to_requests(
            [inv],
            budget_code="321",
            department="總務處",
            payer_name="陳大文",
            purpose="事務用品採購",
        )
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            template_path = tmp_dir / "template.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(template_path)

            files = generate_from_requests(template_path, reqs, out_dir, cfg)
            self.assertEqual(len(files), 1)
            wb = openpyxl.load_workbook(files[0], data_only=False)
            ws = wb["採購申請單(一般)"]
            self.assertEqual(ws["H6"].value, 321)
            self.assertEqual(ws["I5"].value, "黏貼單據  1  張")
            self.assertEqual(ws["H16"].value, "中華民國 115 年 11 月 18 日")
            self.assertEqual(ws["A19"].value, "3M菜瓜布")
            self.assertEqual(ws["C19"].value, "5片裝")
            self.assertEqual(ws["F19"].value, 2)
            self.assertEqual(ws["G19"].value, 49)
            self.assertIsNone(ws["I19"].value)
            self.assertEqual(ws["A21"].value, "延長線")
            self.assertEqual(ws["C21"].value, "3孔4座")

    def test_einvoice_csv_over_max_rows_summarizes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            csv_path = tmp_dir / "einvoice.csv"
            rows = ["發票號碼,發票日期,賣方名稱,品名,數量,單價,小計"]
            for i in range(1, 9):
                rows.append(f"CD11223344,2026/11/18,小北百貨,商品{i},1,50,50")
            csv_path.write_text("\n".join(rows), encoding="utf-8-sig")

            invs = parse_einvoice_csv(csv_path)
            self.assertEqual(len(invs), 1)
            self.assertEqual(invs[0].invoice_no, "CD-11223344")
            self.assertEqual(invs[0].total_amount, 400)

            reqs = invoices_to_requests(invs, budget_code="321", max_rows=7)
            self.assertEqual(len(reqs[0].items), 1)
            self.assertIn("共8項", reqs[0].items[0].name)
            self.assertEqual(reqs[0].items[0].manual_subtotal, 400)

    def test_cpsat_optimal_form_splitting_amount_and_rows(self) -> None:
        # Optimal packing into exactly 2 forms of 3000 each
        items = [
            SplitItem(item_id=1, name="A", amount=1600, budget_code="321", department="總務處"),
            SplitItem(item_id=2, name="B", amount=1300, budget_code="321", department="總務處"),
            SplitItem(item_id=3, name="C", amount=1400, budget_code="321", department="總務處"),
            SplitItem(item_id=4, name="D", amount=1200, budget_code="321", department="總務處"),
            SplitItem(item_id=5, name="E", amount=500, budget_code="321", department="總務處"),
        ]
        plans = solve_form_splitting(items, max_amount=3000, max_rows=7)
        self.assertEqual(len(plans), 2)
        for p in plans:
            self.assertLessEqual(p.total_amount, 3000)
            self.assertLessEqual(len(p.items), 7)

    def test_cpsat_respects_max_rows_and_budget_code_separation(self) -> None:
        items = [
            SplitItem(item_id=i, name=f"品項{i}", amount=100, budget_code="321", department="總務處")
            for i in range(1, 10)
        ] + [
            SplitItem(item_id=99, name="他科品項", amount=200, budget_code="322", department="總務處")
        ]
        plans = solve_form_splitting(items, max_amount=3000, max_rows=7)
        # 9 items in 321 -> 2 forms (max 7 rows), 1 item in 322 -> 1 form => total 3 forms
        self.assertEqual(len(plans), 3)
        codes_per_form = [{it.budget_code for it in p.items} for p in plans]
        self.assertTrue(all(len(c) == 1 for c in codes_per_form))

    def test_cpsat_rejects_item_exceeding_max_amount(self) -> None:
        req = PurchaseRequest(
            request_id="OVER-1",
            budget_code="321",
            department="總務處",
            items=[LineItem(name="高額設備", quantity=1, unit_price=3500)],
        )
        with self.assertRaises(ValueError):
            split_requests_opt([req], max_amount=3000, max_rows=7, allow_oversized=False)

    def test_cpsat_routes_oversized_item_to_over_3000_sheet(self) -> None:
        base_dir = Path(__file__).parent
        cfg = load_config(base_dir / "config.json")
        req = PurchaseRequest(
            request_id="OVER-3788",
            budget_code="321",
            department="總務處",
            items=[
                LineItem(name="碎紙機", quantity=1, unit_price=3788),
                LineItem(name="原子筆", quantity=2, unit_price=30),
            ],
        )
        split_reqs = split_requests_opt([req], max_amount=3000, max_rows=7)
        self.assertEqual(len(split_reqs), 2)
        totals = sorted(r.total_amount for r in split_reqs)
        self.assertEqual(totals, [60, 3788])

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "two.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(tpl_path)
            wb = openpyxl.load_workbook(tpl_path)
            ws_over = wb.active
            ws_over.title = "請購單-3000元以上一定要用此表"
            ws_under = wb.copy_worksheet(ws_over)
            ws_under.title = "請購單簡化流程-3000元以下適用"
            wb.save(tpl_path)

            files = generate_from_requests(tpl_path, split_reqs, out_dir, cfg)
            sheet_sets = [openpyxl.load_workbook(f).sheetnames[0] for f in files]
            self.assertEqual(
                sorted(sheet_sets),
                ["請購單-3000元以上一定要用此表", "請購單簡化流程-3000元以下適用"],
            )

    def test_multisheet_template_outputs_only_target_sheet_and_resolves_vlookup(self) -> None:
        base_dir = Path(__file__).parent
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "multi_template.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(tpl_path)

            # Add extra sheets (代碼表 & 費用動支請示單) to simulate the school's full workbook
            wb = openpyxl.load_workbook(tpl_path)
            ws_code = wb.create_sheet("代碼表")
            ws_code["A1"] = "代號"
            ws_code["B1"] = "計畫名稱"
            ws_code["A2"] = "321"
            ws_code["B2"] = "各校經常門分支計畫"
            wb.create_sheet("費用動支請示單")
            wb.save(tpl_path)

            cfg = load_config(base_dir / "config.json")
            req = PurchaseRequest(
                request_id="MULTI-1",
                budget_code="321",
                department="總務處",
                request_date="中華民國 115 年 11 月 18 日",
                purpose="測試單一工作表輸出",
                items=[LineItem(name="原子筆", unit="支", quantity=10, unit_price=15)],
            )
            files = generate_from_requests(tpl_path, [req], out_dir, cfg)
            out_wb = openpyxl.load_workbook(files[0], data_only=False)
            self.assertEqual(out_wb.sheetnames, ["採購申請單(一般)"])
            out_ws = out_wb["採購申請單(一般)"]
            self.assertEqual(out_ws["B8"].value, "各校經常門分支計畫")
            self.assertEqual(out_ws["H19"].value, "=F19*G19")
            self.assertEqual(out_ws["H26"].value, "=SUM(H19:H25)")

    def test_auto_select_3000_threshold_sheets(self) -> None:
        base_dir = Path(__file__).parent
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "two_sheets.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(tpl_path)

            wb = openpyxl.load_workbook(tpl_path)
            ws_over = wb.active
            ws_over.title = "請購單-3000元以上一定要用此表"
            ws_under = wb.copy_worksheet(ws_over)
            ws_under.title = "請購單簡化流程-3000元以下適用"
            wb.create_sheet("其他不相關工作表")
            wb.save(tpl_path)

            cfg = load_config(base_dir / "config.json")
            req_under = PurchaseRequest(
                request_id="UNDER-1700",
                budget_code="321",
                items=[LineItem(name="名牌", quantity=25, manual_subtotal=1700)],
            )
            req_over = PurchaseRequest(
                request_id="OVER-3500",
                budget_code="321",
                items=[LineItem(name="交換器", quantity=1, unit_price=3500)],
            )
            files = generate_from_requests(tpl_path, [req_under, req_over], out_dir, cfg)
            wb_u = openpyxl.load_workbook(files[0])
            wb_o = openpyxl.load_workbook(files[1])
            self.assertEqual(wb_u.sheetnames, ["請購單簡化流程-3000元以下適用"])
            self.assertEqual(wb_o.sheetnames, ["請購單-3000元以上一定要用此表"])

    def test_autodetect_header_row_20_and_clears_leftover_sample(self) -> None:
        base_dir = Path(__file__).parent
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "real_layout.xlsx"
            out_dir = tmp_dir / "output"

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "請購單簡化流程-3000元以下適用"
            ws["F18"] = "請 購 日 期 "
            ws.merge_cells("H18:I18")
            ws["H18"] = "中華民國 115 年 11 月 16 日"

            # Merged table header at rows 19-20 (just like the user's screenshot)
            ws.merge_cells("A19:B20")
            ws["A19"] = "名          稱"
            ws.merge_cells("F19:F20")
            ws["F19"] = "數 量"
            ws.merge_cells("G19:H19")
            ws["G19"] = "估    價"
            ws["G20"] = "單 價"
            ws["H20"] = "小  計"

            for r in range(21, 28):
                ws.merge_cells(f"A{r}:B{r}")
                ws[f"H{r}"] = f"=F{r}*G{r}"

            # Leftover sample item in row 21 & purpose in row 30
            ws["A21"] = "職務名牌"
            ws["A28"] = "合      計"
            ws["H28"] = "=SUM(H21:H27)"
            ws["A29"] = "用     途     說     明"
            ws.merge_cells("A30:I30")
            ws["A30"] = "高中部辦公室老師座位名牌"
            wb.save(tpl_path)

            cfg = load_config(base_dir / "config.json")
            req = PurchaseRequest(
                request_id="SCAN-1",
                budget_code="321",
                department="總務處",
                request_date="中華民國 115 年 11 月 18 日",
                purpose="小北百貨採購",
                items=[
                    LineItem(name="3M菜瓜布5片裝", unit="包", quantity=2, unit_price=49),
                    LineItem(name="奇異筆黑色", unit="支", quantity=4, unit_price=25),
                    LineItem(name="延長線3孔4座", unit="個", quantity=1, unit_price=100),
                ],
            )
            files = generate_from_requests(tpl_path, [req], out_dir, cfg)
            out_wb = openpyxl.load_workbook(files[0], data_only=False)
            out_ws = out_wb["請購單簡化流程-3000元以下適用"]
            # Header rows 19-20 must NOT be overwritten
            self.assertEqual(out_ws["A19"].value, "名          稱")
            self.assertEqual(out_ws["G20"].value, "單 價")
            # Items start at row 21
            self.assertEqual(out_ws["A21"].value, "3M菜瓜布5片裝")
            self.assertEqual(out_ws["A22"].value, "奇異筆黑色")
            self.assertEqual(out_ws["A23"].value, "延長線3孔4座")
            # Total formula at row 28 & purpose at row 30
            self.assertEqual(out_ws["H28"].value, "=SUM(H21:H27)")
            self.assertEqual(out_ws["A30"].value, "小北百貨採購")

    def test_up_to_12_items_expands_single_form_and_over_12_chunks(self) -> None:
        base_dir = Path(__file__).parent
        cfg = load_config(base_dir / "config.json")

        # 12 items on real template.xlsx -> expands to 1 single sheet!
        req_12 = PurchaseRequest(
            request_id="TWELVE",
            budget_code="321",
            department="總務處",
            purpose="12項單頁採購",
            items=[
                LineItem(name=f"品項{i}", unit="個", quantity=1, unit_price=300)
                for i in range(1, 13)
            ],
        )
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "output"
            tpl_real = base_dir / "template.xlsx"
            files_12 = generate_from_requests(tpl_real, [req_12], out_dir, cfg)
            self.assertEqual(len(files_12), 1)
            wb = openpyxl.load_workbook(files_12[0], data_only=False)
            ws = wb.active
            self.assertEqual(ws["A21"].value, "品項1")
            self.assertEqual(ws["A32"].value, "品項12")
            self.assertEqual(ws["H33"].value, "=SUM(H21:H32)")
            self.assertEqual(ws["I8"].value, "=H33")
            self.assertEqual(ws["A35"].value, "12項單頁採購")

        # 14 items (>12) -> auto-chunks into 2 sheets (12 + 2)
        req_14 = PurchaseRequest(
            request_id="MANY-14",
            budget_code="321",
            department="總務處",
            purpose="多品項採購",
            items=[
                LineItem(name=f"文具品項{i}", unit="個", quantity=1, unit_price=100)
                for i in range(1, 15)
            ],
        )
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "output"
            files = generate_from_requests(None, [req_14], out_dir, cfg)
            self.assertEqual(len(files), 2)
            self.assertEqual(files[0].name, "MANY-14-1_採購申請單.xlsx")
            self.assertEqual(files[1].name, "MANY-14-2_採購申請單.xlsx")
            wb1 = openpyxl.load_workbook(files[0])
            wb2 = openpyxl.load_workbook(files[1])
            self.assertEqual(wb1.active["A18"].value, "文具品項1")
            self.assertEqual(wb1.active["A29"].value, "文具品項12")
            self.assertEqual(wb2.active["A18"].value, "文具品項13")
            self.assertEqual(wb2.active["A19"].value, "文具品項14")

    def test_numeric_cells_prevent_hash_overflow(self) -> None:
        base_dir = Path(__file__).parent
        cfg = load_config(base_dir / "config.json")
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "narrow.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(tpl_path)

            wb = openpyxl.load_workbook(tpl_path)
            ws = wb.active
            ws.column_dimensions["F"].width = 5.0
            ws.column_dimensions["G"].width = 6.0
            ws.column_dimensions["H"].width = 7.0
            ws["A33"] = "請購人"
            ws["A33"].alignment = openpyxl.styles.Alignment(vertical="center", textRotation=255)
            wb.save(tpl_path)

            req = PurchaseRequest(
                request_id="NUM-FIT",
                budget_code="321",
                department="總務處",
                purpose="大額測試",
                items=[LineItem(name="投影機", unit="台", quantity=2, unit_price=12500)],
            )
            files = generate_from_requests(tpl_path, [req], out_dir, cfg)
            out_ws = openpyxl.load_workbook(files[0]).active
            self.assertEqual(out_ws["A33"].alignment.horizontal, "center")
            self.assertGreaterEqual(out_ws.column_dimensions["F"].width, 8.5)
            self.assertGreaterEqual(out_ws.column_dimensions["G"].width, 11.5)
            self.assertGreaterEqual(out_ws.column_dimensions["H"].width, 13.5)
            for coord in ("F19", "G19", "H19", "H26", "I8"):
                self.assertTrue(out_ws[coord].alignment.shrinkToFit)
                self.assertFalse(out_ws[coord].alignment.wrapText)


if __name__ == "__main__":
    unittest.main()
