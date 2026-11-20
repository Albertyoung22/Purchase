"""Integration tests for the Flask + Waitress desktop app."""

import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from PIL import Image, ImageDraw

import app as web_app
from test_fill_purchase_forms import build_mock_template


class TestDesktopApp(unittest.TestCase):
    def setUp(self) -> None:
        self.client = web_app.app.test_client()

    def test_index_renders(self) -> None:
        res = self.client.get("/")
        self.assertEqual(res.status_code, 200)
        self.assertIn("採購申請單自動化系統", res.get_data(as_text=True))
        for icon_url in ("/static/favicon.svg", "/static/favicon.png"):
            icon_res = self.client.get(icon_url)
            self.assertEqual(icon_res.status_code, 200)
            icon_res.close()

    def test_parse_invoice_text_api(self) -> None:
        sample = Path(__file__).parent / "invoice_showba_sample.txt"
        res = self.client.post(
            "/api/parse-invoice-text",
            json={"text": sample.read_text(encoding="utf-8")},
        )
        data = res.get_json()
        self.assertTrue(data["ok"])
        self.assertEqual(data["invoice"]["invoice_no"], "AB-87654321")
        self.assertEqual(len(data["invoice"]["items"]), 3)

    def test_preview_with_cpsat_auto_split(self) -> None:
        res = self.client.post(
            "/api/preview",
            json={
                "mode": "manual",
                "source_type": "明細",
                "budget_code": "321",
                "department": "總務處",
                "auto_split": True,
                "max_amount": 3000,
                "items": [
                    {"name": "A", "quantity": 1, "unit_price": 1800},
                    {"name": "B", "quantity": 1, "unit_price": 1200},
                    {"name": "C", "quantity": 1, "unit_price": 1500},
                ],
            },
        )
        data = res.get_json()
        self.assertTrue(data["ok"])
        self.assertEqual(len(data["requests"]), 2)

    def test_generate_and_download_excel(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "template.xlsx"
            out_dir = tmp_dir / "output"
            build_mock_template(tpl_path)

            with (
                patch.object(web_app, "_active_template_path", return_value=tpl_path),
                patch.object(web_app, "OUTPUT_DIR", out_dir),
                patch.object(web_app, "HISTORY_PATH", tmp_dir / "history.json"),
            ):
                sample = Path(__file__).parent / "invoice_showba_sample.txt"
                res = self.client.post(
                    "/api/generate",
                    json={
                        "mode": "invoice_text",
                        "budget_code": "321",
                        "department": "總務處",
                        "payer_name": "楊敦州",
                        "purpose": "事務用品",
                        "invoice_text": sample.read_text(encoding="utf-8"),
                    },
                )
                data = res.get_json()
                self.assertTrue(data["ok"])
                self.assertEqual(len(data["files"]), 1)

                dl = self.client.get(f"/api/download/{data['files'][0]}")
                self.assertEqual(dl.status_code, 200)
                dl.close()

    def test_generate_standalone_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "output"
            with (
                patch.object(web_app, "_active_template_path", return_value=None),
                patch.object(web_app, "OUTPUT_DIR", out_dir),
                patch.object(web_app, "HISTORY_PATH", Path(tmp) / "history.json"),
            ):
                sample = Path(__file__).parent / "invoice_showba_sample.txt"
                res = self.client.post(
                    "/api/generate",
                    json={
                        "mode": "invoice_text",
                        "budget_code": "321",
                        "department": "總務處",
                        "payer_name": "楊敦州",
                        "purpose": "事務用品",
                        "invoice_text": sample.read_text(encoding="utf-8"),
                    },
                )
                data = res.get_json()
                self.assertTrue(data["ok"])
                out_file = out_dir / data["files"][0]
                wb = openpyxl.load_workbook(out_file, data_only=False)
                self.assertEqual(wb.sheetnames, ["採購申請單暨支出憑證黏存單"])
                ws = wb["採購申請單暨支出憑證黏存單"]
                self.assertEqual(ws["A2"].value, "採購申請單暨支出憑證黏存單")
                self.assertEqual(ws["H5"].value, "321")
                self.assertEqual(ws["B7"].value, "各校經常門分支計畫")
                self.assertEqual(ws["A18"].value, "3M菜瓜布")
                self.assertEqual(ws["C18"].value, "5片裝")
                self.assertEqual(ws["H25"].value, "=SUM(H18:H24)")

    def test_scan_image_api_with_rapidocr(self) -> None:
        img = Image.new("RGB", (600, 260), "white")
        draw = ImageDraw.Draw(img)
        lines = [
            "SHOWBA",
            "AB-12345678",
            "2026-11-18",
            "Pen  2  50  100TX",
            "Total: 100",
        ]
        y = 20
        for line in lines:
            draw.text((20, y), line, fill="black")
            y += 40

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)

        res = self.client.post(
            "/api/scan-image",
            data={"file": (buf, "receipt.png")},
            content_type="multipart/form-data",
        )
        data = res.get_json()
        self.assertTrue(data["ok"])
        self.assertEqual(data["engine"], "rapidocr+cpsat")
        self.assertEqual(data["recommended_sheet"], "請購單簡化流程-3000元以下適用")

    def test_budget_options_and_manual_plan_levels_override(self) -> None:
        opt_res = self.client.get("/api/budget-options")
        opt_data = opt_res.get_json()
        self.assertTrue(opt_data["ok"])
        self.assertIn("各校經常門分支計畫", opt_data["plan_names"])
        self.assertIn("材料及用品費", opt_data["level1_options"])

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            tpl_path = tmp_dir / "tpl.xlsx"
            out_dir = tmp_dir / "output"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "請購單簡化流程-3000元以下適用"
            ws["B7"] = "計畫名稱"
            ws["C7"] = "一級用途別"
            ws["E7"] = "二、三級用途別"
            ws["B8"] = "=VLOOKUP(H6,輸入區!A:E,2,FALSE)"
            ws["C8"] = "=VLOOKUP(H6,輸入區!A:E,3,FALSE)"
            ws["E8"] = "=VLOOKUP(H6,輸入區!A:E,4,FALSE)"
            ws["E9"] = "=VLOOKUP(H6,輸入區!A:E,5,FALSE)"
            ws["G20"] = "單 價"
            ws["A28"] = "合 計"
            wb.save(tpl_path)

            with (
                patch.object(web_app, "_active_template_path", return_value=tpl_path),
                patch.object(web_app, "OUTPUT_DIR", out_dir),
                patch.object(web_app, "HISTORY_PATH", tmp_dir / "history.json"),
            ):
                res = self.client.post(
                    "/api/generate",
                    json={
                        "mode": "manual",
                        "budget_code": "321",
                        "plan_name": "各校經常門分支計畫",
                        "level1": "材料及用品費",
                        "level2": "用品消耗",
                        "level3": "辦公（事務）用品",
                        "department": "總務處",
                        "items": [{"name": "原子筆", "quantity": 2, "unit_price": 30}],
                    },
                )
                data = res.get_json()
                self.assertTrue(data["ok"])
                out_wb = openpyxl.load_workbook(out_dir / data["files"][0], data_only=False)
                out_ws = out_wb["請購單簡化流程-3000元以下適用"]
                self.assertEqual(out_ws["H6"].value, 321)
                self.assertEqual(out_ws["B8"].value, "各校經常門分支計畫")
                self.assertEqual(out_ws["C8"].value, "材料及用品費")
                self.assertEqual(out_ws["E8"].value, "用品消耗")
                self.assertEqual(out_ws["E9"].value, "辦公（事務）用品")

    def test_collect_budget_options_scans_all_worksheets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tpl_path = Path(tmp) / "all_sheets.xlsx"
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "傳統藝術"
            ws1["B6"] = "計畫名稱"
            ws1["C6"] = "一級用途別"
            ws1["F6"] = "二、三級用途別"
            ws1["B7"] = "52121601推展傳統藝術教育"
            ws1["C7"] = "材料及用品費"
            ws1["F7"] = "用品消耗"
            ws1["F8"] = "其他用品消耗"
            ws1["J7"] = "因業務執行需要由(簽章) 楊敦州 代墊"
            ws1["A27"] = "用途說明"
            ws1["A28"] = "高中部辦公室老師座位名牌"

            ws2 = wb.create_sheet("輸入區")
            ws2["F10"] = "L32014"
            ws2["G10"] = "應付代收款"
            ws2["H10"] = "射箭訓練環境及設備改善經費"
            wb.save(tpl_path)

            with patch.object(web_app, "_active_template_path", return_value=tpl_path):
                res = self.client.get("/api/budget-options")
                data = res.get_json()
                self.assertTrue(data["ok"])
                self.assertIn("52121601推展傳統藝術教育", data["plan_names"])
                self.assertIn("總務處", data["departments"])
                self.assertIn("楊敦州", data["payers"])
                self.assertIn("高中部辦公室老師座位名牌", data["purposes"])
                codes = {e["code"]: e for e in data["entries"]}
                self.assertIn("L32014", codes)
                self.assertEqual(codes["L32014"]["plan_name"], "射箭訓練環境及設備改善經費")

    def test_scan_pdf_einvoice_api(self) -> None:
        import pymupdf

        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text(
            (50, 50),
            "SHOWBA\nAB-87654321\n2026-11-18\nSponge 49 x 2 98TX\nMarker 25 x 4 100TX\nTotal: 198\n",
        )
        pdf_bytes = doc.tobytes()
        doc.close()

        res = self.client.post(
            "/api/scan-image",
            data={"file": (io.BytesIO(pdf_bytes), "einvoice.pdf")},
            content_type="multipart/form-data",
        )
        data = res.get_json()
        self.assertTrue(data["ok"])
        self.assertEqual(data["engine"], "pdf-text+cpsat")
        self.assertTrue(data["cpsat_verified"])
        self.assertEqual(data["invoice"]["invoice_no"], "AB-87654321")
        self.assertEqual(data["invoice"]["total_amount"], 198)
        self.assertEqual(len(data["invoice"]["items"]), 2)
        self.assertEqual(data["recommended_sheet"], "請購單簡化流程-3000元以下適用")

    def test_cpsat_ocr_reconciliation_filters_payment_noise(self) -> None:
        from ocr_cpsat_parser import CandidateItem, reconcile_items_with_cpsat

        candidates = [
            CandidateItem(name="3M菜瓜布", quantity=2, unit_price=49, subtotal=98, has_tax_marker=True, arithmetic_ok=True),
            CandidateItem(name="奇異筆", quantity=4, unit_price=25, subtotal=100, has_tax_marker=True, arithmetic_ok=True),
            CandidateItem(name="現金", quantity=1, unit_price=500, subtotal=500, is_noise_hint=True),
            CandidateItem(name="找零", quantity=1, unit_price=302, subtotal=302, is_noise_hint=True),
        ]
        selected, exact = reconcile_items_with_cpsat(candidates, detected_total=198)
        self.assertTrue(exact)
        self.assertEqual([it.name for it in selected], ["3M菜瓜布", "奇異筆"])

    def test_scan_encrypted_pdf_with_password(self) -> None:
        import pymupdf

        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text(
            (50, 50),
            "SHOWBA\nAB-11223344\n2026-11-18\nTape 30 x 2 60TX\nTotal: 60\n",
        )
        pdf_bytes = doc.tobytes(encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="69123456", owner_pw="owner")
        doc.close()

        # Without password -> friendly error
        res_fail = self.client.post(
            "/api/scan-image",
            data={"file": (io.BytesIO(pdf_bytes), "enc.pdf")},
            content_type="multipart/form-data",
        )
        self.assertEqual(res_fail.status_code, 400)
        self.assertIn("已加密保護", res_fail.get_json()["error"])

        # With correct password -> succeeds
        res_ok = self.client.post(
            "/api/scan-image",
            data={"file": (io.BytesIO(pdf_bytes), "enc.pdf"), "pdf_password": "69123456"},
            content_type="multipart/form-data",
        )
        data_ok = res_ok.get_json()
        self.assertTrue(data_ok["ok"])
        self.assertEqual(data_ok["invoice"]["invoice_no"], "AB-11223344")
        self.assertEqual(data_ok["invoice"]["total_amount"], 60)

    def test_real_excel_print_preview_api(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "output"
            with (
                patch.object(web_app, "_active_template_path", return_value=None),
                patch.object(web_app, "OUTPUT_DIR", out_dir),
                patch.object(web_app, "HISTORY_PATH", Path(tmp) / "history.json"),
            ):
                res = self.client.post(
                    "/api/print-preview",
                    json={
                        "mode": "manual",
                        "budget_code": "321",
                        "department": "總務處",
                        "purpose": "真實列印預覽測試",
                        "items": [{"name": "測試品項", "unit": "個", "quantity": 2, "unit_price": 100}],
                    },
                )
                data = res.get_json()
                self.assertTrue(data["ok"])
                self.assertEqual(data["mode"], "excel_pdf")
                self.assertGreaterEqual(len(data["page_urls"]), 1)
                img_res = self.client.get(data["page_urls"][0])
                self.assertEqual(img_res.status_code, 200)
                img_res.close()

    def test_cumulative_history_api(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            out_dir = tmp_dir / "output"
            hist_path = tmp_dir / "history.json"
            with (
                patch.object(web_app, "_active_template_path", return_value=None),
                patch.object(web_app, "OUTPUT_DIR", out_dir),
                patch.object(web_app, "HISTORY_PATH", hist_path),
            ):
                # Generate 1st form (321, 200)
                self.client.post(
                    "/api/generate",
                    json={
                        "mode": "manual",
                        "budget_code": "321",
                        "department": "總務處",
                        "purpose": "文具1",
                        "items": [{"name": "原子筆", "quantity": 2, "unit_price": 100}],
                    },
                )
                # Generate 2nd form (322, 500)
                self.client.post(
                    "/api/generate",
                    json={
                        "mode": "manual",
                        "budget_code": "322",
                        "department": "教務處",
                        "purpose": "雜誌",
                        "items": [{"name": "期刊", "quantity": 1, "unit_price": 500}],
                    },
                )

                h_res = self.client.get("/api/history")
                h_data = h_res.get_json()
                self.assertTrue(h_data["ok"])
                self.assertEqual(h_data["summary"]["total_count"], 2)
                self.assertEqual(h_data["summary"]["total_amount"], 700)
                self.assertEqual(h_data["summary"]["by_budget_code"]["321"], 200)
                self.assertEqual(h_data["summary"]["by_budget_code"]["322"], 500)

                exp = self.client.get("/api/history/export")
                self.assertEqual(exp.status_code, 200)
                csv_text = exp.get_data(as_text=True)
                self.assertIn("原子筆", csv_text)
                self.assertIn("期刊", csv_text)
                exp.close()

                # Delete 1st record
                rec_id = h_data["records"][0]["id"]
                del_res = self.client.delete(f"/api/history/{rec_id}")
                self.assertEqual(del_res.get_json()["summary"]["total_count"], 1)

                # Clear remaining
                clr_res = self.client.post("/api/history/clear")
                self.assertEqual(clr_res.get_json()["summary"]["total_count"], 0)

    def test_session_upload_isolation_and_preview_security(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            upload_dir = tmp_dir / "uploads"
            out_dir = tmp_dir / "output"
            preview_root = out_dir / "_preview"
            stale_dir = preview_root / "stale123"
            stale_dir.mkdir(parents=True, exist_ok=True)
            stale_file = stale_dir / "old.png"
            stale_file.write_bytes(b"old")
            import os
            old_ts = stale_dir.stat().st_mtime - 3600
            os.utime(stale_dir, (old_ts, old_ts))

            web_app._cleanup_old_previews(preview_root, max_age_sec=1800)
            self.assertFalse(stale_dir.exists())

            with (
                patch.object(web_app, "UPLOAD_DIR", upload_dir),
                patch.object(web_app, "OUTPUT_DIR", out_dir),
            ):
                c1 = web_app.app.test_client()
                c2 = web_app.app.test_client()

                c1.post("/api/upload-csv", data={"file": (io.BytesIO(b"a,b\n1,2\n"), "a.csv")}, content_type="multipart/form-data")
                c2.post("/api/upload-csv", data={"file": (io.BytesIO(b"a,b\n3,4\n"), "b.csv")}, content_type="multipart/form-data")

                subdirs = [p for p in upload_dir.iterdir() if p.is_dir()]
                self.assertEqual(len(subdirs), 2)
                contents = sorted((p / "input.csv").read_text(encoding="utf-8") for p in subdirs)
                self.assertEqual(contents, ["a,b\n1,2\n", "a,b\n3,4\n"])

                bad_res = c1.get("/api/preview-file/../../app.py")
                self.assertEqual(bad_res.status_code, 404)
                bad_res.close()

    def test_gemini_api_key_sent_in_header_not_url(self) -> None:
        from unittest.mock import MagicMock
        import ocr_scanner

        img = Image.new("RGB", (60, 40), "white")
        with tempfile.TemporaryDirectory() as tmp:
            img_path = Path(tmp) / "inv.png"
            img.save(img_path)

            fake_resp = MagicMock()
            fake_resp.read.return_value = (
                b'{"candidates":[{"content":{"parts":[{"text":"{\\"document_type\\":\\"\\u767c\\u7968\\",\\"invoice_no\\":\\"AB-12345678\\",\\"date\\":\\"2026-11-18\\",\\"seller_name\\":\\"Test\\",\\"total_amount\\":100,\\"items\\":[{\\"name\\":\\"Pen\\",\\"spec\\":\\"\\",\\"unit\\":\\"\\",\\"quantity\\":1,\\"unit_price\\":100,\\"subtotal\\":100}]}"}]}}]}'
            )
            fake_resp.__enter__.return_value = fake_resp

            captured_req = {}

            def _fake_urlopen(req, timeout=30):
                captured_req["url"] = req.full_url
                captured_req["headers"] = {k.lower(): v for k, v in req.header_items()}
                return fake_resp

            with patch("urllib.request.urlopen", side_effect=_fake_urlopen):
                doc_type, _, inv = ocr_scanner._scan_with_gemini(img_path, "secret-key-123")

            self.assertEqual(doc_type, "發票")
            self.assertEqual(inv.invoice_no, "AB-12345678")
            self.assertNotIn("secret-key-123", captured_req["url"])
            self.assertEqual(captured_req["headers"].get("x-goog-api-key"), "secret-key-123")

    def test_travel_preview_generate_and_print_preview_api(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "output"
            payload = {
                "traveler_name": "王大明",
                "job_title": "組長",
                "rank_level": "簡任級以下",
                "budget_code": "321",
                "plan_name": "各校經常門分支計畫",
                "purpose": "教育局業務會議",
                "destination": "永華市政中心",
                "start_date": "115/11/20",
                "end_date": "115/11/20",
                "days": [
                    {
                        "date": "11/20",
                        "transport_mode": "self_car",
                        "segments_text": "土城高中->教育局:14, 教育局->土城高中:14",
                        "summary": "參加會議",
                    }
                ],
            }
            with patch.object(web_app, "OUTPUT_DIR", out_dir):
                prev_res = self.client.post("/api/travel/preview", json=payload)
                prev_data = prev_res.get_json()
                self.assertTrue(prev_data["ok"])
                # 28km * 3.25 = 91, one-way 14km -> misc 120, total = 211
                self.assertEqual(prev_data["grand_total"], 211)

                gen_res = self.client.post("/api/travel/generate", json=payload)
                gen_data = gen_res.get_json()
                self.assertTrue(gen_data["ok"])
                dl_res = self.client.get(f"/api/download/{gen_data['filename']}")
                self.assertEqual(dl_res.status_code, 200)
                dl_res.close()

                pp_res = self.client.post("/api/travel/print-preview", json=payload)
                pp_data = pp_res.get_json()
                self.assertTrue(pp_data["ok"])


if __name__ == "__main__":
    unittest.main()

