"""Flask + Waitress desktop application for school purchase requisition automation."""

from __future__ import annotations

import io
import json
import os
import re
import shutil
import socket
import subprocess
import tempfile
import threading
import time
import webbrowser
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from flask import Flask, jsonify, render_template, request, send_file
from waitress import serve

from fill_purchase_forms import (
    LineItem,
    PurchaseRequest,
    chunk_requests_by_max_rows,
    generate_from_requests,
    invoices_to_requests,
    load_config,
    normalize_budget_code,
    parse_number,
    read_requests_csv,
    split_requests_opt,
)
from form_builder import BUDGET_CODE_MAP
from history_store import (
    append_history,
    clear_history,
    delete_history_record,
    export_history_csv_bytes,
    load_history,
    summarize_history,
)
from invoice_parser import ParsedInvoice, parse_einvoice_csv, parse_invoice_text
from ocr_cpsat_parser import extract_candidates_from_text, reconcile_items_with_cpsat
from ocr_scanner import scan_document
from print_preview import (
    export_xlsx_to_pdf,
    render_pdf_to_pngs,
    render_xlsx_fallback_html,
)

BASE_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = BASE_DIR / "config.json"
HISTORY_PATH = BASE_DIR / "history.json"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB


_BUDGET_OPTIONS_CACHE: dict[str, Any] = {"key": None, "data": None}


def _is_cloud_mode() -> bool:
    return bool(os.environ.get("RENDER") or os.environ.get("PORT"))


def _active_template_path() -> Path | None:
    uploaded = UPLOAD_DIR / "template.xlsx"
    if uploaded.exists():
        return uploaded
    local = BASE_DIR / "template.xlsx"
    if local.exists():
        return local
    return None


def _collect_budget_options() -> dict[str, Any]:
    """Collect budget code mappings & dropdown options from every worksheet in the template."""
    tpl = _active_template_path()
    cache_key = (str(tpl.resolve()), tpl.stat().st_mtime_ns) if tpl and tpl.exists() else None
    if _BUDGET_OPTIONS_CACHE["key"] == cache_key and _BUDGET_OPTIONS_CACHE["data"] is not None:
        return _BUDGET_OPTIONS_CACHE["data"]

    entries: dict[str, dict[str, str]] = {}
    for code, (plan, lv1, lv2, lv3) in BUDGET_CODE_MAP.items():
        entries[code] = {
            "code": code,
            "plan_name": plan,
            "level1": lv1,
            "level2": lv2,
            "level3": lv3,
        }

    extra_plans: list[str] = []
    extra_lv1: list[str] = []
    extra_lv2: list[str] = []
    extra_lv3: list[str] = []
    extra_depts: list[str] = ["總務處", "教務處", "學務處", "輔導室", "人事室", "秘書室", "圖書館", "國中部辦公室"]
    extra_payers: list[str] = []
    extra_purposes: list[str] = ["事務用品採購", "辦公用品採購", "校內維修", "教學設備採購", "活動用品採購"]

    payer_re = re.compile(r"由\(簽章\)\s*(\S+?)\s*代墊")

    if tpl and tpl.exists():
        try:
            wb = openpyxl.load_workbook(tpl, data_only=True, read_only=True)
            for name in wb.sheetnames:
                ws = wb[name]
                rows_data = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 400, 400), max_col=25, values_only=True))

                # 1. Extract from each sheet's header, payer, department & purpose
                for r_idx, row in enumerate(rows_data[:40]):
                    for c_idx, cell_val in enumerate(row):
                        if not isinstance(cell_val, str):
                            continue
                        text_val = cell_val.strip()
                        compact = text_val.replace(" ", "").replace("\u3000", "")
                        if compact == "計畫名稱" and r_idx + 1 < len(rows_data):
                            v = rows_data[r_idx + 1][c_idx]
                            if v and not str(v).startswith("#"):
                                extra_plans.append(str(v).strip())
                        elif compact == "一級用途別" and r_idx + 1 < len(rows_data):
                            v = rows_data[r_idx + 1][c_idx]
                            if v and not str(v).startswith("#"):
                                extra_lv1.append(str(v).strip())
                        elif "二、三級用途別" in compact:
                            if r_idx + 1 < len(rows_data):
                                v2 = rows_data[r_idx + 1][c_idx]
                                if v2 and not str(v2).startswith("#"):
                                    extra_lv2.append(str(v2).strip())
                            if r_idx + 2 < len(rows_data):
                                v3 = rows_data[r_idx + 2][c_idx]
                                if v3 and not str(v3).startswith("#"):
                                    extra_lv3.append(str(v3).strip())
                        elif "代墊" in compact:
                            m_p = payer_re.search(text_val)
                            if m_p:
                                extra_payers.append(m_p.group(1).strip())
                        elif compact == "用途說明" and r_idx + 1 < len(rows_data):
                            pv = rows_data[r_idx + 1][0]
                            if pv and not str(pv).startswith("#"):
                                extra_purposes.append(str(pv).strip())
                        elif compact in ("總務處", "教務處", "學務處", "輔導室", "人事室", "秘書室", "圖書館"):
                            extra_depts.append(compact)

                # 2. Extract from 5-col budget code table and 3-col L/R/T/E agency-fund tables
                for row in rows_data:
                    for idx in range(len(row) - 2):
                        c_val = row[idx]
                        if c_val is None:
                            continue
                        code_str = normalize_budget_code(str(c_val))
                        if code_str in ("代號", "代碼", "") or len(code_str) > 16:
                            continue
                        if not any(ch.isdigit() for ch in code_str):
                            continue

                        # 3-column L/R/T/E table: 代號, 應付代收款/存入保證金..., 計畫名稱
                        if code_str[0] in ("L", "R", "T", "E") and row[idx + 1] is not None:
                            cat = str(row[idx + 1]).strip()
                            sub = str(row[idx + 2]).strip() if row[idx + 2] else ""
                            if (cat or sub) and code_str not in entries:
                                entries[code_str] = {
                                    "code": code_str,
                                    "plan_name": sub or cat,
                                    "level1": cat if sub else "",
                                    "level2": "",
                                    "level3": "",
                                }
                            continue

                        # 5-column table: 代號, 計畫名稱, 一級用途別, 二級用途別, 三級用途別
                        if (
                            idx + 4 < len(row)
                            and row[idx + 1] is not None
                            and row[idx + 2] is not None
                            and (row[idx + 3] is not None or row[idx + 4] is not None)
                        ):
                            p_str = str(row[idx + 1]).strip()
                            l1_str = str(row[idx + 2]).strip()
                            l2_str = str(row[idx + 3]).strip() if row[idx + 3] else ""
                            l3_str = str(row[idx + 4]).strip() if row[idx + 4] else ""
                            if p_str and p_str not in ("計畫名稱", "23") and l1_str not in ("23",):
                                entries[code_str] = {
                                    "code": code_str,
                                    "plan_name": p_str,
                                    "level1": l1_str,
                                    "level2": l2_str,
                                    "level3": l3_str,
                                }
                                continue
            wb.close()
        except Exception:
            pass

    ordered = list(entries.values())

    def _uniq(vals: list[str]) -> list[str]:
        seen: list[str] = []
        for v in vals:
            if v and v not in seen:
                seen.append(v)
        return seen

    result = {
        "entries": ordered,
        "plan_names": _uniq([e["plan_name"] for e in ordered] + extra_plans),
        "level1_options": _uniq([e["level1"] for e in ordered] + extra_lv1),
        "level2_options": _uniq([e["level2"] for e in ordered] + extra_lv2),
        "level3_options": _uniq([e["level3"] for e in ordered] + extra_lv3),
        "departments": _uniq(extra_depts),
        "payers": _uniq(extra_payers),
        "purposes": _uniq(extra_purposes),
    }
    _BUDGET_OPTIONS_CACHE["key"] = cache_key
    _BUDGET_OPTIONS_CACHE["data"] = result
    return result


def _requests_from_payload(payload: dict[str, Any], cfg: dict[str, Any]) -> list[PurchaseRequest]:
    """Build PurchaseRequest list from UI JSON payload."""
    mode = payload.get("mode", "invoice_text")
    max_rows = int(cfg["items_table"]["max_rows"])
    budget_code = payload.get("budget_code", "")
    plan_name = str(payload.get("plan_name", "")).strip()
    level1 = str(payload.get("level1", "")).strip()
    level2 = str(payload.get("level2", "")).strip()
    level3 = str(payload.get("level3", "")).strip()
    department = payload.get("department", "")
    payer_name = payload.get("payer_name", "")
    purpose = payload.get("purpose", "")
    request_date = payload.get("request_date", "")

    def _apply_budget_fields(reqs: list[PurchaseRequest]) -> list[PurchaseRequest]:
        for r in reqs:
            if plan_name:
                r.plan_name = plan_name
            if level1:
                r.level1 = level1
            if level2:
                r.level2 = level2
            if level3:
                r.level3 = level3
        return reqs

    if mode == "invoice_text":
        text = payload.get("invoice_text", "").strip()
        if not text:
            raise ValueError("請貼上電子發票文字明細")
        inv = parse_invoice_text(text)
        reqs = invoices_to_requests(
            [inv],
            budget_code=budget_code,
            department=department,
            payer_name=payer_name,
            purpose=purpose,
            max_rows=max_rows,
            summarize_over_max=not bool(payload.get("auto_split")),
        )
        if request_date:
            reqs[0].request_date = request_date
        return _apply_budget_fields(reqs)

    if mode == "manual":
        raw_items = payload.get("items", [])
        if not raw_items:
            raise ValueError("請至少新增一筆品項明細")
        source_type = payload.get("source_type", "明細")
        invoice_no = payload.get("invoice_no", "").strip()
        rc_raw = parse_number(str(payload.get("receipt_count", "")))
        items: list[LineItem] = []
        for row in raw_items:
            name = str(row.get("name", "")).strip()
            if not name:
                continue
            qty = parse_number(str(row.get("quantity", "")))
            price = parse_number(str(row.get("unit_price", "")))
            manual_sub = parse_number(str(row.get("subtotal", "")))
            items.append(
                LineItem(
                    name=name,
                    spec=str(row.get("spec", "")).strip(),
                    unit=str(row.get("unit", "")).strip(),
                    quantity=qty,
                    unit_price=price,
                    manual_subtotal=manual_sub if price is None else None,
                    invoice_no=invoice_no,
                    note=str(row.get("note", "")).strip(),
                )
            )
        if not items:
            raise ValueError("請填寫至少一筆有效的品名")
        return _apply_budget_fields([
            PurchaseRequest(
                request_id=payload.get("request_id") or (invoice_no or "REQ-001"),
                source_type=source_type if source_type in ("明細", "發票") else "明細",
                budget_code=normalize_budget_code(budget_code),
                department=department,
                request_date=request_date,
                payer_name=payer_name,
                receipt_count=int(rc_raw) if rc_raw is not None else None,
                purpose=purpose,
                items=items,
            )
        ])

    if mode == "csv":
        csv_kind = payload.get("csv_kind", "standard")
        uploaded_csv = UPLOAD_DIR / "input.csv"
        if not uploaded_csv.exists():
            raise ValueError("請先上傳 CSV 檔案")
        if csv_kind == "einvoice":
            invs = parse_einvoice_csv(uploaded_csv)
            if not invs:
                raise ValueError("電子發票 CSV 中找不到有效發票資料")
            return _apply_budget_fields(
                invoices_to_requests(
                    invs,
                    budget_code=budget_code,
                    department=department,
                    payer_name=payer_name,
                    purpose=purpose,
                    max_rows=max_rows,
                    summarize_over_max=not bool(payload.get("auto_split")),
                )
            )
        return _apply_budget_fields(read_requests_csv(uploaded_csv))

    raise ValueError(f"不支援的資料來源模式：{mode}")


def _serialize_requests(reqs: list[PurchaseRequest]) -> list[dict[str, Any]]:
    return [
        {
            "request_id": r.request_id,
            "source_type": r.source_type,
            "budget_code": r.budget_code,
            "plan_name": r.plan_name,
            "level1": r.level1,
            "level2": r.level2,
            "level3": r.level3,
            "department": r.department,
            "request_date": r.request_date,
            "payer_name": r.payer_name,
            "receipt_count": r.receipt_count,
            "purpose": r.purpose,
            "total_amount": r.total_amount,
            "items": [
                {
                    "name": it.name,
                    "spec": it.spec,
                    "unit": it.unit,
                    "quantity": it.quantity,
                    "unit_price": it.unit_price,
                    "subtotal": it.subtotal,
                    "note": it.display_note,
                }
                for it in r.items
            ],
        }
        for r in reqs
    ]


@app.get("/")
def index() -> str:
    cfg = load_config(CONFIG_PATH)
    tpl = _active_template_path()
    sample_inv_path = BASE_DIR / "invoice_showba_sample.txt"
    sample_inv = sample_inv_path.read_text(encoding="utf-8") if sample_inv_path.exists() else ""
    return render_template(
        "index.html",
        config_json=json.dumps(cfg, ensure_ascii=False, indent=2),
        has_template=tpl is not None,
        template_name=tpl.name if tpl else "",
        sample_invoice=sample_inv,
        is_cloud=_is_cloud_mode(),
    )


@app.get("/api/budget-options")
def api_budget_options():
    return jsonify({"ok": True, **_collect_budget_options()})


@app.post("/api/upload-template")
def upload_template():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "請選擇 .xlsx 範本檔案"}), 400
    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({"ok": False, "error": "僅支援 .xlsx 格式，若為 .xls 請先用 Excel 另存為 .xlsx"}), 400
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    dest = UPLOAD_DIR / "template.xlsx"
    file.save(dest)
    return jsonify({"ok": True, "filename": file.filename})


@app.post("/api/upload-csv")
def upload_csv():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "請選擇 CSV 檔案"}), 400
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    dest = UPLOAD_DIR / "input.csv"
    file.save(dest)
    return jsonify({"ok": True, "filename": file.filename})


@app.get("/api/inspect-template")
def api_inspect_template():
    tpl = _active_template_path()
    if not tpl:
        return jsonify({"ok": False, "error": "尚未上傳或放置 template.xlsx 範本檔"}), 400
    wb = openpyxl.load_workbook(tpl, data_only=False)
    try:
        keywords = ("憑證", "預算", "代墊", "請購日期", "名", "合計", "用途")
        sheets = []
        for name in wb.sheetnames:
            ws = wb[name]
            labels = []
            formulas = []
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 50), max_col=min(ws.max_column, 15)):
                for cell in row:
                    val = cell.value
                    if isinstance(val, str):
                        if val.startswith("="):
                            formulas.append(f"{cell.coordinate}: {val}")
                        elif any(kw in val.replace(" ", "") for kw in keywords):
                            labels.append(f"{cell.coordinate}: {val.strip()[:28]}")
            sheets.append({
                "name": name,
                "rows": ws.max_row,
                "cols": ws.max_column,
                "labels": labels[:15],
                "formulas": formulas[:12],
            })
        return jsonify({"ok": True, "template": tpl.name, "sheets": sheets})
    finally:
        wb.close()


@app.post("/api/save-config")
def api_save_config():
    try:
        data = request.get_json(force=True)
        parsed = json.loads(data["config_text"])
        CONFIG_PATH.write_text(json.dumps(parsed, ensure_ascii=False, indent=2), encoding="utf-8")
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/scan-image")
def api_scan_image():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "請選擇發票或估價單檔案（圖片或 PDF）"}), 400
    ext = Path(file.filename).suffix.lower()
    if ext not in (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".pdf"):
        return jsonify({"ok": False, "error": "僅支援圖片 (.jpg, .png, .webp, .bmp) 或 .pdf 格式"}), 400

    gemini_key = request.form.get("gemini_api_key", "")
    pdf_password = request.form.get("pdf_password", "")
    cfg = load_config(CONFIG_PATH)
    threshold = int(cfg.get("under_3000_threshold", 3000))

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            dest = Path(tmp_dir) / f"scan_input{ext}"
            file.save(dest)
            res = scan_document(dest, gemini_api_key=gemini_key, pdf_password=pdf_password, threshold=threshold)
        inv = res.invoice
        return jsonify({
            "ok": True,
            "engine": res.engine,
            "cpsat_verified": res.cpsat_verified,
            "document_type": res.document_type,
            "recommended_sheet": res.recommended_sheet,
            "raw_text": res.raw_text,
            "invoice": {
                "invoice_no": inv.invoice_no,
                "roc_date": inv.roc_date,
                "seller_name": inv.seller_name,
                "total_amount": inv.total_amount,
                "items": [
                    {
                        "name": it.name,
                        "spec": it.spec,
                        "unit": it.unit,
                        "quantity": it.quantity,
                        "unit_price": it.unit_price,
                        "subtotal": it.subtotal,
                        "note": it.note,
                    }
                    for it in inv.items
                ],
            },
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": f"影像掃描失敗：{exc}"}), 400


@app.post("/api/parse-invoice-text")
def api_parse_invoice_text():
    try:
        data = request.get_json(force=True)
        inv = parse_invoice_text(data.get("text", ""))
        return jsonify({
            "ok": True,
            "invoice": {
                "invoice_no": inv.invoice_no,
                "roc_date": inv.roc_date,
                "seller_name": inv.seller_name,
                "total_amount": inv.total_amount,
                "items": [
                    {
                        "name": it.name,
                        "spec": it.spec,
                        "unit": it.unit,
                        "quantity": it.quantity,
                        "unit_price": it.unit_price,
                        "subtotal": it.subtotal,
                        "note": it.note,
                    }
                    for it in inv.items
                ],
            },
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


def _requests_from_edited(raw_reqs: list[dict[str, Any]], payload: dict[str, Any]) -> list[PurchaseRequest]:
    """Build PurchaseRequest objects directly from user-edited split preview cards."""
    plan_fallback = str(payload.get("plan_name", "")).strip()
    lv1_fallback = str(payload.get("level1", "")).strip()
    lv2_fallback = str(payload.get("level2", "")).strip()
    lv3_fallback = str(payload.get("level3", "")).strip()

    built: list[PurchaseRequest] = []
    for idx, r in enumerate(raw_reqs, start=1):
        items: list[LineItem] = []
        for row in r.get("items", []):
            name = str(row.get("name", "")).strip()
            if not name:
                continue
            qty = parse_number(str(row.get("quantity", ""))) if row.get("quantity") not in (None, "") else None
            price = parse_number(str(row.get("unit_price", ""))) if row.get("unit_price") not in (None, "") else None
            sub = parse_number(str(row.get("subtotal", ""))) if row.get("subtotal") not in (None, "") else None
            items.append(
                LineItem(
                    name=name,
                    spec=str(row.get("spec", "")).strip(),
                    unit=str(row.get("unit", "")).strip(),
                    quantity=qty,
                    unit_price=price,
                    manual_subtotal=sub if price is None else None,
                    note=str(row.get("note", "")).strip(),
                )
            )
        if not items:
            continue
        rc_raw = parse_number(str(r.get("receipt_count", ""))) if r.get("receipt_count") not in (None, "") else None
        built.append(
            PurchaseRequest(
                request_id=str(r.get("request_id") or f"REQ-{idx:03d}").strip(),
                source_type=str(r.get("source_type") or "明細"),
                budget_code=normalize_budget_code(str(r.get("budget_code", ""))),
                plan_name=str(r.get("plan_name") or plan_fallback).strip(),
                level1=str(r.get("level1") or lv1_fallback).strip(),
                level2=str(r.get("level2") or lv2_fallback).strip(),
                level3=str(r.get("level3") or lv3_fallback).strip(),
                department=str(r.get("department", "")).strip(),
                request_date=str(r.get("request_date", "")).strip(),
                payer_name=str(r.get("payer_name", "")).strip(),
                receipt_count=int(rc_raw) if rc_raw is not None else None,
                purpose=str(r.get("purpose", "")).strip(),
                items=items,
            )
        )
    if not built:
        raise ValueError("請至少保留一筆有效的品項明細")
    return built


def _prepare_requests_from_payload(payload: dict[str, Any], cfg: dict[str, Any]) -> list[PurchaseRequest]:
    """Build and split/chunk PurchaseRequest list according to payload options."""
    edited = payload.get("edited_requests")
    if isinstance(edited, list) and edited:
        return _requests_from_edited(edited, payload)

    reqs = _requests_from_payload(payload, cfg)
    max_rows = int(cfg["items_table"]["max_rows"])
    if payload.get("auto_split"):
        max_amount = float(payload.get("max_amount") or 3000)
        return split_requests_opt(reqs, max_amount=max_amount, max_rows=max_rows)
    return chunk_requests_by_max_rows(reqs, max_rows=max_rows)


@app.post("/api/solve-ocr-text")
def api_solve_ocr_text():
    """Re-parse edited raw OCR/PDF text with OR-Tools CP-SAT reconciliation and form splitting."""
    try:
        payload = request.get_json(force=True)
        raw_text = str(payload.get("raw_text", "")).strip()
        if not raw_text:
            raise ValueError("請提供原始 OCR / PDF 文字內容")

        cfg = load_config(CONFIG_PATH)
        max_rows = int(cfg["items_table"]["max_rows"])
        auto_split = bool(payload.get("auto_split", True))
        max_amount = float(payload.get("max_amount") or 3000)

        parsed = parse_invoice_text(raw_text)
        candidates = extract_candidates_from_text(raw_text, parsed)
        selected_items, verified = reconcile_items_with_cpsat(candidates, parsed.total_amount)
        if selected_items:
            parsed.items = selected_items
        if parsed.total_amount is None and parsed.items:
            tot = sum(it.subtotal for it in parsed.items)
            parsed.total_amount = int(tot) if float(tot).is_integer() else tot

        reqs = invoices_to_requests(
            [parsed],
            budget_code=str(payload.get("budget_code", "")),
            department=str(payload.get("department", "")),
            payer_name=str(payload.get("payer_name", "")),
            purpose=str(payload.get("purpose", "")),
            max_rows=max_rows,
            summarize_over_max=not auto_split,
        )
        for r in reqs:
            if payload.get("plan_name"):
                r.plan_name = str(payload["plan_name"]).strip()
            if payload.get("level1"):
                r.level1 = str(payload["level1"]).strip()
            if payload.get("level2"):
                r.level2 = str(payload["level2"]).strip()
            if payload.get("level3"):
                r.level3 = str(payload["level3"]).strip()

        if auto_split:
            reqs = split_requests_opt(reqs, max_amount=max_amount, max_rows=max_rows)
        else:
            reqs = chunk_requests_by_max_rows(reqs, max_rows=max_rows)

        return jsonify({
            "ok": True,
            "cpsat_verified": verified,
            "invoice": {
                "invoice_no": parsed.invoice_no,
                "roc_date": parsed.roc_date,
                "seller_name": parsed.seller_name,
                "total_amount": parsed.total_amount,
                "items": [
                    {
                        "name": it.name,
                        "spec": it.spec,
                        "unit": it.unit,
                        "quantity": it.quantity,
                        "unit_price": it.unit_price,
                        "subtotal": it.subtotal,
                        "note": it.note,
                    }
                    for it in parsed.items
                ],
            },
            "requests": _serialize_requests(reqs),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/preview")
def api_preview():
    try:
        payload = request.get_json(force=True)
        cfg = load_config(CONFIG_PATH)
        reqs = _prepare_requests_from_payload(payload, cfg)
        return jsonify({"ok": True, "requests": _serialize_requests(reqs)})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/generate")
def api_generate():
    try:
        tpl = _active_template_path()
        payload = request.get_json(force=True)
        cfg = load_config(CONFIG_PATH)
        reqs = _prepare_requests_from_payload(payload, cfg)

        files = generate_from_requests(tpl, reqs, OUTPUT_DIR, cfg)
        append_history(HISTORY_PATH, reqs, files)
        return jsonify({
            "ok": True,
            "files": [f.name for f in files],
            "requests": _serialize_requests(reqs),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/api/preview-file/<filename>")
def api_preview_file(filename: str):
    preview_dir = OUTPUT_DIR / "_preview"
    target = (preview_dir / filename).resolve()
    if target.parent != preview_dir.resolve() or not target.exists():
        return jsonify({"ok": False, "error": "找不到預覽檔案"}), 404
    if target.suffix.lower() == ".pdf":
        return send_file(target, mimetype="application/pdf", as_attachment=False)
    return send_file(target)


@app.post("/api/print-preview")
def api_print_preview():
    try:
        tpl = _active_template_path()
        payload = request.get_json(force=True)
        cfg = load_config(CONFIG_PATH)
        reqs = _prepare_requests_from_payload(payload, cfg)

        preview_dir = OUTPUT_DIR / "_preview"
        preview_dir.mkdir(parents=True, exist_ok=True)
        files = generate_from_requests(tpl, reqs, preview_dir, cfg)
        first_xlsx = files[0]
        pdf_path = preview_dir / f"{first_xlsx.stem}.pdf"

        if export_xlsx_to_pdf(first_xlsx, pdf_path):
            pngs = render_pdf_to_pngs(pdf_path, preview_dir, prefix=first_xlsx.stem)
            return jsonify({
                "ok": True,
                "mode": "excel_pdf",
                "pdf_url": f"/api/preview-file/{pdf_path.name}",
                "page_urls": [f"/api/preview-file/{p.name}" for p in pngs],
                "total_forms": len(files),
            })

        return jsonify({
            "ok": True,
            "mode": "html_fallback",
            "html": render_xlsx_fallback_html(first_xlsx),
            "total_forms": len(files),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": f"列印預覽失敗：{exc}"}), 400


@app.get("/api/download/<filename>")
def api_download(filename: str):
    target = (OUTPUT_DIR / filename).resolve()
    if target.parent != OUTPUT_DIR.resolve() or not target.exists():
        return jsonify({"ok": False, "error": "找不到檔案"}), 404
    return send_file(target, as_attachment=True)


@app.get("/api/download-zip")
def api_download_zip():
    names = [n for n in request.args.get("files", "").split(",") if n]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            target = (OUTPUT_DIR / name).resolve()
            if target.parent == OUTPUT_DIR.resolve() and target.exists():
                zf.write(target, arcname=target.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="採購申請單_批次產出.zip")


@app.get("/api/sample-csv")
def api_sample_csv():
    sample = BASE_DIR / "requests_sample.csv"
    return send_file(sample, as_attachment=True, download_name="requests_sample.csv")


@app.get("/api/history")
def api_get_history():
    records = load_history(HISTORY_PATH)
    return jsonify({
        "ok": True,
        "records": records,
        "summary": summarize_history(records),
    })


@app.delete("/api/history/<record_id>")
def api_delete_history(record_id: str):
    deleted = delete_history_record(HISTORY_PATH, record_id)
    if not deleted:
        return jsonify({"ok": False, "error": "找不到該筆紀錄"}), 404
    records = load_history(HISTORY_PATH)
    return jsonify({"ok": True, "records": records, "summary": summarize_history(records)})


@app.post("/api/history/clear")
def api_clear_history():
    clear_history(HISTORY_PATH)
    return jsonify({"ok": True, "records": [], "summary": summarize_history([])})


@app.get("/api/history/export")
def api_export_history():
    records = load_history(HISTORY_PATH)
    csv_bytes = export_history_csv_bytes(records)
    return send_file(
        io.BytesIO(csv_bytes),
        as_attachment=True,
        download_name="採購紀錄累積彙整.csv",
        mimetype="text/csv; charset=utf-8",
    )


@app.post("/api/open-output")
def api_open_output():
    if _is_cloud_mode():
        return jsonify({"ok": False, "error": "雲端模式下無法開啟本機資料夾"}), 403
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(OUTPUT_DIR))  # type: ignore[attr-defined]
    return jsonify({"ok": True})


@app.post("/api/shutdown")
def api_shutdown():
    if _is_cloud_mode():
        return jsonify({"ok": False, "error": "雲端模式下不允許遠端關閉伺服器"}), 403
    threading.Timer(0.3, lambda: os._exit(0)).start()
    return jsonify({"ok": True})


def _find_free_port(preferred: int = 5050) -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", preferred)) != 0:
            return preferred
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def _open_desktop_window(url: str) -> None:
    """Open a standalone Edge/Chrome app window on Windows, or fall back to default browser."""
    time.sleep(0.8)
    for exe in (
        shutil.which("msedge"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        shutil.which("chrome"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    ):
        if exe and Path(exe).exists():
            subprocess.Popen([exe, f"--app={url}", "--window-size=1180,860"])
            return
    webbrowser.open(url)


def run_desktop_app() -> None:
    env_port = os.environ.get("PORT")
    if env_port:
        port = int(env_port)
        print(f"採購單自動化系統雲端模式已啟動：0.0.0.0:{port} (Waitress WSGI)")
        serve(app, host="0.0.0.0", port=port, threads=4)
        return

    port = _find_free_port(5050)
    url = f"http://127.0.0.1:{port}"
    print(f"採購單自動化桌面系統已啟動：{url} (Waitress WSGI)")
    threading.Thread(target=_open_desktop_window, args=(url,), daemon=True).start()
    serve(app, host="127.0.0.1", port=port, threads=4)


if __name__ == "__main__":
    run_desktop_app()

