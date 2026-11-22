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
import uuid
import webbrowser
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from flask import Flask, jsonify, render_template, request, send_file, session
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
    sanitize_filename,
    split_requests_opt,
)
from travel_distance import calculate_route_mileage, optimize_route_with_cpsat, parse_stops_text
from substitute_builder import (
    DEFAULT_OCC_RATE,
    SubstituteEntry,
    SubstituteRoster,
    build_substitute_roster_xlsx,
    summarize_roster,
)
from travel_form_builder import (
    RouteSegment,
    TravelDayEntry,
    TravelRateConfig,
    TravelReport,
    amount_to_chinese_upper,
    build_travel_report_xlsx,
    compute_day_amounts,
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
from invoice_parser import parse_einvoice_csv, parse_invoice_text
from ocr_cpsat_parser import extract_candidates_from_text, reconcile_items_with_cpsat
from ocr_scanner import _detect_document_type, _recommend_sheet, parse_travel_receipt, scan_document
from print_preview import (
    export_multiple_xlsx_to_merged_pdf,
    export_xlsx_to_pdf,
    render_pdf_to_pngs,
    render_xlsx_fallback_html,
)

BASE_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = BASE_DIR / "config.json"
HISTORY_PATH = BASE_DIR / "history.json"
SUBSTITUTE_HISTORY_PATH = BASE_DIR / "substitute_history.json"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(24)

_SID_RE = re.compile(r"^[a-f0-9]{16}$")
_PREVIEW_TTL_SEC = 1800


def _get_session_upload_dir(*, create: bool = False) -> Path:
    """Return an isolated upload directory for the current browser session."""
    sid = session.get("sid", "")
    if not isinstance(sid, str) or not _SID_RE.match(sid):
        if not create:
            return UPLOAD_DIR
        sid = uuid.uuid4().hex[:16]
        session["sid"] = sid
    target = UPLOAD_DIR / sid
    if create:
        target.mkdir(parents=True, exist_ok=True)
    return target


def _active_template_path() -> Path | None:
    session_tpl = _get_session_upload_dir(create=False) / "template.xlsx"
    if session_tpl.exists():
        return session_tpl
    uploaded = UPLOAD_DIR / "template.xlsx"
    if uploaded.exists():
        return uploaded
    local = BASE_DIR / "template.xlsx"
    if local.exists():
        return local
    return None


def _collect_budget_options() -> dict[str, Any]:
    """Collect budget code mappings & dropdown options from every worksheet in the template."""
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

    tpl = _active_template_path()
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

    return {
        "entries": ordered,
        "plan_names": _uniq([e["plan_name"] for e in ordered] + extra_plans),
        "level1_options": _uniq([e["level1"] for e in ordered] + extra_lv1),
        "level2_options": _uniq([e["level2"] for e in ordered] + extra_lv2),
        "level3_options": _uniq([e["level3"] for e in ordered] + extra_lv3),
        "departments": _uniq(extra_depts),
        "payers": _uniq(extra_payers),
        "purposes": _uniq(extra_purposes),
    }


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
        uploaded_csv = _get_session_upload_dir(create=False) / "input.csv"
        if not uploaded_csv.exists():
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
        is_cloud=bool(os.environ.get("RENDER") or os.environ.get("PORT")),
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
    dest = _get_session_upload_dir(create=True) / "template.xlsx"
    file.save(dest)
    return jsonify({"ok": True, "filename": file.filename})


@app.post("/api/upload-csv")
def upload_csv():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "請選擇 CSV 檔案"}), 400
    dest = _get_session_upload_dir(create=True) / "input.csv"
    file.save(dest)
    return jsonify({"ok": True, "filename": file.filename})


@app.get("/api/inspect-template")
def api_inspect_template():
    tpl = _active_template_path()
    if not tpl:
        return jsonify({"ok": False, "error": "尚未上傳或放置 template.xlsx 範本檔"}), 400
    wb = openpyxl.load_workbook(tpl, data_only=False)
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


@app.post("/api/parse-raw-ocr-cpsat")
def api_parse_raw_ocr_cpsat():
    try:
        data = request.get_json(force=True)
        raw_text = str(data.get("raw_text", "")).strip()
        if not raw_text:
            raise ValueError("請先輸入或貼上原始 OCR / PDF 文字")

        inv = parse_invoice_text(raw_text)
        total_raw = parse_number(str(data.get("total_amount", "")))
        if total_raw is not None:
            inv.total_amount = total_raw

        candidates = extract_candidates_from_text(raw_text, inv)
        verified = False
        if candidates:
            selected, verified = reconcile_items_with_cpsat(candidates, inv.total_amount)
            if selected:
                inv.items = selected
                if inv.total_amount is None:
                    tot = sum(i.subtotal for i in selected)
                    inv.total_amount = int(tot) if float(tot).is_integer() else tot

        cfg = load_config(CONFIG_PATH)
        threshold = int(cfg.get("under_3000_threshold", 3000))
        return jsonify({
            "ok": True,
            "engine": "raw-text+cpsat",
            "cpsat_verified": verified,
            "document_type": _detect_document_type(raw_text),
            "recommended_sheet": _recommend_sheet(inv.total_amount, threshold=threshold),
            "raw_text": raw_text,
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


@app.post("/api/preview")
def api_preview():
    try:
        payload = request.get_json(force=True)
        cfg = load_config(CONFIG_PATH)
        reqs = _requests_from_payload(payload, cfg)
        max_rows = int(cfg["items_table"]["max_rows"])
        if payload.get("auto_split"):
            max_amount = float(payload.get("max_amount") or 3000)
            reqs = split_requests_opt(reqs, max_amount=max_amount, max_rows=max_rows)
        else:
            reqs = chunk_requests_by_max_rows(reqs, max_rows=max_rows)
        return jsonify({"ok": True, "requests": _serialize_requests(reqs)})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/generate")
def api_generate():
    try:
        tpl = _active_template_path()
        payload = request.get_json(force=True)
        cfg = load_config(CONFIG_PATH)
        reqs = _requests_from_payload(payload, cfg)
        max_rows = int(cfg["items_table"]["max_rows"])
        if payload.get("auto_split"):
            max_amount = float(payload.get("max_amount") or 3000)
            reqs = split_requests_opt(reqs, max_amount=max_amount, max_rows=max_rows)
        else:
            reqs = chunk_requests_by_max_rows(reqs, max_rows=max_rows)

        files = generate_from_requests(tpl, reqs, OUTPUT_DIR, cfg)
        history_entries = append_history(HISTORY_PATH, reqs, files)
        return jsonify({
            "ok": True,
            "files": [f.name for f in files],
            "requests": _serialize_requests(reqs),
            "history_entries": history_entries,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


def _cleanup_old_previews(preview_root: Path, max_age_sec: int = _PREVIEW_TTL_SEC) -> None:
    """Remove expired temporary preview files and subdirectories."""
    if not preview_root.exists():
        return
    cutoff = time.time() - max_age_sec
    for entry in preview_root.iterdir():
        try:
            if entry.stat().st_mtime < cutoff:
                if entry.is_dir():
                    shutil.rmtree(entry, ignore_errors=True)
                else:
                    entry.unlink(missing_ok=True)
        except OSError:
            pass


@app.get("/api/preview-file/<path:subpath>")
def api_preview_file(subpath: str):
    preview_root = (OUTPUT_DIR / "_preview").resolve()
    target = (preview_root / subpath).resolve()
    if not target.is_relative_to(preview_root) or not target.is_file():
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
        reqs = _requests_from_payload(payload, cfg)
        max_rows = int(cfg["items_table"]["max_rows"])
        if payload.get("auto_split"):
            max_amount = float(payload.get("max_amount") or 3000)
            reqs = split_requests_opt(reqs, max_amount=max_amount, max_rows=max_rows)
        else:
            reqs = chunk_requests_by_max_rows(reqs, max_rows=max_rows)

        preview_root = OUTPUT_DIR / "_preview"
        _cleanup_old_previews(preview_root)
        preview_id = uuid.uuid4().hex[:12]
        preview_dir = preview_root / preview_id
        preview_dir.mkdir(parents=True, exist_ok=True)
        files = generate_from_requests(tpl, reqs, preview_dir, cfg)
        merged_stem = files[0].stem if len(files) == 1 else "採購申請單_合併列印"
        pdf_path = preview_dir / f"{merged_stem}.pdf"

        if export_multiple_xlsx_to_merged_pdf(files, pdf_path):
            pngs = render_pdf_to_pngs(pdf_path, preview_dir, prefix=merged_stem)
            return jsonify({
                "ok": True,
                "mode": "excel_pdf",
                "pdf_url": f"/api/preview-file/{preview_id}/{pdf_path.name}",
                "page_urls": [f"/api/preview-file/{preview_id}/{p.name}" for p in pngs],
                "total_forms": len(files),
            })

        combined_html = '<div style="page-break-after:always; margin:18px 0; border-top:2px dashed #cbd5e1;"></div>'.join(
            render_xlsx_fallback_html(f) for f in files
        )
        return jsonify({
            "ok": True,
            "mode": "html_fallback",
            "html": combined_html,
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


@app.post("/api/history/redownload")
def api_history_redownload():
    """Rebuild a purchase Excel from saved history request_data when Render temp file is gone."""
    try:
        payload = request.get_json(force=True)
        rec = payload.get("record") or {}
        rd = rec.get("request_data") or {}
        if not rd or not rd.get("items"):
            return jsonify({"ok": False, "error": "此筆紀錄缺少明細資料，無法於雲端重啟後自動重建"}), 404
        cfg = load_config(CONFIG_PATH)
        if isinstance(payload.get("custom_config"), dict):
            cfg.update(payload["custom_config"])
        items = [
            LineItem(
                name=str(it.get("name", "")),
                spec=str(it.get("spec", "")),
                unit=str(it.get("unit", "")),
                quantity=parse_number(it.get("quantity")),
                unit_price=parse_number(it.get("unit_price")),
                manual_subtotal=parse_number(it.get("manual_subtotal") if it.get("manual_subtotal") is not None else it.get("subtotal")),
                invoice_no=str(it.get("invoice_no", "")),
                note=str(it.get("note", "")),
            )
            for it in rd.get("items", [])
        ]
        req = PurchaseRequest(
            request_id=str(rd.get("request_id") or rec.get("request_id") or "REQ-001"),
            source_type=str(rd.get("source_type") or "明細"),
            sheet_name=str(rd.get("sheet_name") or ""),
            budget_code=str(rd.get("budget_code") or ""),
            plan_name=str(rd.get("plan_name") or ""),
            level1=str(rd.get("level1") or ""),
            level2=str(rd.get("level2") or ""),
            level3=str(rd.get("level3") or ""),
            department=str(rd.get("department") or ""),
            payer_name=str(rd.get("payer_name") or ""),
            purpose=str(rd.get("purpose") or ""),
            receipt_count=int(rd["receipt_count"]) if rd.get("receipt_count") not in (None, "") else None,
            request_date=str(rd.get("request_date") or ""),
            items=items,
        )
        tpl = _active_template_path()
        files = generate_from_requests(tpl, [req], OUTPUT_DIR, cfg)
        if not files or not files[0].exists():
            return jsonify({"ok": False, "error": "重建 Excel 失敗"}), 500
        dl_name = str(rec.get("output_file") or files[0].name)
        return send_file(files[0], as_attachment=True, download_name=dl_name)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/open-output")
def api_open_output():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(OUTPUT_DIR))  # type: ignore[attr-defined]
    return jsonify({"ok": True})


def _parse_segments_text(raw: str) -> list[RouteSegment]:
    """Parse 'A->B:14, B->A:14' or 'A→B:14' into RouteSegment list."""
    segments: list[RouteSegment] = []
    for part in re.split(r"[,;\n]+", raw.strip()):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(.*?)(?:->|→)(.*?)\s*[:：]\s*(\d+(?:\.\d+)?)\s*(?:km|公里)?$", part, re.IGNORECASE)
        if m:
            segments.append(
                RouteSegment(
                    origin=m.group(1).strip(),
                    destination=m.group(2).strip(),
                    distance_km=float(m.group(3)),
                )
            )
    return segments


def _travel_report_from_payload(payload: dict[str, Any]) -> TravelReport:
    raw_days = payload.get("days", [])
    if not raw_days:
        raise ValueError("請至少填寫一筆出差行程明細")

    car_rate = parse_number(str(payload.get("car_rate_per_km", "3.0")))
    scooter_rate = parse_number(str(payload.get("scooter_rate_per_km", "2.0")))
    rates = TravelRateConfig(
        car_rate_per_km=float(car_rate) if car_rate is not None and float(car_rate) > 0 else 3.0,
        scooter_rate_per_km=float(scooter_rate) if scooter_rate is not None and float(scooter_rate) > 0 else 2.0,
    )

    days: list[TravelDayEntry] = []
    for row in raw_days:
        mode = str(row.get("transport_mode", "self_car")).strip() or "self_car"
        seg_text = str(row.get("segments_text", "")).strip()
        segments = _parse_segments_text(seg_text) if seg_text else []
        for s in row.get("segments", []) or []:
            segments.append(
                RouteSegment(
                    origin=str(s.get("origin", "")).strip(),
                    destination=str(s.get("destination", "")).strip(),
                    distance_km=float(parse_number(str(s.get("distance_km", "0"))) or 0),
                )
            )

        fare_val = float(parse_number(str(row.get("fare", "0"))) or 0)
        bus_equiv = parse_number(str(row.get("bus_equivalent_fare", "")))
        one_way = parse_number(str(row.get("one_way_km", "")))
        misc_raw = parse_number(str(row.get("misc_fee", "")))

        days.append(
            TravelDayEntry(
                date=str(row.get("date", "")).strip(),
                route_text=str(row.get("route_text", "")).strip(),
                summary=str(row.get("summary", "")).strip(),
                transport_mode=mode,
                segments=segments,
                one_way_km=float(one_way) if one_way is not None else None,
                bus_equivalent_fare=float(bus_equiv) if bus_equiv is not None else None,
                train_fare=fare_val if mode == "train" else float(parse_number(str(row.get("train_fare", "0"))) or 0),
                flight_hsr_fare=fare_val if mode == "flight_hsr" else float(parse_number(str(row.get("flight_hsr_fare", "0"))) or 0),
                bus_mrt_fare=fare_val if mode in ("bus_mrt", "taxi") else float(parse_number(str(row.get("bus_mrt_fare", "0"))) or 0),
                ship_fare=fare_val if mode == "ship" else float(parse_number(str(row.get("ship_fare", "0"))) or 0),
                ubike_fare=fare_val if mode == "ubike" else float(parse_number(str(row.get("ubike_fare", "0"))) or 0),
                self_car_fare=fare_val if mode == "self_car" else float(parse_number(str(row.get("self_car_fare", "0"))) or 0),
                self_scooter_fare=fare_val if mode == "self_scooter" else float(parse_number(str(row.get("self_scooter_fare", "0"))) or 0),
                taxi_approved=bool(row.get("taxi_approved", False)),
                is_holiday_lodging=bool(row.get("is_holiday_lodging", False)),
                lodging_actual=float(parse_number(str(row.get("lodging_actual", "0"))) or 0),
                misc_fee=float(misc_raw) if misc_raw is not None else None,
                receipt_count=int(parse_number(str(row.get("receipt_count", "0"))) or 0),
                note=str(row.get("note", "")).strip(),
            )
        )

    return TravelReport(
        school_name=str(payload.get("school_name", "")).strip() or "臺南市立土城高級中學",
        traveler_name=str(payload.get("traveler_name", "")).strip(),
        job_title=str(payload.get("job_title", "")).strip(),
        rank_level=str(payload.get("rank_level", "簡任級以下")).strip() or "簡任級以下",
        purpose=str(payload.get("purpose", "")).strip(),
        destination=str(payload.get("destination", "")).strip(),
        start_date=str(payload.get("start_date", "")).strip(),
        end_date=str(payload.get("end_date", "")).strip(),
        total_days_text=str(payload.get("total_days_text", "")).strip(),
        budget_code=normalize_budget_code(str(payload.get("budget_code", ""))),
        plan_name=str(payload.get("plan_name", "")).strip(),
        days=days,
        rates=rates,
    )


@app.post("/api/travel/scan-receipt")
def api_travel_scan_receipt():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "請選擇車票或住宿單據檔案（圖片或 PDF）"}), 400
    ext = Path(file.filename).suffix.lower()
    if ext not in (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".pdf"):
        return jsonify({"ok": False, "error": "僅支援圖片 (.jpg, .png, .webp, .bmp) 或 .pdf 格式"}), 400

    gemini_key = request.form.get("gemini_api_key", "")
    pdf_password = request.form.get("pdf_password", "")
    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            dest = Path(tmp_dir) / f"travel_receipt{ext}"
            file.save(dest)
            parsed = parse_travel_receipt(dest, gemini_api_key=gemini_key, pdf_password=pdf_password)
        return jsonify({"ok": True, **parsed})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"單據掃描失敗：{exc}"}), 400


@app.post("/api/travel/calc-route")
def api_travel_calc_route():
    try:
        payload = request.get_json(force=True)
        raw_stops = payload.get("stops")
        if isinstance(raw_stops, str):
            stops = parse_stops_text(raw_stops)
        elif isinstance(raw_stops, list):
            stops = [str(s).strip() for s in raw_stops if str(s).strip()]
        else:
            stops = parse_stops_text(str(payload.get("route_text", "")))

        round_trip = bool(payload.get("round_trip", False))
        optimize_order = bool(payload.get("optimize_order", False))
        if optimize_order:
            result = optimize_route_with_cpsat(stops, round_trip=round_trip)
        else:
            result = calculate_route_mileage(stops, round_trip=round_trip)
        return jsonify({"ok": True, **result})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/travel/preview")
def api_travel_preview():
    try:
        payload = request.get_json(force=True)
        report = _travel_report_from_payload(payload)
        is_same_day = bool(report.start_date and report.start_date == report.end_date) or len(report.days) == 1
        computed = [
            compute_day_amounts(d, rates=report.rates, is_same_day_round_trip=is_same_day)
            for d in report.days
        ]
        grand_total = sum(c.subtotal for c in computed)
        return jsonify({
            "ok": True,
            "days": [
                {
                    "date": c.date,
                    "route_text": c.route_text,
                    "summary": c.summary,
                    "train_fare": c.train_fare,
                    "flight_hsr_fare": c.flight_hsr_fare,
                    "bus_mrt_fare": c.bus_mrt_fare,
                    "ship_fare": c.ship_fare,
                    "ubike_fare": c.ubike_fare,
                    "self_car_fare": c.self_car_fare,
                    "self_scooter_fare": c.self_scooter_fare,
                    "lodging_fee": c.lodging_fee,
                    "misc_fee": c.misc_fee,
                    "receipt_count": c.receipt_count,
                    "subtotal": c.subtotal,
                    "note": c.note,
                }
                for c in computed
            ],
            "grand_total": grand_total,
            "chinese_upper": amount_to_chinese_upper(grand_total),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/travel/generate")
def api_travel_generate():
    try:
        payload = request.get_json(force=True)
        report = _travel_report_from_payload(payload)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        safe_name = sanitize_filename(report.traveler_name or "出差人員")
        stamp = time.strftime("%Y%m%d_%H%M%S")
        out_path = OUTPUT_DIR / f"國內出差旅費報告表_{safe_name}_{stamp}.xlsx"
        build_travel_report_xlsx(report, out_path)
        is_same_day = bool(report.start_date and report.start_date == report.end_date) or len(report.days) == 1
        computed = [
            compute_day_amounts(d, rates=report.rates, is_same_day_round_trip=is_same_day)
            for d in report.days
        ]
        grand_total = sum(c.subtotal for c in computed)
        return jsonify({
            "ok": True,
            "filename": out_path.name,
            "grand_total": grand_total,
            "chinese_upper": amount_to_chinese_upper(grand_total),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/travel/print-preview")
def api_travel_print_preview():
    try:
        payload = request.get_json(force=True)
        report = _travel_report_from_payload(payload)
        preview_root = OUTPUT_DIR / "_preview"
        _cleanup_old_previews(preview_root)
        preview_id = uuid.uuid4().hex[:12]
        preview_dir = preview_root / preview_id
        preview_dir.mkdir(parents=True, exist_ok=True)
        xlsx_path = preview_dir / "國內出差旅費報告表.xlsx"
        build_travel_report_xlsx(report, xlsx_path)
        pdf_path = preview_dir / "國內出差旅費報告表.pdf"

        if export_xlsx_to_pdf(xlsx_path, pdf_path):
            pngs = render_pdf_to_pngs(pdf_path, preview_dir, prefix=xlsx_path.stem)
            return jsonify({
                "ok": True,
                "mode": "excel_pdf",
                "pdf_url": f"/api/preview-file/{preview_id}/{pdf_path.name}",
                "page_urls": [f"/api/preview-file/{preview_id}/{p.name}" for p in pngs],
                "total_forms": 1,
            })

        return jsonify({
            "ok": True,
            "mode": "html_fallback",
            "html": render_xlsx_fallback_html(xlsx_path),
            "total_forms": 1,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": f"列印預覽失敗：{exc}"}), 400


def _substitute_roster_from_payload(payload: dict[str, Any]) -> SubstituteRoster:
    raw_entries = payload.get("entries") or []
    entries: list[SubstituteEntry] = []
    for row in raw_entries:
        name = str(row.get("teacher_name", "")).strip()
        if not name:
            continue
        entries.append(
            SubstituteEntry(
                teacher_name=name,
                id_number=str(row.get("id_number", "")).strip(),
                subject_note=str(row.get("subject_note", "")).strip(),
                school_level=str(row.get("school_level", "elementary")),
                hourly_rate=int(parse_number(row.get("hourly_rate")) or 0),
                periods=float(parse_number(row.get("periods")) or 0),
                enroll_mode=str(row.get("enroll_mode", "part_time")),
                insured_days=int(parse_number(row.get("insured_days")) if row.get("insured_days") not in (None, "") else 30),
                dependents=int(parse_number(row.get("dependents")) or 0),
                voluntary_pension_pct=float(parse_number(row.get("voluntary_pension_pct")) or 0),
                nhi_end_of_month=bool(row.get("nhi_end_of_month", True)),
                nhi_enrolled_here=bool(row.get("nhi_enrolled_here", False)),
                income_tax=int(parse_number(row.get("income_tax")) or 0),
                note=str(row.get("note", "")).strip(),
            )
        )
    if not entries:
        raise ValueError("請至少輸入一位代課教師資料")
    occ_raw = parse_number(payload.get("occ_rate"))
    occ_rate = (float(occ_raw) / 100.0) if occ_raw is not None else DEFAULT_OCC_RATE
    return SubstituteRoster(
        school_name=str(payload.get("school_name") or "臺南市立○○國民小學").strip(),
        roc_year=int(parse_number(payload.get("roc_year")) or 115),
        month=int(parse_number(payload.get("month")) or 11),
        plan_name=str(payload.get("plan_name") or "教師請假短期代課鐘點費").strip(),
        budget_code=str(payload.get("budget_code") or "").strip(),
        occ_rate=occ_rate,
        entries=entries,
    )


@app.post("/api/substitute/preview")
def api_substitute_preview():
    try:
        payload = request.get_json(force=True)
        roster = _substitute_roster_from_payload(payload)
        items, totals = summarize_roster(roster)
        return jsonify({
            "ok": True,
            "items": [
                {
                    "teacher_name": i.teacher_name,
                    "subject_note": i.subject_note,
                    "hourly_rate": i.hourly_rate,
                    "periods": i.periods,
                    "gross_pay": i.gross_pay,
                    "enroll_mode": i.enroll_mode,
                    "insured_days": i.insured_days,
                    "li_bracket": i.li_bracket,
                    "nhi_bracket": i.nhi_bracket,
                    "pension_bracket": i.pension_bracket,
                    "li_employee": i.li_employee,
                    "nhi_employee": i.nhi_employee,
                    "pension_employee": i.pension_employee,
                    "nhi_supp": i.nhi_supp,
                    "income_tax": i.income_tax,
                    "total_deduction": i.total_deduction,
                    "net_pay": i.net_pay,
                    "li_employer": i.li_employer,
                    "occ_employer": i.occ_employer,
                    "nhi_employer": i.nhi_employer,
                    "pension_employer": i.pension_employer,
                    "employer_total": i.employer_total,
                    "note": i.note,
                }
                for i in items
            ],
            "totals": totals,
            "chinese_upper_net": amount_to_chinese_upper(totals["net_pay"]),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


def _load_sub_history() -> list[dict[str, Any]]:
    if not SUBSTITUTE_HISTORY_PATH.exists():
        return []
    try:
        data = json.loads(SUBSTITUTE_HISTORY_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_sub_history(records: list[dict[str, Any]]) -> None:
    SUBSTITUTE_HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    SUBSTITUTE_HISTORY_PATH.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def _summarize_sub_history(records: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "total_count": len(records),
        "total_gross": sum(int(r.get("gross_pay") or 0) for r in records),
        "total_net": sum(int(r.get("net_pay") or 0) for r in records),
        "total_employer": sum(int(r.get("employer_total") or 0) for r in records),
    }


@app.post("/api/substitute/generate")
def api_substitute_generate():
    try:
        payload = request.get_json(force=True)
        roster = _substitute_roster_from_payload(payload)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        safe_plan = sanitize_filename(f"{roster.roc_year}年{roster.month}月_代課鐘點費印領清冊")
        stamp = time.strftime("%Y%m%d_%H%M%S")
        out_path = OUTPUT_DIR / f"{safe_plan}_{stamp}.xlsx"
        build_substitute_roster_xlsx(roster, out_path)
        _, totals = summarize_roster(roster)
        names = [e.teacher_name for e in roster.entries]
        summary_names = "、".join(names[:3]) + (f" 等共{len(names)}人" if len(names) > 3 else "")
        history_entry = {
            "id": uuid.uuid4().hex[:10],
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "school_name": roster.school_name,
            "roc_year": roster.roc_year,
            "month": roster.month,
            "plan_name": roster.plan_name,
            "budget_code": roster.budget_code,
            "teacher_count": len(names),
            "teachers_summary": summary_names,
            "gross_pay": totals["gross_pay"],
            "net_pay": totals["net_pay"],
            "employer_total": totals["employer_total"],
            "output_file": out_path.name,
            "payload": payload,
        }
        records = [history_entry] + _load_sub_history()
        _save_sub_history(records)
        return jsonify({
            "ok": True,
            "filename": out_path.name,
            "totals": totals,
            "chinese_upper_net": amount_to_chinese_upper(totals["net_pay"]),
            "history_entry": history_entry,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/api/substitute/history")
def api_substitute_history():
    records = _load_sub_history()
    return jsonify({"ok": True, "records": records, "summary": _summarize_sub_history(records)})


@app.delete("/api/substitute/history/<record_id>")
def api_substitute_history_delete(record_id: str):
    records = _load_sub_history()
    filtered = [r for r in records if r.get("id") != record_id]
    if len(filtered) == len(records):
        return jsonify({"ok": False, "error": "找不到該筆代課清冊紀錄"}), 404
    _save_sub_history(filtered)
    return jsonify({"ok": True, "records": filtered, "summary": _summarize_sub_history(filtered)})


@app.post("/api/substitute/history/clear")
def api_substitute_history_clear():
    _save_sub_history([])
    return jsonify({"ok": True, "records": [], "summary": _summarize_sub_history([])})


@app.post("/api/substitute/history/redownload")
def api_substitute_history_redownload():
    try:
        body = request.get_json(force=True)
        rec = body.get("record") or {}
        payload = rec.get("payload") or {}
        if not payload or not payload.get("entries"):
            return jsonify({"ok": False, "error": "此筆紀錄缺少明細資料，無法重建清冊"}), 404
        roster = _substitute_roster_from_payload(payload)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        dl_name = str(rec.get("output_file") or f"{roster.roc_year}年{roster.month}月_代課鐘點費印領清冊.xlsx")
        out_path = OUTPUT_DIR / sanitize_filename(dl_name.removesuffix(".xlsx"))
        if not str(out_path).endswith(".xlsx"):
            out_path = Path(f"{out_path}.xlsx")
        build_substitute_roster_xlsx(roster, out_path)
        return send_file(out_path, as_attachment=True, download_name=dl_name)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.post("/api/substitute/print-preview")
def api_substitute_print_preview():
    try:
        payload = request.get_json(force=True)
        roster = _substitute_roster_from_payload(payload)
        preview_root = OUTPUT_DIR / "_preview"
        _cleanup_old_previews(preview_root)
        preview_id = uuid.uuid4().hex[:12]
        preview_dir = preview_root / preview_id
        preview_dir.mkdir(parents=True, exist_ok=True)
        xlsx_path = preview_dir / "代課鐘點費印領清冊.xlsx"
        build_substitute_roster_xlsx(roster, xlsx_path)
        pdf_path = preview_dir / "代課鐘點費印領清冊.pdf"

        if export_xlsx_to_pdf(xlsx_path, pdf_path):
            pngs = render_pdf_to_pngs(pdf_path, preview_dir, prefix=xlsx_path.stem)
            return jsonify({
                "ok": True,
                "mode": "excel_pdf",
                "pdf_url": f"/api/preview-file/{preview_id}/{pdf_path.name}",
                "page_urls": [f"/api/preview-file/{preview_id}/{p.name}" for p in pngs],
                "total_forms": 1,
            })

        return jsonify({
            "ok": True,
            "mode": "html_fallback",
            "html": render_xlsx_fallback_html(xlsx_path),
            "total_forms": 1,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": f"列印預覽失敗：{exc}"}), 400


@app.post("/api/shutdown")
def api_shutdown():
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

