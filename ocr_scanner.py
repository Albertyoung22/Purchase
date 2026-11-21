"""Image/document OCR & Vision scanner for Taiwan invoices (發票) and quotations (估價單)."""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from invoice_parser import ParsedInvoice, ParsedInvoiceItem, format_roc_date, parse_invoice_text
from ocr_cpsat_parser import extract_candidates_from_text, reconcile_items_with_cpsat

_OCR_ENGINE: Any = None


@dataclass
class ScanResult:
    """Structured result returned from scanning an invoice or quotation image."""

    engine: str
    document_type: str  # "發票" or "估價單"
    raw_text: str
    invoice: ParsedInvoice
    recommended_sheet: str
    cpsat_verified: bool = False


def _get_rapid_ocr() -> Any:
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR

        try:
            _OCR_ENGINE = RapidOCR(intra_op_num_threads=1, inter_op_num_threads=1)
        except Exception:
            _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _prepare_ocr_image(image_path: Path, tmp_dir: Path, max_dim: int = 1600) -> Path:
    """Fix EXIF orientation and downscale oversized photos to prevent OOM/502 on 512MB cloud hosts."""
    from PIL import Image, ImageOps

    with Image.open(image_path) as img:
        img = ImageOps.exif_transpose(img)
        w, h = img.size
        if max(w, h) <= max_dim:
            return image_path
        scale = max_dim / float(max(w, h))
        new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
        resized = img.resize(new_size, Image.Resampling.LANCZOS)
        if resized.mode not in ("RGB", "L"):
            resized = resized.convert("RGB")
        out_path = tmp_dir / f"{image_path.stem}_resized.jpg"
        resized.save(out_path, format="JPEG", quality=88)
        return out_path


def _group_ocr_boxes_into_lines(ocr_result: list[Any], y_tolerance: float = 14.0) -> str:
    """Sort RapidOCR bounding boxes top-to-bottom, left-to-right and join same-row text."""
    if not ocr_result:
        return ""

    entries: list[tuple[float, float, str]] = []
    for item in ocr_result:
        box, text, _score = item
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        center_y = sum(ys) / len(ys)
        left_x = min(xs)
        entries.append((center_y, left_x, str(text).strip()))

    entries.sort(key=lambda e: (e[0], e[1]))

    lines: list[list[tuple[float, str]]] = []
    current_y: float | None = None

    for cy, lx, text in entries:
        if current_y is None or abs(cy - current_y) > y_tolerance:
            lines.append([(lx, text)])
            current_y = cy
        else:
            lines[-1].append((lx, text))

    joined_lines = []
    for row in lines:
        row.sort(key=lambda pair: pair[0])
        joined_lines.append("  ".join(text for _, text in row))
    return "\n".join(joined_lines)


def _detect_document_type(text: str) -> str:
    if any(kw in text for kw in ("估價單", "報價單", "報價", "估價")):
        return "估價單"
    return "發票"


def _recommend_sheet(total_amount: int | float | None, threshold: int = 3000) -> str:
    if total_amount is not None and total_amount < threshold:
        return "請購單簡化流程-3000元以下適用"
    return "請購單-3000元以上一定要用此表"


def _scan_with_gemini(image_path: Path, api_key: str) -> tuple[str, str, ParsedInvoice]:
    """Use Gemini Vision REST API to extract structured invoice/quotation fields."""
    mime_type, _ = mimetypes.guess_type(str(image_path))
    mime_type = mime_type or "image/jpeg"
    b64_data = base64.b64encode(image_path.read_bytes()).decode("ascii")

    prompt = (
        "你是台灣學校採購單據辨識專家。請辨識這張圖片（可能是電子發票、傳統發票、收據或廠商估價單），"
        "並只回傳純 JSON（不要 markdown 標記），格式如下：\n"
        "{\n"
        '  "document_type": "發票 或 估價單",\n'
        '  "invoice_no": "發票號碼如 AB-12345678，若無則空字串",\n'
        '  "date": "YYYY-MM-DD 或 YYY/MM/DD，若無則空字串",\n'
        '  "seller_name": "賣方或廠商名稱",\n'
        '  "total_amount": 數字或null,\n'
        '  "items": [\n'
        '    {"name": "品名（不含規格型號）", "spec": "規格/型號/尺寸/容量（若無則空字串）", "unit": "單位（如個/包/支/台，若無則空字串）", "quantity": 數量數字, "unit_price": 單價數字或null, "subtotal": 小計數字}\n'
        "  ]\n"
        "}"
    )

    body = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": mime_type, "data": b64_data}},
                ]
            }
        ],
        "generationConfig": {"response_mime_type": "application/json"},
    }

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        "gemini-2.5-flash:generateContent"
    )
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-goog-api-key": api_key,
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        resp_data = json.loads(resp.read().decode("utf-8"))

    raw_json = resp_data["candidates"][0]["content"]["parts"][0]["text"]
    parsed = json.loads(raw_json)

    items = [
        ParsedInvoiceItem(
            name=str(it.get("name", "")).strip(),
            spec=str(it.get("spec", "")).strip(),
            unit=str(it.get("unit", "")).strip(),
            quantity=it.get("quantity") or 1,
            unit_price=it.get("unit_price"),
            subtotal=it.get("subtotal") or 0,
        )
        for it in parsed.get("items", [])
        if str(it.get("name", "")).strip()
    ]
    inv = ParsedInvoice(
        invoice_no=str(parsed.get("invoice_no", "")).strip(),
        roc_date=format_roc_date(str(parsed.get("date", ""))),
        seller_name=str(parsed.get("seller_name", "")).strip(),
        total_amount=parsed.get("total_amount"),
        items=items,
    )
    if inv.total_amount is None and inv.items:
        inv.total_amount = sum(i.subtotal for i in inv.items)

    doc_type = "估價單" if parsed.get("document_type") == "估價單" else "發票"
    return doc_type, raw_json, inv


def _reconcile_with_cpsat(raw_text: str, inv: ParsedInvoice) -> tuple[ParsedInvoice, bool]:
    """Use OR-Tools CP-SAT to filter noise lines and reconcile items against total_amount."""
    candidates = extract_candidates_from_text(raw_text, inv)
    if not candidates:
        return inv, False
    selected, exact_match = reconcile_items_with_cpsat(candidates, inv.total_amount)
    if selected:
        inv.items = selected
        if inv.total_amount is None:
            total = sum(i.subtotal for i in selected)
            inv.total_amount = int(total) if float(total).is_integer() else total
    return inv, exact_match


def _extract_pdf_lines(doc: Any) -> str:
    """Extract PDF text preserving horizontal table rows via word bounding boxes."""
    page_texts: list[str] = []
    for page in doc:
        words = page.get_text("words")  # (x0, y0, x1, y1, word, block, line, word_no)
        if words:
            boxes = [([[w[0], w[1]], [w[2], w[1]], [w[2], w[3]], [w[0], w[3]]], w[4], 1.0) for w in words]
            page_texts.append(_group_ocr_boxes_into_lines(boxes, y_tolerance=6.0))
        else:
            t = page.get_text("text").strip()
            if t:
                page_texts.append(t)
    return "\n".join(t for t in page_texts if t.strip()).strip()


def _parse_num_str(raw: str) -> int | float | None:
    cleaned = raw.replace(",", "").replace("$", "").replace("NT", "").replace("元", "").strip()
    if not cleaned:
        return None
    try:
        v = float(cleaned)
        return int(v) if v.is_integer() else v
    except ValueError:
        return None


def _map_header_role(label: str) -> str | None:
    c = label.replace(" ", "").replace("\n", "").strip()
    if any(k in c for k in ("項次", "序號", "編號")) or c.upper() == "NO":
        return "index"
    if any(k in c for k in ("項目", "品名", "名稱", "商品")):
        return "name"
    if any(k in c for k in ("規格", "尺寸", "型號")):
        return "spec"
    if "顏色" in c:
        return "color"
    if "單位" in c:
        return "unit"
    if "單價" in c:
        return "unit_price"
    if "數量" in c:
        return "quantity"
    if any(k in c for k in ("金額", "小計", "複價")):
        return "subtotal"
    if any(k in c for k in ("備註", "說明")):
        return "note"
    return None


def _join_cell_words(words: list[tuple[float, float, str]]) -> str:
    """Sort words in a table cell by (line_y, x0) and join cleanly without breaking CJK text."""
    if not words:
        return ""
    words.sort(key=lambda w: (round(w[0] / 6.0), w[1]))
    parts = [w[2] for w in words]
    return "".join(parts).strip()


def _extract_pdf_table_invoice(doc: Any, combined_text: str) -> ParsedInvoice | None:
    """Extract structured items from bordered PDF quotation/invoice tables via header column X-coords."""
    base_inv = parse_invoice_text(combined_text)
    extracted_items: list[ParsedInvoiceItem] = []

    for page in doc:
        tabs = page.find_tables()
        if not tabs or not tabs.tables:
            continue

        words = page.get_text("words")  # (x0, y0, x1, y1, word, block, line, word_no)
        for table in tabs.tables:
            if not table.header or not table.header.cells or not table.header.names:
                continue

            col_ranges: list[tuple[float, float, str]] = []
            roles_present: set[str] = set()
            for cell_bbox, col_name in zip(table.header.cells, table.header.names):
                if cell_bbox is None or not col_name:
                    continue
                role = _map_header_role(str(col_name))
                if role:
                    col_ranges.append((cell_bbox[0], cell_bbox[2], role))
                    roles_present.add(role)

            if "name" not in roles_present or "subtotal" not in roles_present:
                continue

            header_y1 = table.header.bbox[3]
            table_y1 = table.bbox[3]

            # Collect words inside table body
            body_words = [
                w for w in words if w[1] >= header_y1 - 2 and w[3] <= table_y1 + 2
            ]

            # Find '合計' / '總計' row Y boundary so we don't treat it as an item
            total_y = table_y1
            for w in body_words:
                txt = w[4].replace(" ", "")
                if txt in ("合計", "總計", "總金額"):
                    total_y = min(total_y, w[1] - 3)

            body_words = [w for w in body_words if w[3] <= total_y + 2]

            def _col_role_for(wx0: float, wx1: float) -> str | None:
                cx = (wx0 + wx1) / 2.0
                for rx0, rx1, rname in col_ranges:
                    if rx0 - 2 <= cx <= rx1 + 2:
                        return rname
                return None

            # Determine row Y-anchors from 'index' column (項次 1, 2, 3...) or 'subtotal' column
            anchor_role = "index" if "index" in roles_present else "subtotal"
            anchors: list[float] = []
            for w in body_words:
                if _col_role_for(w[0], w[2]) == anchor_role:
                    if _parse_num_str(w[4]) is not None:
                        anchors.append(w[1])
            anchors.sort()

            # Deduplicate anchors that are very close vertically
            row_starts: list[float] = []
            for ay in anchors:
                if not row_starts or ay - row_starts[-1] > 8.0:
                    row_starts.append(ay)

            if not row_starts:
                continue

            for idx, y_start in enumerate(row_starts):
                y_end = row_starts[idx + 1] - 4.0 if idx + 1 < len(row_starts) else total_y
                y_min = y_start - 6.0

                col_buckets: dict[str, list[tuple[float, float, str]]] = {}
                for w in body_words:
                    cy = (w[1] + w[3]) / 2.0
                    if y_min <= cy < y_end:
                        role = _col_role_for(w[0], w[2])
                        if role:
                            col_buckets.setdefault(role, []).append((w[1], w[0], w[4]))

                name = _join_cell_words(col_buckets.get("name", []))
                sub_val = _parse_num_str(_join_cell_words(col_buckets.get("subtotal", [])))
                if not name and sub_val is None:
                    continue

                spec_parts = [
                    _join_cell_words(col_buckets.get("spec", [])),
                    _join_cell_words(col_buckets.get("color", [])),
                ]
                spec = " / ".join(p for p in spec_parts if p)
                unit = _join_cell_words(col_buckets.get("unit", []))
                qty = _parse_num_str(_join_cell_words(col_buckets.get("quantity", [])))
                price = _parse_num_str(_join_cell_words(col_buckets.get("unit_price", [])))
                note = _join_cell_words(col_buckets.get("note", []))

                if sub_val is None and qty is not None and price is not None:
                    calc = qty * price
                    sub_val = int(calc) if float(calc).is_integer() else calc
                if sub_val is None:
                    continue
                if qty is None:
                    qty = 1
                if price is None and qty:
                    calc_p = sub_val / qty
                    price = int(calc_p) if float(calc_p).is_integer() else round(calc_p, 2)

                extracted_items.append(
                    ParsedInvoiceItem(
                        name=name,
                        spec=spec,
                        unit=unit,
                        quantity=qty,
                        unit_price=price,
                        subtotal=sub_val,
                        note=note,
                    )
                )

    if not extracted_items:
        return None

    base_inv.items = extracted_items
    if base_inv.total_amount is None:
        tot = sum(i.subtotal for i in extracted_items)
        base_inv.total_amount = int(tot) if float(tot).is_integer() else tot
    return base_inv


def _scan_pdf(
    pdf_path: Path,
    *,
    gemini_api_key: str,
    pdf_password: str = "",
    threshold: int,
) -> ScanResult:
    """Extract text directly from digital PDF e-invoices, or render scanned PDFs for OCR."""
    import tempfile
    import pymupdf

    doc = pymupdf.open(str(pdf_path))
    try:
        if doc.needs_pass:
            unlocked = bool(pdf_password and doc.authenticate(pdf_password)) or bool(doc.authenticate(""))
            if not unlocked:
                raise ValueError("此 PDF 電子發票已加密保護，請填入「PDF 密碼」（常見為統一編號或身分證字號）後再掃描。")

        combined_text = _extract_pdf_lines(doc)

        if len(combined_text) >= 20 and not gemini_api_key:
            table_inv = _extract_pdf_table_invoice(doc, combined_text)
            if table_inv is not None:
                # Reconcile table items against detected total with OR-Tools CP-SAT
                table_inv, verified = _reconcile_with_cpsat(combined_text, table_inv)
                return ScanResult(
                    engine="pdf-table+cpsat",
                    document_type=_detect_document_type(combined_text),
                    raw_text=combined_text,
                    invoice=table_inv,
                    recommended_sheet=_recommend_sheet(table_inv.total_amount, threshold=threshold),
                    cpsat_verified=verified,
                )

            inv = parse_invoice_text(combined_text)
            if inv.items or inv.invoice_no or inv.total_amount is not None:
                inv, verified = _reconcile_with_cpsat(combined_text, inv)
                return ScanResult(
                    engine="pdf-text+cpsat",
                    document_type=_detect_document_type(combined_text),
                    raw_text=combined_text,
                    invoice=inv,
                    recommended_sheet=_recommend_sheet(inv.total_amount, threshold=threshold),
                    cpsat_verified=verified,
                )

        # Fallback for scanned/image-only PDFs (or when Gemini Vision is requested)
        with tempfile.TemporaryDirectory() as tmp:
            img_path = Path(tmp) / "page_1.png"
            pix = doc[0].get_pixmap(dpi=150)
            pix.save(str(img_path))
            return scan_document(img_path, gemini_api_key=gemini_api_key, threshold=threshold)
    finally:
        doc.close()


def scan_document(
    image_path: Path,
    *,
    gemini_api_key: str = "",
    pdf_password: str = "",
    threshold: int = 3000,
) -> ScanResult:
    """Scan an invoice/quotation (image or PDF) using Gemini Vision, PDF text, or RapidOCR."""
    import gc
    import tempfile

    api_key = gemini_api_key.strip() or os.environ.get("GEMINI_API_KEY", "").strip()
    if image_path.suffix.lower() == ".pdf":
        return _scan_pdf(
            image_path,
            gemini_api_key=api_key,
            pdf_password=pdf_password.strip(),
            threshold=threshold,
        )
    if api_key:
        with tempfile.TemporaryDirectory() as tmp:
            prepped = _prepare_ocr_image(image_path, Path(tmp), max_dim=1600)
            doc_type, raw_text, inv = _scan_with_gemini(prepped, api_key)
        return ScanResult(
            engine="gemini",
            document_type=doc_type,
            raw_text=raw_text,
            invoice=inv,
            recommended_sheet=_recommend_sheet(inv.total_amount, threshold=threshold),
        )

    ocr = _get_rapid_ocr()
    with tempfile.TemporaryDirectory() as tmp:
        prepped = _prepare_ocr_image(image_path, Path(tmp), max_dim=1600)
        result, _elapse = ocr(str(prepped))
    gc.collect()
    raw_text = _group_ocr_boxes_into_lines(result or [])
    inv = parse_invoice_text(raw_text)
    inv, verified = _reconcile_with_cpsat(raw_text, inv)
    doc_type = _detect_document_type(raw_text)
    return ScanResult(
        engine="rapidocr+cpsat",
        document_type=doc_type,
        raw_text=raw_text,
        invoice=inv,
        recommended_sheet=_recommend_sheet(inv.total_amount, threshold=threshold),
        cpsat_verified=verified,
    )


_THSR_STATIONS = ("南港", "台北", "臺北", "板橋", "桃園", "新竹", "苗栗", "台中", "臺中", "彰化", "雲林", "嘉義", "台南", "臺南", "左營")


def parse_travel_receipt(
    file_path: Path,
    *,
    gemini_api_key: str = "",
    pdf_password: str = "",
) -> dict[str, Any]:
    """Extract travel receipt fields (HSR/TRA/Lodging/Bus) from a PDF or image."""
    import re
    import tempfile

    raw_text = ""
    engine = "rapidocr"
    inv: ParsedInvoice | None = None

    if file_path.suffix.lower() == ".pdf":
        import pymupdf

        doc = pymupdf.open(str(file_path))
        try:
            if doc.needs_pass:
                unlocked = bool(pdf_password and doc.authenticate(pdf_password)) or bool(doc.authenticate(""))
                if not unlocked:
                    raise ValueError("此 PDF 已加密保護，請填入密碼後再掃描。")
            raw_text = _extract_pdf_lines(doc)
            if len(raw_text) >= 15 and not gemini_api_key.strip():
                engine = "pdf-text"
                inv = parse_invoice_text(raw_text)
            else:
                with tempfile.TemporaryDirectory() as tmp:
                    img_p = Path(tmp) / "p1.png"
                    doc[0].get_pixmap(dpi=150).save(str(img_p))
                    return parse_travel_receipt(img_p, gemini_api_key=gemini_api_key)
        finally:
            doc.close()
    else:
        res = scan_document(file_path, gemini_api_key=gemini_api_key)
        engine = res.engine
        raw_text = res.raw_text
        inv = res.invoice

    compact = raw_text.replace(" ", "").replace("\u3000", "")
    amount = inv.total_amount if (inv and inv.total_amount) else None
    if amount is None:
        m_amt = re.search(r"(?:票價|總計|合計|實付|金額|TWD|NT\$|\$)\s*[:：]?\s*(\d+(?:,\d{3})*)", compact, re.IGNORECASE)
        if m_amt:
            amount = int(m_amt.group(1).replace(",", ""))

    # Detect date (YYYY/MM/DD or YYYY-MM-DD or YYY/MM/DD)
    mm_dd = ""
    m_dt = re.search(r"(20\d{2}|1[01]\d)[/\-.](\d{1,2})[/\-.](\d{1,2})", raw_text)
    if m_dt:
        mm_dd = f"{int(m_dt.group(2)):02d}/{int(m_dt.group(3)):02d}"

    # Detect stations in the order they appear in the ticket text
    found_stations = sorted(
        (st for st in _THSR_STATIONS if st in compact),
        key=lambda st: compact.find(st),
    )
    route_text = "→".join(found_stations[:2]) if len(found_stations) >= 2 else ""

    if any(k in compact.upper() for k in ("高鐵", "高速鐵路", "THSR", "T-EX")):
        mode = "flight_hsr"
        summary = "搭乘高鐵"
    elif any(k in compact.upper() for k in ("臺鐵", "台鐵", "臺灣鐵路", "台灣鐵路", "自強", "莒光", "區間")):
        mode = "train"
        summary = "搭乘台鐵"
    elif any(k in compact.upper() for k in ("飯店", "旅店", "商旅", "酒店", "民宿", "會館", "HOTEL", "INN", "住宿")):
        mode = "lodging"
        summary = f"住宿（{inv.seller_name}）" if (inv and inv.seller_name) else "出差住宿"
    else:
        mode = "bus_mrt"
        summary = inv.seller_name if (inv and inv.seller_name) else "交通費單據"

    return {
        "engine": engine,
        "receipt_type": mode,
        "date": mm_dd,
        "route_text": route_text,
        "fare": 0 if mode == "lodging" else (amount or 0),
        "lodging_actual": (amount or 0) if mode == "lodging" else 0,
        "summary": summary,
        "raw_text": raw_text,
    }

