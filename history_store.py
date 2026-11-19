"""Persistent cumulative history store for generated purchase requisitions."""

from __future__ import annotations

import csv
import io
import json
import threading
import uuid
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from fill_purchase_forms import PurchaseRequest

_LOCK = threading.Lock()


def load_history(history_path: Path) -> list[dict[str, Any]]:
    """Load all saved purchase records (newest first)."""
    if not history_path.exists():
        return []
    with _LOCK:
        try:
            data = json.loads(history_path.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else []
        except Exception:
            return []


def _save_history(history_path: Path, records: list[dict[str, Any]]) -> None:
    history_path.parent.mkdir(parents=True, exist_ok=True)
    history_path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def append_history(
    history_path: Path,
    requests: list[PurchaseRequest],
    files: list[Path],
) -> list[dict[str, Any]]:
    """Append newly generated PurchaseRequest records to the cumulative history file."""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_entries: list[dict[str, Any]] = []

    for req, out_file in zip(requests, files):
        item_names = "、".join(it.name for it in req.items[:3])
        if len(req.items) > 3:
            item_names += f" 等共{len(req.items)}項"
        new_entries.append({
            "id": uuid.uuid4().hex[:10],
            "created_at": now_str,
            "request_id": req.request_id,
            "source_type": req.source_type,
            "budget_code": req.budget_code,
            "plan_name": req.plan_name,
            "department": req.department,
            "payer_name": req.payer_name,
            "purpose": req.purpose,
            "total_amount": req.total_amount,
            "item_count": len(req.items),
            "items_summary": item_names,
            "output_file": out_file.name,
        })

    with _LOCK:
        existing = []
        if history_path.exists():
            try:
                existing = json.loads(history_path.read_text(encoding="utf-8"))
                if not isinstance(existing, list):
                    existing = []
            except Exception:
                existing = []
        # Keep newest first
        updated = list(reversed(new_entries)) + existing
        _save_history(history_path, updated)
    return new_entries


def delete_history_record(history_path: Path, record_id: str) -> bool:
    """Remove a single record by ID."""
    with _LOCK:
        records = load_history_unlocked(history_path)
        filtered = [r for r in records if r.get("id") != record_id]
        if len(filtered) == len(records):
            return False
        _save_history(history_path, filtered)
        return True


def load_history_unlocked(history_path: Path) -> list[dict[str, Any]]:
    if not history_path.exists():
        return []
    try:
        data = json.loads(history_path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def clear_history(history_path: Path) -> None:
    """Clear all cumulative history records."""
    with _LOCK:
        _save_history(history_path, [])


def summarize_history(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Compute total count, total amount, and breakdown by budget_code."""
    total_amount = 0.0
    by_code: dict[str, float] = defaultdict(float)

    for r in records:
        amt = float(r.get("total_amount") or 0)
        total_amount += amt
        code = r.get("budget_code") or "未指定"
        by_code[code] += amt

    def _clean(v: float) -> int | float:
        return int(v) if v.is_integer() else round(v, 2)

    return {
        "total_count": len(records),
        "total_amount": _clean(total_amount),
        "by_budget_code": {k: _clean(v) for k, v in by_code.items()},
    }


def export_history_csv_bytes(records: list[dict[str, Any]]) -> bytes:
    """Export cumulative history records as UTF-8 BOM CSV bytes for Excel."""
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow([
        "產出時間", "單號", "資料來源", "預算科目代碼", "計畫名稱",
        "請購單位", "代墊人", "用途說明", "品項摘要", "品項數", "總金額", "檔案名稱",
    ])
    for r in records:
        writer.writerow([
            r.get("created_at", ""),
            r.get("request_id", ""),
            r.get("source_type", ""),
            r.get("budget_code", ""),
            r.get("plan_name", ""),
            r.get("department", ""),
            r.get("payer_name", ""),
            r.get("purpose", ""),
            r.get("items_summary", ""),
            r.get("item_count", 0),
            r.get("total_amount", 0),
            r.get("output_file", ""),
        ])
    return buf.getvalue().encode("utf-8-sig")

