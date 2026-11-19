"""OR-Tools CP-SAT solver for reconciling OCR/PDF invoice and quotation line items."""

from __future__ import annotations

import re
from dataclasses import dataclass

from ortools.sat.python import cp_model

from invoice_parser import ParsedInvoice, ParsedInvoiceItem, split_name_and_spec


@dataclass
class CandidateItem:
    """A candidate line item extracted from OCR or PDF text before CP-SAT filtering."""

    name: str
    spec: str = ""
    unit: str = ""
    quantity: int | float = 1
    unit_price: int | float | None = None
    subtotal: int | float = 0
    note: str = ""
    has_tax_marker: bool = False
    arithmetic_ok: bool = False
    is_noise_hint: bool = False


_NOISE_WORDS = ("現金", "找零", "信用卡", "悠遊卡", "LINEPAY", "刷卡", "實付", "應付", "稅額", "免稅", "應稅")


def extract_candidates_from_text(raw_text: str, parsed: ParsedInvoice) -> list[CandidateItem]:
    """Build CandidateItem list from parsed items and any extra numeric lines in raw_text."""
    candidates: list[CandidateItem] = []
    seen_names: set[str] = set()

    for it in parsed.items:
        arith_ok = (
            it.unit_price is not None
            and abs((it.quantity * it.unit_price) - it.subtotal) < 0.01
        )
        is_noise = any(nw in it.name.replace(" ", "").upper() for nw in _NOISE_WORDS)
        candidates.append(
            CandidateItem(
                name=it.name,
                spec=it.spec,
                unit=it.unit,
                quantity=it.quantity,
                unit_price=it.unit_price,
                subtotal=it.subtotal,
                note=it.note,
                has_tax_marker=f"{it.subtotal}TX" in raw_text.replace(" ", ""),
                arithmetic_ok=arith_ok,
                is_noise_hint=is_noise,
            )
        )
        seen_names.add(it.name)

    # Also scan raw lines for fallback "name  amount" patterns that might be missed or payment lines
    line_num_re = re.compile(r"^(?P<name>[^\d\s].*?)\s+\$?(?P<amt>\d+(?:,\d{3})*(?:\.\d+)?)\s*(?P<tx>TX|T)?\s*$")
    for line in raw_text.splitlines():
        s = line.strip()
        m = line_num_re.match(s)
        if not m:
            continue
        raw_name = m.group("name").strip()
        name, spec = split_name_and_spec(raw_name)
        if name in seen_names or any(k in raw_name for k in ("總計", "合計", "隨機碼", "統編", "統一編號", "日期")):
            continue
        amt = float(m.group("amt").replace(",", ""))
        amt_clean: int | float = int(amt) if amt.is_integer() else amt
        is_noise = any(nw in name.replace(" ", "").upper() for nw in _NOISE_WORDS)
        candidates.append(
            CandidateItem(
                name=name,
                spec=spec,
                quantity=1,
                unit_price=amt_clean,
                subtotal=amt_clean,
                has_tax_marker=bool(m.group("tx")),
                arithmetic_ok=True,
                is_noise_hint=is_noise,
            )
        )
        seen_names.add(name)

    return candidates


def reconcile_items_with_cpsat(
    candidates: list[CandidateItem],
    detected_total: int | float | None,
    *,
    time_limit_sec: float = 2.0,
) -> tuple[list[ParsedInvoiceItem], bool]:
    """Use OR-Tools CP-SAT to select the optimal subset of items matching `detected_total`.

    Returns (selected_items, exact_total_match).
    """
    if not candidates:
        return [], False

    # Filter out negative/zero or huge non-item numbers if total is known
    has_float = any(isinstance(c.subtotal, float) and not c.subtotal.is_integer() for c in candidates) or (
        detected_total is not None and isinstance(detected_total, float) and not float(detected_total).is_integer()
    )
    scale = 100 if has_float else 1
    scaled_subs = [max(0, int(round(c.subtotal * scale))) for c in candidates]
    n = len(candidates)

    model = cp_model.CpModel()
    keep = [model.new_bool_var(f"keep_{i}") for i in range(n)]

    # Scoring: reward verified arithmetic & tax markers, penalize noise words (e.g. 現金/找零)
    scores: list[int] = []
    for c in candidates:
        sc = 10
        if c.arithmetic_ok:
            sc += 20
        if c.has_tax_marker:
            sc += 15
        if c.is_noise_hint:
            sc -= 50
        scores.append(sc)

    if detected_total is not None and detected_total > 0:
        scaled_total = int(round(detected_total * scale))
        max_sum = max(sum(scaled_subs), scaled_total)
        selected_sum = sum(scaled_subs[i] * keep[i] for i in range(n))
        diff = model.new_int_var(0, max_sum, "total_diff")
        model.add(selected_sum - scaled_total <= diff)
        model.add(scaled_total - selected_sum <= diff)

        # Primary objective: match detected_total (minimize diff); secondary: maximize item quality score
        model.minimize(diff * 1000 - sum(scores[i] * keep[i] for i in range(n)))
    else:
        # No total detected: keep all positive-score items and drop noise hints
        model.maximize(sum(scores[i] * keep[i] for i in range(n)))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_sec
    status = solver.solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return [
            ParsedInvoiceItem(
                name=c.name, spec=c.spec, unit=c.unit, quantity=c.quantity, unit_price=c.unit_price, subtotal=c.subtotal, note=c.note
            )
            for c in candidates
            if not c.is_noise_hint
        ], False

    selected = [
        ParsedInvoiceItem(
            name=candidates[i].name,
            spec=candidates[i].spec,
            unit=candidates[i].unit,
            quantity=candidates[i].quantity,
            unit_price=candidates[i].unit_price,
            subtotal=candidates[i].subtotal,
            note=candidates[i].note,
        )
        for i in range(n)
        if solver.value(keep[i]) == 1
    ]

    exact_match = False
    if detected_total is not None and detected_total > 0:
        sel_sum = sum(it.subtotal for it in selected)
        exact_match = abs(sel_sum - detected_total) < 0.01

    return selected, exact_match

