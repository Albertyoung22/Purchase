"""OR-Tools CP-SAT solver for optimal splitting of purchase items into requisition forms."""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from ortools.sat.python import cp_model


@dataclass
class SplitItem:
    """An item or invoice entry to be assigned to a purchase requisition form."""

    item_id: int
    name: str
    amount: int | float
    budget_code: str = ""
    plan_name: str = ""
    level1: str = ""
    level2: str = ""
    level3: str = ""
    department: str = ""
    request_date: str = ""
    payer_name: str = ""
    purpose: str = ""
    spec: str = ""
    unit: str = ""
    quantity: int | float | None = None
    unit_price: int | float | None = None
    manual_subtotal: int | float | None = None
    invoice_no: str = ""
    note: str = ""


@dataclass
class FormPlan:
    """A single grouped purchase form produced by the CP-SAT solver."""

    form_index: int
    budget_code: str
    department: str
    total_amount: int | float
    items: list[SplitItem] = field(default_factory=list)


def _solve_partition(
    items: list[SplitItem],
    *,
    max_amount: int | float,
    max_rows: int,
    time_limit_sec: float,
    allow_oversized: bool = True,
) -> list[list[SplitItem]]:
    """Solve bin-packing with cardinality & capacity limits using OR-Tools CP-SAT."""
    if not items:
        return []

    oversized_bins: list[list[SplitItem]] = []
    normal_items: list[SplitItem] = []

    for it in items:
        if it.amount < 0:
            raise ValueError(f"品項「{it.name}」金額不可為負數：{it.amount}")
        if it.amount > max_amount:
            if not allow_oversized:
                raise ValueError(
                    f"品項「{it.name}」金額 ({it.amount}) 已超過單張上限 ({max_amount})，無法分單"
                )
            oversized_bins.append([it])
        else:
            normal_items.append(it)

    if not normal_items:
        return oversized_bins

    n = len(normal_items)
    # Scale amounts to integers (x100) only if decimals exist, so CP-SAT stays exact.
    has_float = any(isinstance(it.amount, float) and not it.amount.is_integer() for it in normal_items) or (
        isinstance(max_amount, float) and not max_amount.is_integer()
    )
    scale = 100 if has_float else 1
    scaled_amounts = [int(round(it.amount * scale)) for it in normal_items]
    scaled_cap = int(round(max_amount * scale))

    model = cp_model.CpModel()
    max_bins = n

    # x[i][b] == 1 iff item i is placed in form b
    x = [[model.new_bool_var(f"x_{i}_{b}") for b in range(max_bins)] for i in range(n)]
    # y[b] == 1 iff form b is used
    y = [model.new_bool_var(f"y_{b}") for b in range(max_bins)]

    # Each item must be assigned to exactly one form
    for i in range(n):
        model.add_exactly_one(x[i][b] for b in range(max_bins))

    # Form capacity (amount) and cardinality (max_rows) constraints
    for b in range(max_bins):
        model.add(sum(scaled_amounts[i] * x[i][b] for i in range(n)) <= scaled_cap * y[b])
        model.add(sum(x[i][b] for i in range(n)) <= max_rows * y[b])

    # Symmetry breaking: fill lower-indexed forms first
    for b in range(max_bins - 1):
        model.add(y[b] >= y[b + 1])

    model.minimize(sum(y))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_sec
    status = solver.solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(f"CP-SAT 無法找到可行分單解 (status={solver.status_name(status)})")

    bins: list[list[SplitItem]] = []
    for b in range(max_bins):
        if solver.value(y[b]) == 1:
            bin_items = [normal_items[i] for i in range(n) if solver.value(x[i][b]) == 1]
            bins.append(bin_items)
    return bins + oversized_bins


def solve_form_splitting(
    items: Iterable[SplitItem],
    *,
    max_amount: int | float = 3000,
    max_rows: int = 7,
    group_by_budget_code: bool = True,
    time_limit_sec: float = 10.0,
    allow_oversized: bool = True,
) -> list[FormPlan]:
    """Group items into the minimum number of forms respecting amount and row limits."""
    if max_rows <= 0:
        raise ValueError("max_rows 必須大於 0")
    if max_amount <= 0:
        raise ValueError("max_amount 必須大於 0")

    partitions: dict[tuple[str, str], list[SplitItem]] = defaultdict(list)
    for it in items:
        key = (it.budget_code, it.department) if group_by_budget_code else ("", "")
        partitions[key].append(it)

    plans: list[FormPlan] = []
    form_idx = 1
    for (budget_code, department), part_items in partitions.items():
        bins = _solve_partition(
            part_items,
            max_amount=max_amount,
            max_rows=max_rows,
            time_limit_sec=time_limit_sec,
            allow_oversized=allow_oversized,
        )
        for bin_items in bins:
            total = sum(i.amount for i in bin_items)
            total_clean = int(total) if float(total).is_integer() else round(total, 2)
            plans.append(
                FormPlan(
                    form_index=form_idx,
                    budget_code=budget_code,
                    department=department,
                    total_amount=total_clean,
                    items=bin_items,
                )
            )
            form_idx += 1

    return plans

