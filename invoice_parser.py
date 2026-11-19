"""Parser for Taiwan electronic invoices (e.g. Showba / 小北百貨 receipt text & MOF CSV)."""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ParsedInvoiceItem:
    """Single line item extracted from an electronic invoice."""

    name: str
    spec: str = ""
    unit: str = ""
    quantity: int | float = 1
    unit_price: int | float | None = None
    subtotal: int | float = 0
    note: str = ""


@dataclass
class ParsedInvoice:
    """Structured data extracted from a Taiwan electronic invoice."""

    invoice_no: str = ""
    roc_date: str = ""
    seller_name: str = ""
    total_amount: int | float | None = None
    items: list[ParsedInvoiceItem] = field(default_factory=list)


_INVOICE_NO_RE = re.compile(r"\b([A-Z]{2})[-\s]?(\d{8})\b")
_DATE_RE = re.compile(
    r"(?:(?P<y>\d{3,4})[-/](?P<m>\d{1,2})[-/](?P<d>\d{1,2}))|(?:(?P<y2>\d{3,4})\s*年\s*(?P<m2>\d{1,2})\s*月\s*(?P<d2>\d{1,2})\s*日)"
)
# Matches: "49 x 2  98TX" or "49 * 2 = 98" or "$49 x 2 $98TX"
_QTY_PRICE_LINE_RE = re.compile(
    r"^\s*\$?(?P<price>\d+(?:,\d{3})*(?:\.\d+)?)\s*[xX×*]\s*(?P<qty>\d+(?:\.\d+)?)\s*(?:=\s*)?\$?(?P<sub>-?\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:TX|T)?\s*$"
)
# Matches single-line: "3M菜瓜布  2  49  98TX" or "3M菜瓜布  $49 x 2  $98TX"
_SINGLE_LINE_MUL_RE = re.compile(
    r"^(?P<name>.+?)\s+\$?(?P<price>\d+(?:,\d{3})*(?:\.\d+)?)\s*[xX×*]\s*(?P<qty>\d+(?:\.\d+)?)\s+\$?(?P<sub>\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:TX|T)?\s*$"
)
_SINGLE_LINE_COLS_RE = re.compile(
    r"^(?P<name>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+\$?(?P<price>\d+(?:,\d{3})*(?:\.\d+)?)\s+\$?(?P<sub>\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:TX|T)?\s*$"
)
_SINGLE_LINE_SIMPLE_RE = re.compile(
    r"^(?P<name>.+?)\s+(?:NT\$|\$)?(?P<sub>\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:TX|T|元)\s*$"
)
_SINGLE_LINE_UNIT_COLS_RE = re.compile(
    r"^(?P<name>.+?)\s+(?P<unit>[個包支箱式批台組張瓶盒捲條雙份本顆粒罐袋])\s+(?P<qty>\d+(?:\.\d+)?)\s+(?:NT\$|\$)?(?P<price>\d+(?:,\d{3})*(?:\.\d+)?)\s+(?:NT\$|\$)?(?P<sub>\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:TX|T|元)?\s*$"
)
_SINGLE_LINE_QTY_SUB_RE = re.compile(
    r"^(?P<name>.+?)\s+(?P<qty>\d{1,3})\s+(?:NT\$|\$)?(?P<sub>\d+(?:,\d{3})*(?:\.\d+)?)\s*(?:TX|T|元)?\s*$"
)
_TOTAL_RE = re.compile(
    r"(?:總\s*計\s*金\s*額|合\s*計\s*金\s*額|總\s*金\s*額|發\s*票\s*金\s*額|實\s*付\s*金\s*額|應\s*付\s*金\s*額|總\s*計|合\s*計|應稅銷售額)\s*[:：]?\s*(?:NT\$|\$)?\s*(?P<total>\d+(?:,\d{3})*(?:\.\d+)?)\s*元?"
)
_TOTAL_LABEL_ONLY_RE = re.compile(
    r"^(?:總\s*計\s*金\s*額|合\s*計\s*金\s*額|總\s*金\s*額|發\s*票\s*金\s*額|總\s*計|合\s*計)\s*[:：]?\s*$"
)
_MONEY_ONLY_RE = re.compile(r"^(?:NT\$|\$)?\s*(?P<amt>\d+(?:,\d{3})*(?:\.\d+)?)\s*元?\s*$")

_SKIP_PREFIXES = (
    "電子發票",
    "證明聯",
    "交易明細",
    "賣方",
    "買方",
    "統一編號",
    "隨機碼",
    "機號",
    "序號",
    "店別",
    "門市",
    "電話",
    "地址",
    "退貨",
    "信用卡",
    "現金",
    "找零",
    "載具",
    "應稅",
    "免稅",
    "稅額",
    "總計",
    "合計",
    "小計",
    "=",
    "-",
    "*",
)


_PAREN_SPEC_RE = re.compile(r"^(?P<name>.+?)\s*[（(](?P<spec>[^）)]+)[）)]?\s*$")
_TRAILING_SPEC_RE = re.compile(
    r"^(?P<name>.+?)\s*(?P<spec>(?:\d+(?:\.\d+)?\s*(?:片裝|入|張|孔\d*座|座|埠|磅|ml|mL|ML|L|cm|mm|m|M|g|kg|W|V|吋|寸|公升|毫升|公分|公尺))+|(?:A3|A4|A5|B4|B5)(?:\s*\d+磅)?|(?:黑色|藍色|紅色|綠色|白色|黃色|灰色|透明))$",
    re.IGNORECASE,
)


def split_name_and_spec(raw_name: str) -> tuple[str, str]:
    """Separate a raw product description into (name, spec) cleanly."""
    text = raw_name.strip().lstrip("/").strip()
    if not text:
        return "", ""

    # 1. Parenthesized spec at end: e.g. "高保密碎紙機(大容量紙桶)" or "碎紙機(大容量紙桶"
    m_p = _PAREN_SPEC_RE.match(text)
    if m_p and len(m_p.group("name").strip()) >= 2:
        return m_p.group("name").strip(), m_p.group("spec").strip()

    # 2. Explicit separator like "品名 / 規格" or "品名 - 規格"
    for sep in (" / ", " - ", "｜", "|"):
        if sep in text:
            left, right = text.split(sep, 1)
            if len(left.strip()) >= 2 and right.strip():
                return left.strip(), right.strip()

    # 3. Trailing numeric/unit/color spec: e.g. "3M菜瓜布5片裝", "延長線3孔4座", "奇異筆黑色"
    m_t = _TRAILING_SPEC_RE.match(text)
    if m_t and len(m_t.group("name").strip()) >= 2:
        return m_t.group("name").strip(), m_t.group("spec").strip()

    return text, ""


def _make_item(
    raw_name: str,
    *,
    quantity: int | float = 1,
    unit_price: int | float | None = None,
    subtotal: int | float = 0,
    unit: str = "",
    spec: str = "",
) -> ParsedInvoiceItem:
    if spec:
        name = raw_name.strip().lstrip("/").strip()
    else:
        name, spec = split_name_and_spec(raw_name)
    return ParsedInvoiceItem(
        name=name,
        spec=spec,
        unit=unit,
        quantity=quantity,
        unit_price=unit_price,
        subtotal=subtotal,
    )


def _to_num(raw: str) -> int | float:
    val = float(raw.replace(",", "").strip())
    return int(val) if val.is_integer() else val


def format_roc_date(raw_date: str) -> str:
    """Convert Western or ROC date text into '中華民國 YYY 年 MM 月 DD 日'."""
    m = _DATE_RE.search(raw_date)
    if not m:
        return ""
    year = int(m.group("y") or m.group("y2"))
    if year > 1911:
        year -= 1911
    month = int(m.group("m") or m.group("m2"))
    day = int(m.group("d") or m.group("d2"))
    return f"中華民國 {year} 年 {month} 月 {day} 日"


def parse_invoice_text(text: str) -> ParsedInvoice:
    """Parse pasted or extracted Taiwan e-invoice receipt text (e.g. 小北百貨)."""
    inv = ParsedInvoice()
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    pending_name: str | None = None
    expect_total_next_line = False

    for line in lines:
        if expect_total_next_line:
            m_money = _MONEY_ONLY_RE.match(line)
            if m_money:
                inv.total_amount = _to_num(m_money.group("amt"))
                expect_total_next_line = False
                pending_name = None
                continue
            expect_total_next_line = False

        if not inv.seller_name and ("小北百貨" in line or "股份有限公司" in line or "百貨" in line):
            inv.seller_name = line

        if not inv.invoice_no:
            m_no = _INVOICE_NO_RE.search(line)
            if m_no:
                inv.invoice_no = f"{m_no.group(1)}-{m_no.group(2)}"

        if not inv.roc_date:
            roc = format_roc_date(line)
            if roc:
                inv.roc_date = roc

        m_total = _TOTAL_RE.search(line)
        if m_total:
            inv.total_amount = _to_num(m_total.group("total"))
            pending_name = None
            continue

        if _TOTAL_LABEL_ONLY_RE.match(line):
            expect_total_next_line = True
            pending_name = None
            continue

        # Two-line item format: previous line was item name, current line is "price x qty subTX"
        m_qp = _QTY_PRICE_LINE_RE.match(line)
        if m_qp and pending_name:
            inv.items.append(
                _make_item(
                    pending_name,
                    quantity=_to_num(m_qp.group("qty")),
                    unit_price=_to_num(m_qp.group("price")),
                    subtotal=_to_num(m_qp.group("sub")),
                )
            )
            pending_name = None
            continue

        # Single-line: "name price x qty subTX"
        m_sm = _SINGLE_LINE_MUL_RE.match(line)
        if m_sm:
            inv.items.append(
                _make_item(
                    m_sm.group("name"),
                    quantity=_to_num(m_sm.group("qty")),
                    unit_price=_to_num(m_sm.group("price")),
                    subtotal=_to_num(m_sm.group("sub")),
                )
            )
            pending_name = None
            continue

        # Single-line with unit: "name unit qty price sub"
        m_su = _SINGLE_LINE_UNIT_COLS_RE.match(line)
        if m_su and not any(line.startswith(p) for p in _SKIP_PREFIXES):
            inv.items.append(
                _make_item(
                    m_su.group("name"),
                    unit=m_su.group("unit").strip(),
                    quantity=_to_num(m_su.group("qty")),
                    unit_price=_to_num(m_su.group("price")),
                    subtotal=_to_num(m_su.group("sub")),
                )
            )
            pending_name = None
            continue

        # Single-line: "name qty price subTX"
        m_sc = _SINGLE_LINE_COLS_RE.match(line)
        if m_sc and not any(line.startswith(p) for p in _SKIP_PREFIXES):
            inv.items.append(
                _make_item(
                    m_sc.group("name"),
                    quantity=_to_num(m_sc.group("qty")),
                    unit_price=_to_num(m_sc.group("price")),
                    subtotal=_to_num(m_sc.group("sub")),
                )
            )
            pending_name = None
            continue

        # Single-line: "name qty sub" (e.g. "碎紙機 1 3,788")
        m_qs = _SINGLE_LINE_QTY_SUB_RE.match(line)
        if m_qs and not any(line.startswith(p) for p in _SKIP_PREFIXES):
            qty = _to_num(m_qs.group("qty"))
            sub = _to_num(m_qs.group("sub"))
            unit_p = sub / qty if qty else sub
            unit_p_clean = int(unit_p) if float(unit_p).is_integer() else round(unit_p, 2)
            inv.items.append(
                _make_item(
                    m_qs.group("name"),
                    quantity=qty,
                    unit_price=unit_p_clean,
                    subtotal=sub,
                )
            )
            pending_name = None
            continue

        # Single-line: "name 98TX"
        m_ss = _SINGLE_LINE_SIMPLE_RE.match(line)
        if m_ss and not any(line.startswith(p) for p in _SKIP_PREFIXES):
            sub = _to_num(m_ss.group("sub"))
            inv.items.append(
                _make_item(
                    m_ss.group("name"),
                    quantity=1,
                    unit_price=sub,
                    subtotal=sub,
                )
            )
            pending_name = None
            continue

        # Skip metadata / header lines before treating as a pending item name
        normalized = line.replace(" ", "")
        if (
            any(normalized.startswith(p) for p in _SKIP_PREFIXES)
            or _INVOICE_NO_RE.search(line)
            or _DATE_RE.search(line)
            or line == inv.seller_name
        ):
            pending_name = None
            continue

        pending_name = line

    if inv.total_amount is None and inv.items:
        total = sum(i.subtotal for i in inv.items)
        inv.total_amount = int(total) if float(total).is_integer() else total

    return inv


def parse_einvoice_csv(csv_path: Path) -> list[ParsedInvoice]:
    """Parse Taiwan MOF e-invoice exported CSV into a list of ParsedInvoice objects."""
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
        raise RuntimeError(f"無法讀取電子發票 CSV {csv_path}：{last_err}")

    grouped: dict[str, ParsedInvoice] = {}
    order: list[str] = []

    for row in rows:
        inv_no = row.get("發票號碼") or row.get("載具發票號碼") or ""
        if not inv_no:
            continue
        m_no = _INVOICE_NO_RE.search(inv_no)
        if m_no:
            inv_no = f"{m_no.group(1)}-{m_no.group(2)}"

        if inv_no not in grouped:
            raw_date = row.get("發票日期") or row.get("日期") or ""
            seller = row.get("賣方名稱") or row.get("商店名稱") or row.get("營業人名稱") or ""
            grouped[inv_no] = ParsedInvoice(
                invoice_no=inv_no,
                roc_date=format_roc_date(raw_date),
                seller_name=seller,
            )
            order.append(inv_no)

        name = row.get("品名") or row.get("商品名稱") or ""
        qty_raw = row.get("數量") or "1"
        price_raw = row.get("單價") or ""
        sub_raw = row.get("小計") or row.get("金額") or row.get("總金額") or ""

        if name:
            qty = _to_num(qty_raw)
            price = _to_num(price_raw) if price_raw else None
            if sub_raw:
                sub = _to_num(sub_raw)
            elif price is not None:
                sub = qty * price
                sub = int(sub) if float(sub).is_integer() else sub
            else:
                sub = 0

            grouped[inv_no].items.append(
                _make_item(
                    name,
                    spec=row.get("規格", ""),
                    unit=row.get("單位", ""),
                    quantity=qty,
                    unit_price=price,
                    subtotal=sub,
                )
            )

    for inv in grouped.values():
        if inv.items:
            total = sum(i.subtotal for i in inv.items)
            inv.total_amount = int(total) if float(total).is_integer() else total

    return [grouped[no] for no in order]
