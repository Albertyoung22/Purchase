"""Real Excel print preview generator using Microsoft Excel COM -> PDF -> PNG."""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

import openpyxl


def export_xlsx_to_pdf_with_libreoffice(xlsx_path: Path, pdf_path: Path, timeout_sec: int = 60) -> bool:
    """Use LibreOffice Calc headless mode (Linux/Render) to export .xlsx to PDF."""
    import tempfile

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # LibreOffice does not auto-rotate ASCII '(' ')' in vertical CJK cells (e.g. J8 '(簽章)').
    # Convert to vertical presentation forms '︵' '︶' on a temporary copy before PDF export.
    with tempfile.TemporaryDirectory() as tmp:
        tmp_xlsx = Path(tmp) / xlsx_path.name
        wb = openpyxl.load_workbook(xlsx_path)
        for ws in wb.worksheets:
            for coord in ("J8", "J7", "J9"):
                val = ws[coord].value
                if isinstance(val, str) and "簽章" in val:
                    ws[coord].value = val.replace("(", "︵").replace(")", "︶").replace("（", "︵").replace("）", "︶")
        wb.save(tmp_xlsx)

        proc = subprocess.run(
            [
                soffice,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_path.parent.resolve()),
                str(tmp_xlsx.resolve()),
            ],
            capture_output=True,
            text=True,
            timeout=timeout_sec,
        )
    expected = pdf_path.parent / f"{xlsx_path.stem}.pdf"
    if proc.returncode == 0 and expected.exists():
        if expected.resolve() != pdf_path.resolve():
            expected.replace(pdf_path)
        return True
    return False


def export_xlsx_to_pdf_with_excel(xlsx_path: Path, pdf_path: Path, timeout_sec: int = 30) -> bool:
    """Use installed Microsoft Excel COM to export an .xlsx file to an exact print-layout PDF."""
    if os.name != "nt":
        return False
    xlsx_abs = str(xlsx_path.resolve()).replace("'", "''")
    pdf_abs = str(pdf_path.resolve()).replace("'", "''")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    ps_script = (
        "$excel = $null; $wb = $null; "
        "try { "
        "$excel = New-Object -ComObject Excel.Application; "
        "$excel.Visible = $false; "
        "$excel.DisplayAlerts = $false; "
        f"$wb = $excel.Workbooks.Open('{xlsx_abs}'); "
        "$wb.RefreshAll(); "
        f"$wb.ExportAsFixedFormat(0, '{pdf_abs}'); "
        "} finally { "
        "if ($wb -ne $null) { $wb.Close($false) }; "
        "if ($excel -ne $null) { $excel.Quit() }; "
        "}"
    )

    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        capture_output=True,
        text=True,
        timeout=timeout_sec,
    )
    return proc.returncode == 0 and pdf_path.exists()


def export_xlsx_to_pdf(xlsx_path: Path, pdf_path: Path) -> bool:
    """Export .xlsx to PDF using Microsoft Excel on Windows, or LibreOffice on Linux/Render."""
    if os.name == "nt" and export_xlsx_to_pdf_with_excel(xlsx_path, pdf_path):
        return True
    return export_xlsx_to_pdf_with_libreoffice(xlsx_path, pdf_path)


def export_multiple_xlsx_to_merged_pdf(xlsx_paths: list[Path], merged_pdf_path: Path) -> bool:
    """Export all generated .xlsx forms to PDF and merge them into a single multi-page PDF."""
    if not xlsx_paths:
        return False
    if len(xlsx_paths) == 1:
        return export_xlsx_to_pdf(xlsx_paths[0], merged_pdf_path)

    import pymupdf

    part_pdfs: list[Path] = []
    for idx, x_path in enumerate(xlsx_paths, start=1):
        p_pdf = merged_pdf_path.parent / f"_part_{idx}_{x_path.stem}.pdf"
        if not export_xlsx_to_pdf(x_path, p_pdf):
            return False
        part_pdfs.append(p_pdf)

    merged = pymupdf.open()
    try:
        for p_pdf in part_pdfs:
            with pymupdf.open(str(p_pdf)) as src:
                merged.insert_pdf(src)
        merged_pdf_path.parent.mkdir(parents=True, exist_ok=True)
        merged.save(str(merged_pdf_path))
    finally:
        merged.close()
        for p_pdf in part_pdfs:
            p_pdf.unlink(missing_ok=True)

    return merged_pdf_path.exists()


def render_pdf_to_pngs(pdf_path: Path, out_dir: Path, prefix: str = "preview") -> list[Path]:
    """Render each page of a PDF to PNG using PyMuPDF for crisp inline preview."""
    import pymupdf

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(str(pdf_path))
    images: list[Path] = []
    try:
        for idx, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=150)
            img_path = out_dir / f"{prefix}_p{idx}.png"
            pix.save(str(img_path))
            images.append(img_path)
    finally:
        doc.close()
    return images


def _eval_simple_formula(ws: openpyxl.worksheet.worksheet.Worksheet, val: object) -> str:
    """Evaluate basic =F*G, =ROUND(F*G,0), =SUM(H1:H2), and =H25 formulas for HTML preview."""
    import re

    if val is None:
        return ""
    if not (isinstance(val, str) and val.startswith("=")):
        return str(val)

    expr = val[1:].strip().upper()
    # =H28 or =H25
    m_ref = re.fullmatch(r"([A-Z]+)(\d+)", expr)
    if m_ref:
        return _eval_simple_formula(ws, ws[expr].value)

    # =F21*G21 or =ROUND(F21*G21,0)
    m_mul = re.fullmatch(r"(?:ROUND\()?([A-Z]+\d+)\s*\*\s*([A-Z]+\d+)(?:,\s*0\))?", expr)
    if m_mul:
        v1 = ws[m_mul.group(1)].value
        v2 = ws[m_mul.group(2)].value
        if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
            prod = v1 * v2
            return f"{int(round(prod)):,}"
        return ""

    # =SUM(H18:H24)
    m_sum = re.fullmatch(r"SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", expr)
    if m_sum:
        col = m_sum.group(1)
        r1, r2 = int(m_sum.group(2)), int(m_sum.group(4))
        total = 0.0
        for r in range(r1, r2 + 1):
            raw_txt = _eval_simple_formula(ws, ws[f"{col}{r}"].value).replace(",", "")
            try:
                if raw_txt:
                    total += float(raw_txt)
            except ValueError:
                pass
        return f"{int(round(total)):,}"

    return val


def render_xlsx_fallback_html(xlsx_path: Path) -> str:
    """Fallback HTML table preview if Microsoft Excel COM is not available (e.g. Render Linux)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws = wb.active

    merged_skip: set[tuple[int, int]] = set()
    merged_span: dict[tuple[int, int], tuple[int, int]] = {}
    for rng in ws.merged_cells.ranges:
        merged_span[(rng.min_row, rng.min_col)] = (
            rng.max_row - rng.min_row + 1,
            rng.max_col - rng.min_col + 1,
        )
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    merged_skip.add((r, c))

    rows_html: list[str] = []
    for r in range(1, min(ws.max_row, 40) + 1):
        if ws.row_dimensions[r].hidden:
            continue
        cells_html: list[str] = []
        for c in range(1, min(ws.max_column, 10) + 1):
            if (r, c) in merged_skip:
                continue
            rs, cs = merged_span.get((r, c), (1, 1))
            text = _eval_simple_formula(ws, ws.cell(r, c).value)
            attrs = ""
            if rs > 1:
                attrs += f' rowspan="{rs}"'
            if cs > 1:
                attrs += f' colspan="{cs}"'
            cells_html.append(f"<td{attrs}>{text}</td>")
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")

    return f'<table class="excel-fallback">{"".join(rows_html)}</table>'

