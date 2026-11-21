"""
דוח "פירוט הזמנות ששודרו לפי נהג" מפריוריטי (PDF או אקסל) - מקור לעמודת "מספר אריזות".

הדוח לא חובה: כשהמשתמש מוסיף אותו בחלון העריכה, מספר האריזות של כל שורת הזמנה
נלקח ממנו (התאמה לפי מספר הזמנה + מק"ט, ואם המק"ט מופיע כמה פעמים באותה הזמנה -
גם לפי הכמות). שורה שלא נמצאה בדוח מחושבת כרגיל לפי package_sizes.xlsx.
"""

from __future__ import annotations

import pathlib
import re
from typing import Any

import openpyxl
from pypdf import PdfReader

from pdf_orders import Rect, _cell_lines, _cell_text, _page_layout

# כותרות הדוח (אחרי איחוד השורות של כל כותרת) -> שם פנימי
COLUMNS = {
    "הזמנה": "order",
    'מק"ט': "sku",
    "מק'ט": "sku",
    "כמות למשלוח": "qty",
    "כמות": "qty",
    "מספר אריזות": "packages",
}


def _num(val: Any) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    t = str(val).replace(",", "").strip()
    try:
        return float(t) if t else None
    except ValueError:
        return None


def _clean_sku(val: Any) -> str:
    return re.sub(r"\s+", "", str(val or "")).strip().upper()


def _clean_order(val: Any) -> str:
    return re.sub(r"\s+", "", str(val or "")).strip().upper()


def _read_pdf(path: pathlib.Path) -> list[dict[str, Any]]:
    """שורות הטבלה מה-PDF: כל מילה משויכת לתא (מלבן) שלה, התאים לעמודות לפי הכותרת."""
    reader = PdfReader(str(path))
    rows: list[dict[str, Any]] = []
    columns: list[tuple[float, float, str]] | None = None
    for page in reader.pages:
        rects, tokens = _page_layout(page, reader)
        cells = [r for r in rects if not ((r.x1 - r.x0) > 1000 and (r.y1 - r.y0) > 100)]
        for tok in tokens:
            hits = [r for r in cells if r.contains(tok[0] + 1, tok[1] + 2)]
            if hits:
                min(hits, key=lambda r: r.area).tokens.append(tok)
        by_row: dict[int, list[tuple[Rect, str]]] = {}
        for r in cells:
            name = " ".join(_cell_lines(r))
            if name in COLUMNS:
                by_row.setdefault(round(r.y1), []).append((r, COLUMNS[name]))
        header = max(by_row.values(), key=len) if by_row else []
        header_y = None
        if len({c for _r, c in header}) >= 3:
            columns = [(r.x0, r.x1, c) for r, c in header]
            header_y = min(r.y0 for r, _c in header)
        if not columns:
            continue
        bands: dict[tuple[int, int], list[Rect]] = {}
        for r in cells:
            if r.tokens and (header_y is None or r.y1 <= header_y + 1):
                bands.setdefault((round(r.y1), round(r.y0)), []).append(r)
        for _key, band in sorted(bands.items(), key=lambda kv: -kv[0][0]):
            values: dict[str, Any] = {}
            for r in band:
                mid = (r.x0 + r.x1) / 2
                col = next((c for x0, x1, c in columns if x0 - 1 <= mid <= x1 + 1), None)
                if col:
                    values[col] = _cell_text(r, col).replace(" ", "") if col != "qty" else _cell_text(r, col)
            if values.get("order") or values.get("sku"):
                rows.append(values)
    return rows


def _read_excel(path: pathlib.Path) -> list[dict[str, Any]]:
    try:
        ws = openpyxl.load_workbook(path, data_only=True).active
    except TypeError:
        # קובץ אקסל מפריוריטי עם עיצוב שמשבש את openpyxl - אותו תיקון כמו בקובץ ההזמנות
        from sort_orders_by_region import load_workbook_safe

        ws = load_workbook_safe(path).active
    header_row, idx = None, {}
    for r in range(1, min(ws.max_row, 15) + 1):
        names = [str(ws.cell(r, c).value or "").replace("\n", " ").strip()
                 for c in range(1, ws.max_column + 1)]
        found = {COLUMNS[re.sub(r"\s+", " ", n)]: i for i, n in enumerate(names)
                 if re.sub(r"\s+", " ", n) in COLUMNS}
        if "packages" in found and "order" in found:
            header_row, idx = r, found
            break
    if header_row is None:
        return []
    rows = []
    for vals in ws.iter_rows(min_row=header_row + 1, values_only=True):
        rows.append({k: vals[i] if i < len(vals) else None for k, i in idx.items()})
    return rows


def read_packages_report(path: str | pathlib.Path) -> list[tuple[str, str, float | None, float]]:
    """[(הזמנה, מק"ט, כמות, מספר אריזות)] מהדוח (PDF או אקסל)."""
    path = pathlib.Path(path)
    raw = _read_pdf(path) if path.suffix.lower() == ".pdf" else _read_excel(path)
    out = []
    for r in raw:
        order, sku, pk = _clean_order(r.get("order")), _clean_sku(r.get("sku")), _num(r.get("packages"))
        if order and sku and pk is not None:
            out.append((order, sku, _num(r.get("qty")), pk))
    if not out:
        raise ValueError(
            f"לא נמצאו שורות עם \"מספר אריזות\" בקובץ {path.name}.\n"
            "צריך את הדוח \"פירוט הזמנות ששודרו לפי נהג\" מפריוריטי (עם העמודות הזמנה, מק\"ט ומספר אריזות)."
        )
    return out


class PackagesLookup:
    """מספר האריזות לשורת הזמנה: לפי הזמנה + מק"ט (ואם יש כמה - לפי הכמות)."""

    def __init__(self, entries: list[tuple[str, str, float | None, float]], source: str = "") -> None:
        self.source = source
        self.by_key: dict[tuple[str, str], list[tuple[float | None, float]]] = {}
        for order, sku, qty, pk in entries:
            self.by_key.setdefault((order, sku), []).append((qty, pk))

    def __len__(self) -> int:
        return sum(len(v) for v in self.by_key.values())

    def get(self, order: Any, sku: Any, qty: Any = None) -> float | None:
        cands = self.by_key.get((_clean_order(order), _clean_sku(sku)))
        if not cands:
            return None
        if len(cands) == 1:
            return cands[0][1]
        q = _num(qty)
        for cq, pk in cands:
            if q is not None and cq is not None and abs(cq - q) < 0.01:
                return pk
        return None  # כמה שורות של אותו מק"ט ואף אחת לא באותה כמות - לא מנחשים
