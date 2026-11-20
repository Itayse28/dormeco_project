"""
קריאת דוח "הזמנות מוכנות לאספקה" מפריוריטי בפורמט PDF.

ה-PDF בנוי כטבלה שבה כל תא מצויר כמלבן. לכן כל מילה משויכת למלבן הקטן ביותר
שמכיל אותה, והתאים משויכים לעמודות לפי מלבני כותרת הטבלה. כך מתקבלת טבלה
מדויקת גם כשהטקסט בתא נשבר לכמה שורות.

התוצאה היא טבלה באותו מבנה כמו קובץ האקסל מפריוריטי (אותן כותרות באותו סדר),
ואחריהן עמודות נוספות שיש רק ב-PDF: כתובת אתר, עיר אתר, טלפון אתר, שם איש קשר,
טלפון איש קשר, מחסן מועדף, מספר משטחים, כמות במשטח.
"""

from __future__ import annotations

import datetime
import pathlib
import re
from typing import Any

import openpyxl
from pypdf import PdfReader
from pypdf.generic import ContentStream

# הכותרות של קובץ האקסל (אותו סדר - כך שזיהוי העמודות בסקריפט לא משתנה)
EXCEL_HEADERS = [
    "מס. לקוח", "שם לקוח", "כתובת", "יעד", "שם האתר", "מספר טלפון", "הזמנה",
    "תאריך אספקה", "שורה", "מק'ט", "תאור מוצר", "כמות", "הערה למוצר ללקוח .",
    "יח'", "הערה פנימית להזמנה.",
]
EXTRA_HEADERS = [
    "כתובת אתר", "עיר אתר", "טלפון אתר", "שם איש קשר", "טלפון איש קשר",
    "מחסן מועדף", "מספר משטחים", "כמות במשטח", "דגשים ללקוח.",
]

# כותרת עמודה ב-PDF (אחרי איחוד השורות שלה) -> שם העמודה בטבלה
PDF_COLUMN_NAMES = {
    "הזמנה": "הזמנה",
    "תאריך אספקה": "תאריך אספקה",
    "שורה": "שורה",
    'מק"ט': "מק'ט",
    "תאור מוצר": "תאור מוצר",
    "כמות": "כמות",
    "יח'": "יח'",
    "מחסן מועדף": "מחסן מועדף",
    "מספר משטחים": "מספר משטחים",
    "כמות במשטח": "כמות במשטח",
    "שם האתר": "שם האתר",
    "כתובת אתר": "כתובת אתר",
    "עיר אתר": "עיר אתר",
    "טלפון אתר": "טלפון אתר",
    "שם איש קשר": "שם איש קשר",
    "טלפון איש קשר": "טלפון איש קשר",
    "הערה למוצר ללקוח .": "הערה למוצר ללקוח .",
    "הערה פנימית להזמנה.": "הערה פנימית להזמנה.",
    # פורמט חדש של הדוח: פרטי הלקוח בעמודות (במקום שורת לקוח מעל ההזמנות)
    "מס. לקוח": "מס. לקוח",
    "שם לקוח": "שם לקוח",
    "כתובת": "כתובת",
    "מספר טלפון": "מספר טלפון",
    "יעד": "יעד",
    "דגשים ללקוח.": "דגשים ללקוח.",
}
CUSTOMER_COLUMNS = ("מס. לקוח", "שם לקוח", "כתובת", "מספר טלפון", "יעד")
PHONE_COLUMNS = {"טלפון אתר", "טלפון איש קשר"}
NUMBER_COLUMNS = {"כמות", "מספר משטחים", "כמות במשטח"}

_HEB = re.compile(r"[\u0590-\u05FF]")
_MIRROR = str.maketrans("()[]", ")(][")


class Rect:
    __slots__ = ("x0", "y0", "x1", "y1", "tokens")

    def __init__(self, x0: float, y0: float, x1: float, y1: float) -> None:
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1
        self.tokens: list[tuple[float, float, str]] = []

    @property
    def area(self) -> float:
        return (self.x1 - self.x0) * (self.y1 - self.y0)

    def contains(self, x: float, y: float) -> bool:
        return self.x0 - 0.5 <= x <= self.x1 + 0.5 and self.y0 - 0.5 <= y <= self.y1 + 0.5


def _paths_to_rects(operations: list) -> list[Rect]:
    """כל מסלול סגור (m + l...) או re הופך למלבן התוחם שלו."""
    rects: list[Rect] = []
    pts: list[tuple[float, float]] = []

    def flush() -> None:
        if len(pts) >= 3:
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            if max(xs) - min(xs) > 2 and max(ys) - min(ys) > 2:
                rects.append(Rect(min(xs), min(ys), max(xs), max(ys)))
        pts.clear()

    for args, op in operations:
        if op == b"m":
            flush()
            pts.append((float(args[0]), float(args[1])))
        elif op == b"l":
            pts.append((float(args[0]), float(args[1])))
        elif op == b"re":
            flush()
            x, y, w, h = (float(a) for a in args)
            if w > 2 and h > 2:
                rects.append(Rect(min(x, x + w), min(y, y + h), max(x, x + w), max(y, y + h)))
    flush()
    return rects


class _Placement:
    """
    מיקום של אובייקט Form בעמוד. הדוח נשמר כ"צילום" ארוך של דף אינטרנט שמחולק
    לחלונות: כל עמוד מציג חלקים (לפעמים משותפים לכמה עמודים) עם הזזה וחיתוך משלו.
    הקואורדינטות מנורמלות לקנה המידה של ה-Form (כך שהספים בקוד לא תלויים בהקטנה).
    """

    def __init__(self, name: str, cm: list[float], clip: tuple[float, float, float, float] | None):
        self.name = name
        self.scale = cm[0] or 1.0
        self.dx = cm[4] / self.scale
        self.dy = cm[5] / self.scale
        self.clip = (
            tuple(c / self.scale for c in clip) if clip else None
        )  # (x0, y0, x1, y1) בקואורדינטות מנורמלות

    def point(self, x: float, y: float) -> tuple[float, float]:
        return x + self.dx, y + self.dy

    def visible(self, x: float, y: float) -> bool:
        if self.clip is None:
            return True
        x0, y0, x1, y1 = self.clip
        return x0 - 1 <= x <= x1 + 1 and y0 - 1 <= y <= y1 + 1

    def place_rect(self, r: Rect) -> Rect | None:
        x0, y0 = self.point(r.x0, r.y0)
        x1, y1 = self.point(r.x1, r.y1)
        if self.clip is not None:
            cx0, cy0, cx1, cy1 = self.clip
            x0, y0, x1, y1 = max(x0, cx0), max(y0, cy0), min(x1, cx1), min(y1, cy1)
            if x1 - x0 <= 2 or y1 - y0 <= 2:
                return None
        return Rect(x0, y0, x1, y1)


def _mat_mul(a: list[float], b: list[float]) -> list[float]:
    """מכפלת מטריצות PDF (a אחרי b)."""
    return [
        a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
        a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
        a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5],
    ]


def _placements(page: Any, reader: PdfReader) -> list[_Placement]:
    """כל פקודות Do בעמוד, עם המטריצה וחיתוך ה-re W שלפניהן."""
    out: list[_Placement] = []
    cm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    clip: tuple[float, float, float, float] | None = None
    stack: list[tuple[list[float], Any]] = []
    pending_re: tuple[float, float, float, float] | None = None
    for args, op in ContentStream(page.get_contents(), reader).operations:
        if op == b"q":
            stack.append((list(cm), clip))
        elif op == b"Q" and stack:
            cm, clip = stack.pop()
        elif op == b"cm":
            cm = _mat_mul([float(a) for a in args], cm)
        elif op == b"re":
            x, y, w, h = (float(a) for a in args)
            pending_re = (
                cm[0] * min(x, x + w) + cm[4], cm[3] * min(y, y + h) + cm[5],
                cm[0] * max(x, x + w) + cm[4], cm[3] * max(y, y + h) + cm[5],
            )
        elif op == b"W" and pending_re is not None:
            clip = pending_re if clip is None else (
                max(clip[0], pending_re[0]), max(clip[1], pending_re[1]),
                min(clip[2], pending_re[2]), min(clip[3], pending_re[3]),
            )
        elif op == b"n":
            pending_re = None
        elif op == b"Do":
            out.append(_Placement(str(args[0]), cm, clip))
    return out


def _page_layout(page: Any, reader: PdfReader) -> tuple[list[Rect], list[tuple[float, float, str]]]:
    """מלבני התאים והמילים בעמוד, בקואורדינטות של העמוד (מנורמלות)."""
    resources = page.get("/Resources") or {}
    xobjects = resources.get("/XObject") or {}
    placements = [p for p in _placements(page, reader) if p.name in xobjects]

    rects: list[Rect] = []
    for pl in placements:
        xo = xobjects[pl.name].get_object()
        if xo.get("/Subtype") != "/Form":
            continue
        try:
            ops = ContentStream(xo, reader).operations
        except Exception:
            continue
        for r in _paths_to_rects(ops):
            placed = pl.place_rect(r)
            if placed is not None:
                rects.append(placed)

    tokens: list[tuple[float, float, str]] = []
    current: list[_Placement | None] = [None]
    queue = list(placements)

    def before(op: bytes, args: Any, _cm: Any, _tm: Any) -> None:
        if op == b"Do":
            current[0] = queue.pop(0) if queue else None

    def visit(text: str, cm: list, tm: list, _fd: Any, _fs: Any) -> None:
        if not text or not text.strip():
            return
        x = cm[4] + tm[4] * cm[0]
        y = cm[5] + tm[5] * cm[3]
        pl = current[0]
        if pl is not None:
            x, y = pl.point(x, y)
            if not pl.visible(x, y + 2):
                return
        tokens.append((x, y, text))

    page.extract_text(visitor_operand_before=before, visitor_text=visit)
    return rects, tokens


def _ordered_runs(tokens: list[tuple[float, float, str]], has_hebrew: bool) -> list[str]:
    """
    קטעי הטקסט בסדר קריאה מימין לשמאל. קטעים באותו מיקום (אותו בלוק טקסט) מופיעים
    בקובץ בסדר הפוך. בשורה עם עברית הסוגריים מתהפכים (שיקוף של כתיבה מימין לשמאל).
    """
    indexed = list(enumerate(tokens))
    if not has_hebrew:
        # שורה בלי עברית (אנגלית / מספרים) נקראת משמאל לימין
        indexed.sort(key=lambda it: (round(it[1][0], 1), it[0]))
        return [t for _i, (_x, _y, t) in indexed]
    indexed.sort(key=lambda it: (-round(it[1][0], 1), -it[0]))
    runs = []
    for _i, (_x, _y, t) in indexed:
        if _HEB.search(t) or not re.search(r"[A-Za-z0-9]", t):
            t = t.translate(_MIRROR)
        runs.append(t)
    return runs


def _tidy(text: str) -> str:
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"([(\[])\s+", r"\1", text)
    text = re.sub(r"\s+([)\]])", r"\1", text)
    text = re.sub(r"(?<=[\u0590-\u05FF])\s+'", "'", text)
    return text.strip()


def _line_text(tokens: list[tuple[float, float, str]]) -> str:
    """טקסט של שורה אחת בתא."""
    has_hebrew = any(_HEB.search(t[2]) for t in tokens)
    return _tidy(" ".join(r.strip() for r in _ordered_runs(tokens, has_hebrew) if r.strip()))


def _cell_lines(rect: Rect) -> list[str]:
    """שורות הטקסט בתא (מלמעלה למטה)."""
    lines: list[list[tuple[float, float, str]]] = []
    for tok in sorted(rect.tokens, key=lambda t: -t[1]):
        if lines and abs(lines[-1][0][1] - tok[1]) < 3:
            lines[-1].append(tok)
        else:
            lines.append([tok])
    return [s for s in (_line_text(l) for l in lines) if s]


def _cell_text(rect: Rect, column: str) -> str:
    lines = _cell_lines(rect)
    if column in PHONE_COLUMNS:
        return "".join(lines).replace(" ", "")
    return _tidy(" ".join(lines))


def _parse_number(text: str) -> float | str | None:
    t = text.replace(",", "").strip()
    if not t:
        return None
    try:
        return float(t)
    except ValueError:
        return text


def _parse_date(text: str) -> datetime.datetime | str | None:
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", text or "")
    if not m:
        return text or None
    d, mo, y = (int(g) for g in m.groups())
    if y < 100:
        y += 2000
    try:
        return datetime.datetime(y, mo, d)
    except ValueError:
        return text


def _join_customer_line(rect: Rect) -> str:
    """שורת הלקוח ("מס. לקוח: ..., שם לקוח: ..., כתובת: ..., מספר טלפון: ..., יעד: ...")."""
    text = "".join(_ordered_runs(rect.tokens, True))
    # רווח בין עברית למספר שהודבקו ("החרושת19" -> "החרושת 19")
    text = re.sub(r"(?<=[\u0590-\u05FF\"'.])(?=\d)", " ", text)
    text = re.sub(r"(?<=\d)(?=[\u0590-\u05FF])", " ", text)
    return re.sub(r"\s+", " ", text).strip()


_CUST_RE = re.compile(
    r"מס\.? לקוח\W*(?P<no>\d+)[^\u0590-\u05FF]*שם לקוח\s*:?\s*(?P<name>.*?)\s*,\s*כתובת\s*:\s*"
    r"(?P<addr>.*?)\s*,?\s*מספר טלפון\s*:?\s*(?P<phone>.*?)\s*,?\s*יעד\s*:\s*(?P<dest>.*)$"
)


def parse_customer_line(text: str) -> dict[str, str]:
    m = _CUST_RE.search(text)
    if not m:
        no = re.search(r"מס\.? לקוח\W*(\d+)", text)
        return {"no": no.group(1) if no else "", "name": "", "addr": "", "phone": "", "dest": ""}
    d = {k: (v or "").strip() for k, v in m.groupdict().items()}
    for k in ("name", "addr", "dest"):
        v = d[k].replace(":", " ") if k == "addr" else d[k]
        v = re.sub(r"\(\s*\)", " ", v)
        v = re.sub(r"(?<=\S)\(", " (", _tidy(v))   # "בע"מ( 1400)" -> "בע"מ (1400)"
        v = re.sub(r"ת\.ד\s*\.", "ת.ד.", v)        # "ת.ד . 260" -> "ת.ד. 260"
        v = re.sub(r"\s+,", ",", v)
        d[k] = re.sub(r"\s+", " ", v).strip(" ,.:")
    phone = re.sub(r"[^\d\-+/]", "", d["phone"])
    if phone.endswith("+"):
        phone = "+" + phone[:-1]  # "972-4-8758780+" -> "+972-4-8758780"
    d["phone"] = phone
    return d


def read_orders_pdf(path: str | pathlib.Path) -> tuple[list[str], list[list[Any]]]:
    """קורא את הדוח ומחזיר (כותרות, שורות) במבנה של קובץ האקסל + העמודות הנוספות."""
    reader = PdfReader(str(path))
    headers = EXCEL_HEADERS + EXTRA_HEADERS
    rows: list[dict[str, Any]] = []
    customer: dict[str, str] = {}
    order_no, order_date = "", None
    columns: list[tuple[float, float, str]] | None = None  # (x0, x1, שם)

    for page in reader.pages:
        rects, tokens = _page_layout(page, reader)
        if not rects:
            continue
        # בלי מסגרות של העמוד / הטבלה כולה (רחבות וגבוהות) - רק תאים
        cells = [r for r in rects if not ((r.x1 - r.x0) > 1000 and (r.y1 - r.y0) > 100)]
        for tok in tokens:
            hits = [r for r in cells if r.contains(tok[0] + 1, tok[1] + 2)]
            if hits:
                min(hits, key=lambda r: r.area).tokens.append(tok)

        # כותרת הטבלה: התאים שהטקסט שלהם הוא שם עמודה מוכר
        # (תא בגוף הטבלה יכול להכיל מילה כמו "יח'" - לכן רק השורה עם הכי הרבה שמות)
        by_row: dict[int, list[tuple[float, float, str, float]]] = {}
        for r in cells:
            name = " ".join(_cell_lines(r))
            if name in PDF_COLUMN_NAMES:
                by_row.setdefault(round(r.y1), []).append((r.x0, r.x1, PDF_COLUMN_NAMES[name], r.y0))
        page_cols = max(by_row.values(), key=len) if by_row else []
        if len(page_cols) >= 8:
            header_y = min(c[3] for c in page_cols)
            columns = [(c[0], c[1], c[2]) for c in page_cols]
        else:
            header_y = None
        if not columns:
            continue

        # שורות: תאים מתחת לכותרת, מקובצים לפי הגובה שלהם. שורת לקוח = תא רחב.
        width = max(c[1] for c in columns) - min(c[0] for c in columns)
        body = [
            r for r in cells
            if (header_y is None or r.y1 <= header_y + 1) and r.tokens
        ]
        bands: dict[tuple[int, int], list[Rect]] = {}
        for r in body:
            bands.setdefault((round(r.y1), round(r.y0)), []).append(r)
        for (top, _bottom), band in sorted(bands.items(), key=lambda kv: -kv[0][0]):
            wide = [r for r in band if (r.x1 - r.x0) > width * 0.6]
            if wide:
                text = _join_customer_line(wide[0])
                if "לקוח" in text:
                    customer = parse_customer_line(text)
                    order_no, order_date = "", None
                continue
            values: dict[str, str] = {}
            for r in band:
                mid = (r.x0 + r.x1) / 2
                col = next((c[2] for c in columns if c[0] - 1 <= mid <= c[1] + 1), None)
                if col:
                    values[col] = _cell_text(r, col)
            if not any(values.values()):
                continue
            is_item = bool(values.get("מק'ט") or values.get("שורה"))
            if not is_item:
                # המשך של שורה שנחתכה בסוף העמוד הקודם
                if rows:
                    prev = rows[-1]
                    for k, v in values.items():
                        if v:
                            prev[k] = f"{prev.get(k) or ''} {v}".strip()
                continue
            if values.get("מס. לקוח"):
                # פורמט חדש: פרטי הלקוח בעמודות של השורה (שורות המשך של אותה
                # הזמנה משאירות אותם ריקים - ממשיכים עם הלקוח הקודם)
                no = re.sub(r"\D", "", values["מס. לקוח"])
                if no != customer.get("no"):
                    order_no, order_date = "", None
                customer = {
                    "no": no,
                    "name": _tidy(values.get("שם לקוח", "")),
                    "addr": re.sub(r"ת\.ד\s*\.", "ת.ד.", _tidy(values.get("כתובת", ""))).strip(" ,"),
                    "dest": _tidy(values.get("יעד", "")),
                    "phone": re.sub(r"[^\d\-+/]", "", values.get("מספר טלפון", "")),
                }
                if customer["phone"].endswith("+"):
                    customer["phone"] = "+" + customer["phone"][:-1]
            if values.get("הזמנה"):
                order_no = values["הזמנה"]
            if values.get("תאריך אספקה"):
                order_date = _parse_date(values["תאריך אספקה"])
            row: dict[str, Any] = dict(values)
            row.update({
                "מס. לקוח": customer.get("no", ""),
                "שם לקוח": customer.get("name", ""),
                "כתובת": customer.get("addr", ""),
                "יעד": customer.get("dest", ""),
                "מספר טלפון": customer.get("phone", ""),
                "הזמנה": order_no,
                "תאריך אספקה": order_date,
            })
            rows.append(row)

    out_rows = []
    for row in rows:
        # בלי "כמות במשטח" פריוריטי כותב במספר המשטחים את הכמות עצמה
        # (1,100 ק"ג -> "1,100 משטחים") - זה לא מספר משטחים אמיתי
        if not str(row.get("כמות במשטח") or "").strip():
            row["מספר משטחים"] = None
        vals = []
        for h in headers:
            v = row.get(h)
            if h in NUMBER_COLUMNS and isinstance(v, str):
                v = _parse_number(v)
            elif h == "שורה" and isinstance(v, str):
                v = int(v) if v.strip().isdigit() else (v or None)
            elif isinstance(v, str):
                v = v.strip() or None
            vals.append(v)
        out_rows.append(vals)
    return headers, out_rows


def pdf_to_workbook(path: str | pathlib.Path) -> openpyxl.Workbook:
    headers, rows = read_orders_pdf(path)
    if not rows:
        raise ValueError(
            f"לא נמצאו הזמנות בקובץ ה-PDF {pathlib.Path(path).name} - "
            "האם זה דוח 'הזמנות מוכנות לאספקה' מפריוריטי?"
        )
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "הזמנות"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    return wb
