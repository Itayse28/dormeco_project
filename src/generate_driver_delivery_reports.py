#!/usr/bin/env python3
"""
מערכת ייצור דוחות הובלה לנהגים (דורמקו) - generate_driver_delivery_reports.py
=============================================================================
- קורא את הקובץ הממוין (<ListName>_sorted.xlsx) מתיקיית output_files.
- מייצר תיקייה בתוך output_files בשם: "דוחות הובלה <INPUT_FILE_NAME>"
- לכל נהג X מייצר קובץ אקסל בשם: "הובלות ע״י נהג X.xlsx"
- מותאם אוטומטית להדפסה בעמוד A4 אחד (רוחב וגובה מותאמים לעמוד אחד).
- ההערות בסוף הדוח מופיעות ללא מספור וללא סימני קריאה.
"""

from __future__ import annotations

import argparse
import math
import io
import pathlib
import re
from urllib.parse import quote

try:
    import mail_sender  # מודול השליחה במייל (mail_sender.py באותה תיקייה)
except ImportError:  # בלי המודול הדוחות עדיין נוצרים, רק בלי קובצי השליחה במייל
    mail_sender = None
import subprocess
import sys
import zipfile
from typing import Any

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from config_files import read_local_overrides


# ==============================================================================
# קישורים לפריוריטי (מספר הזמנה / לקוח נפתחים בפריוריטי בלחיצה)
# ==============================================================================
PRIORITY_DEFAULT_SERVER = "https://pri.dormeco.co.il"
# כותרות אפשריות לעמודת מספר הלקוח בדוח (נדרשת לקישור ללקוח)
CUST_NO_HEADERS = [
    "מס' לקוח", 'מס" לקוח', "מספר לקוח", "מס. לקוח", "מס לקוח", "קוד לקוח", "מספר הלקוח",
]


def priority_cfg_values(raw_kv: dict[str, str]) -> dict[str, str]:
    return {
        "PRI_DOR_CODE": str(raw_kv.get("PRI_DOR_CODE", "") or "").strip(),
        "PRIORITY_SERVER_URL": str(
            raw_kv.get("PRIORITY_SERVER_URL", "") or PRIORITY_DEFAULT_SERVER
        ).strip(),
        "PRIORITY_ORDER_FORM": str(raw_kv.get("PRIORITY_ORDER_FORM", "") or "ORDERS").strip(),
        "PRIORITY_CUSTOMER_FORM": str(
            raw_kv.get("PRIORITY_CUSTOMER_FORM", "") or "CUSTOMERS"
        ).strip(),
    }


def find_cust_no_col(headers: list[str]) -> int:
    for name in CUST_NO_HEADERS:
        for idx, h in enumerate(headers):
            if name in str(h or ""):
                return idx
    return -1


def _key_text(key: Any) -> str:
    if isinstance(key, float) and key.is_integer():
        key = int(key)
    return str(key or "").strip()


def priority_link(cfg: dict[str, Any], form: str, key: Any) -> str | None:
    """קישור לפתיחת רשומה בפריוריטי, או None אם חסר קוד חברה / מפתח."""
    code = str(cfg.get("PRI_DOR_CODE", "") or "").strip()
    key_txt = _key_text(key)
    if not code or not key_txt:
        return None
    server = str(cfg.get("PRIORITY_SERVER_URL") or PRIORITY_DEFAULT_SERVER).strip().rstrip("/")
    return (
        f"{server}/priority/openmail.htm?priority:priform@{form}:"
        f"{quote(key_txt, safe='')}:{code}:tabula.ini:1"
    )


def order_link(cfg: dict[str, Any], order_no: Any) -> str | None:
    return priority_link(cfg, cfg.get("PRIORITY_ORDER_FORM") or "ORDERS", order_no)


def customer_link(cfg: dict[str, Any], cust_no: Any) -> str | None:
    return priority_link(cfg, cfg.get("PRIORITY_CUSTOMER_FORM") or "CUSTOMERS", cust_no)


def set_link(cell: Any, url: str | None) -> None:
    """הופך משבצת להיפרלינק (כחול עם קו תחתון), תוך שמירה על גודל/הדגשה."""
    if not url:
        return
    cell.hyperlink = url
    f = cell.font
    cell.font = Font(
        name=f.name, size=f.size, bold=f.bold, color="0563C1", underline="single"
    )

try:
    import pypdf
except ImportError:
    pypdf = None


DEFAULT_BASE_DIR = r"C:\Users\itaik\OneDrive\מסמכים\dormeco-sorting-script"
CONFIG_FILE_NAME = "config_parameters.txt"

MANDATORY_DRIVER_NOTES: list[str] = [
    "נא לשים לב להערות - חובה לפעול לפי סדר משימות הכתוב שינויים - בתאום איתי בלבד",
    "לאחר כל תחילת נסיעה 15 ועד 20 קמ חובה לעצור ולהדק את רצועות הקשירה של הסחורה",
    "לשים לב אם יש אנליזות",
    "אם מתקבל צ'ק חובה להעביר מידי",
    "לשים לב אם יש אנליזות מצורפות לתעודות ואם חסר להודיע",
]

DEFAULT_COOLING_SKUS: set[str] = {
    "BANKBIOPRO10L25",
    "CCCPPP0000",
    "JSCURVE2025",
}

DEFAULT_HAZMAT_SKUS: set[str] = {
    "250000010",
    "ADD2230IBC",
    "ADD642000",
    "AKCASC2830",
    "BPSA4001",
    "BPSA4002",
    "CHACAB3800",
    "CHACAB5510",
    "CSPCAR9500",
    "DIAC10325",
    "DOWDPMA210",
    "DOWka50000",
    "DOWROC2900",
    "DOWSIL8040",
    "ECOCAS2500",
    "ECOCAS8000",
    "EVOTEGbetain",
    "gerform000",
    "huasod0000",
    "KADAC6100IBC",
    "KADAC6244",
    "KADAC6344IBC",
    "KADAMC16220",
    "KADets000",
    "KADKemira0",
    "luxkem0000",
    "nourform01",
    "shasol0000",
    "WeylCareATBS",
}

DEFAULT_GOLDBOND_SKUS: set[str] = {
    "BANKBIOPR3025",
    "BANKBIOPRO10L25",
    "CCCIVC040",
    "CCCPPP0000",
    "INFSHEBUO25",
}


def parse_bool(val: str, default: bool = True) -> bool:
    v = str(val or "").strip().lower()
    if v in ("true", "1", "yes", "y", "כן", "אמת"):
        return True
    if v in ("false", "0", "no", "n", "לא", "שקר"):
        return False
    return default


def load_config(custom_config_path: pathlib.Path | None = None) -> dict[str, Any]:
    script_dir = pathlib.Path(__file__).resolve().parent.parent  # תיקיית הפרויקט (הקוד נמצא ב-src)
    default_base = pathlib.Path(DEFAULT_BASE_DIR)
    base_dir = default_base if default_base.exists() else script_dir

    in_dir = base_dir / "input_files"
    cfg_path = custom_config_path or (in_dir / CONFIG_FILE_NAME)

    raw_kv: dict[str, str] = {}
    if cfg_path.exists():
        try:
            for line in cfg_path.read_text(encoding="utf-8-sig").splitlines():
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    raw_kv[k.strip()] = v.strip()
        except Exception:
            pass
    # הגדרות מקומיות (BASE_DIR, SAVE_AS_PATH...) גוברות על הקובץ הראשי
    raw_kv.update(read_local_overrides(cfg_path))

    configured_base = raw_kv.get("BASE_DIR", DEFAULT_BASE_DIR).strip()
    BASE_FILE = configured_base if configured_base else DEFAULT_BASE_DIR
    INPUT_FOLDER = BASE_FILE + r"\input_files"
    OUTPUT_FOLDER = BASE_FILE + r"\output_files"

    base_path_obj = pathlib.Path(BASE_FILE)
    if not base_path_obj.exists() and script_dir.exists():
        in_dir_obj = script_dir / "input_files"
        out_dir_obj = script_dir / "output_files"
    else:
        in_dir_obj = base_path_obj / "input_files"
        out_dir_obj = base_path_obj / "output_files"

    main_wh_name = raw_kv.get("MAIN_WAREHOUSE_NAME", "אוברסיז").strip()
    if main_wh_name in ("מחסן ראשי", "המחסן הראשי", ""):
        main_wh_name = "אוברסיז"

    return {
        "BASE_DIR": BASE_FILE,
        "INPUT_FOLDER": INPUT_FOLDER,
        "OUTPUT_FOLDER": OUTPUT_FOLDER,
        "INPUT_DIR_PATH": in_dir_obj,
        "OUTPUT_DIR_PATH": out_dir_obj,
        "INPUT_FILE_NAME": raw_kv.get("INPUT_FILE_NAME", "list.xlsx"),
        "COOLING_FILE_NAME": raw_kv.get(
            "COOLING_FILE_NAME", "Cooling_products.xlsx"
        ),
        "HAZMAT_FILE_NAME": raw_kv.get(
            "HAZMAT_FILE_NAME", "dangerous_products.xlsx"
        ),
        "GOLDBOND_FILE_NAME": raw_kv.get(
            "GOLDBOND_FILE_NAME", "Gold_bond_gili.xlsx"
        ),
        "MAIN_WAREHOUSE_NAME": main_wh_name,
        "MAIN_WAREHOUSE_ADDRESS": raw_kv.get(
            "MAIN_WAREHOUSE_ADDRESS", "מבוא הספנים 1 אשדוד"
        ),
        "COOLING_WAREHOUSE_NAME": raw_kv.get(
            "COOLING_WAREHOUSE_NAME", "מחסן קירור"
        ),
        "COOLING_WAREHOUSE_ADDRESS": raw_kv.get(
            "COOLING_WAREHOUSE_ADDRESS", "מבוא הספנים 6 אשדוד"
        ),
        "HAZMAT_WAREHOUSE_NAME": raw_kv.get(
            "HAZMAT_WAREHOUSE_NAME", 'פרידנזון חומ"ס'
        ),
        "HAZMAT_WAREHOUSE_ADDRESS": raw_kv.get(
            "HAZMAT_WAREHOUSE_ADDRESS", "פארק נמלי ישראל, אשדוד"
        ),
        "GOLDBOND_WAREHOUSE_NAME": raw_kv.get(
            "GOLDBOND_WAREHOUSE_NAME", "גולד בונד גילי"
        ),
        "GOLDBOND_WAREHOUSE_ADDRESS": raw_kv.get(
            "GOLDBOND_WAREHOUSE_ADDRESS", "הקידמה 1 אשדוד"
        ),
        "AUTO_OPEN_OUTPUT_FOLDER": parse_bool(
            raw_kv.get("AUTO_OPEN_OUTPUT_FOLDER", "true"), True
        ),
        "EXTRA_COOLING_SKUS": raw_kv.get("EXTRA_COOLING_SKUS", ""),
        "EXTRA_HAZMAT_SKUS": raw_kv.get("EXTRA_HAZMAT_SKUS", ""),
        "EXTRA_GOLDBOND_SKUS": raw_kv.get("EXTRA_GOLDBOND_SKUS", ""),
        # דוחות הובלה של נהגי "אוברסיז <מספר>" נשלחים לכתובות האלה
        # (השם הקודם SEND_SHIPPING_LIST_TO_OVERSEAS עדיין נתמך)
        "SEND_SHIPPING_LIST_TO_OVERSEAS_DRIVERS": (
            mail_sender.parse_email_list(
                mail_sender.kv_get(raw_kv, "SEND_SHIPPING_LIST_TO_OVERSEAS_DRIVERS")
                or mail_sender.kv_get(raw_kv, "SEND_SHIPPING_LIST_TO_OVERSEAS")
            )
            if mail_sender is not None
            else []
        ),
        # כל דוחות ההובלה (של כל הנהגים) נשלחים לכתובות האלה
        "SEND_ALL_SHIPPING_LISTS_TO": (
            mail_sender.parse_email_list(
                mail_sender.kv_get(raw_kv, "SEND_ALL_SHIPPING_LISTS_TO")
            )
            if mail_sender is not None
            else []
        ),
        **priority_cfg_values(raw_kv),
    }


def load_workbook_safe(excel_path: str | pathlib.Path) -> openpyxl.Workbook:
    excel_path = pathlib.Path(excel_path)
    try:
        return openpyxl.load_workbook(excel_path)
    except TypeError:
        bio = io.BytesIO()
        with zipfile.ZipFile(excel_path, "r") as zin, zipfile.ZipFile(
            bio, "w"
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    text = data.decode("utf-8")
                    text = re.sub(
                        r"<([a-zA-Z0-9_]+:)?fill\s*/>",
                        lambda m: (
                            f"<{m.group(1) or ''}fill>"
                            f'<{m.group(1) or ""}patternFill patternType="none"/>'
                            f"</{m.group(1) or ''}fill>"
                        ),
                        text,
                    )
                    data = text.encode("utf-8")
                zout.writestr(item, data)
        bio.seek(0)
        return openpyxl.load_workbook(bio)


def clean_sku_token(raw_val: Any) -> str:
    if raw_val is None:
        return ""
    s = str(raw_val).strip()
    s = re.sub(r'^="?|"?$', "", s).strip()
    return s.upper()


def load_skus_from_excel(xlsx_path: pathlib.Path) -> set[str]:
    found: set[str] = set()
    if not xlsx_path.exists():
        return found
    try:
        wb = load_workbook_safe(xlsx_path)
        for ws in wb.worksheets:
            sku_col_idx = 1
            for r in range(1, min(10, ws.max_row + 1)):
                row_vals = [
                    str(ws.cell(r, c).value or "").strip()
                    for c in range(1, ws.max_column + 1)
                ]
                for c_idx, val in enumerate(row_vals, start=1):
                    if any(
                        k in val
                        for k in ('מק"ט', "מק'ט", "מקט", "SKU", "פריט")
                    ):
                        sku_col_idx = c_idx
                        break
            for r in range(1, ws.max_row + 1):
                token = clean_sku_token(ws.cell(r, sku_col_idx).value)
                if (
                    token
                    and len(token) >= 4
                    and re.match(r"^[A-Z0-9\-_]+$", token)
                    and token not in ("NONE", "NULL")
                ):
                    found.add(token)
    except Exception:
        pass
    return found


def load_cooling_skus_from_file(file_path: pathlib.Path) -> set[str]:
    found: set[str] = set()
    if not file_path.exists():
        return found
    suffix = file_path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return load_skus_from_excel(file_path)
    if suffix == ".pdf" and pypdf is not None:
        try:
            reader = pypdf.PdfReader(str(file_path))
            for page in reader.pages:
                text = page.extract_text() or ""
                for line in text.splitlines():
                    tokens = re.findall(r"\b[A-Z0-9]{6,20}\b", line.upper())
                    if (
                        "GOODS" in line.upper()
                        or re.search(r"\d{2}/\d{2}/\d{4}", line)
                    ) and tokens:
                        for t in tokens:
                            if not t.isdigit():
                                found.add(t)
        except Exception:
            pass
    return found


def load_warehouse_sku_sets(
    cfg: dict[str, Any],
) -> tuple[set[str], set[str], set[str]]:
    # הקבצים במחסנים הם המקור היחיד, והם נקראים מחדש בכל הרצה: מק"ט שנמחק
    # מקובץ כבר לא שייך למחסן. הרשימות המובנות משמשות רק כשקובץ חסר או ריק.
    cooling_skus: set[str] = set()
    hazmat_skus: set[str] = set()
    goldbond_skus: set[str] = set()
    files_found: set[str] = set()

    in_dir: pathlib.Path = cfg["INPUT_DIR_PATH"]
    script_dir = pathlib.Path(__file__).resolve().parent.parent  # תיקיית הפרויקט (הקוד נמצא ב-src)

    for folder in (in_dir, script_dir):
        c_file = folder / cfg["COOLING_FILE_NAME"]
        if not c_file.exists():
            # מחסן קירור: הקובץ הוחלף מ-PDF לאקסל - מחפשים גם בסיומת השנייה
            for alt in ("Cooling_products.xlsx", "Cooling_products.pdf"):
                if (folder / alt).exists():
                    c_file = folder / alt
                    break
        if c_file.exists():
            files_found.add("cooling")
            cooling_skus.update(load_cooling_skus_from_file(c_file))
        h_file = folder / cfg["HAZMAT_FILE_NAME"]
        if h_file.exists():
            files_found.add("hazmat")
            hazmat_skus.update(load_skus_from_excel(h_file))
        g_file = folder / cfg["GOLDBOND_FILE_NAME"]
        if not g_file.exists():
            # מחסן גולד בונד גילי: הקובץ נקרא היום Gold_bond_gili.xlsx (בעבר Gold_bond.xlsx)
            for alt in ("Gold_bond_gili.xlsx", "Gold_bond.xlsx"):
                if (folder / alt).exists():
                    g_file = folder / alt
                    break
        if g_file.exists():
            files_found.add("goldbond")
            goldbond_skus.update(load_skus_from_excel(g_file))

    # קובץ ריק = אין מק"טים במחסן הזה. רק קובץ שלא קיים בכלל -> הרשימה המובנית
    for key, label, skus, defaults in (
        ("cooling", cfg["COOLING_FILE_NAME"], cooling_skus, DEFAULT_COOLING_SKUS),
        ("hazmat", cfg["HAZMAT_FILE_NAME"], hazmat_skus, DEFAULT_HAZMAT_SKUS),
        ("goldbond", cfg["GOLDBOND_FILE_NAME"], goldbond_skus, DEFAULT_GOLDBOND_SKUS),
    ):
        if key not in files_found:
            print(f"[!] קובץ המחסן {label} לא נמצא - משתמשים ברשימה המובנית בקוד.")
            skus.update(s.upper() for s in defaults)

    for extra_c in str(cfg.get("EXTRA_COOLING_SKUS", "")).split(","):
        tok = clean_sku_token(extra_c)
        if tok:
            cooling_skus.add(tok)
    for extra_h in str(cfg.get("EXTRA_HAZMAT_SKUS", "")).split(","):
        tok = clean_sku_token(extra_h)
        if tok:
            hazmat_skus.add(tok)
    for extra_g in str(cfg.get("EXTRA_GOLDBOND_SKUS", "")).split(","):
        tok = clean_sku_token(extra_g)
        if tok:
            goldbond_skus.add(tok)

    return cooling_skus, hazmat_skus, goldbond_skus


# ==============================================================================
# שליחת דוחות ההובלה של נהגי אוברסיז במייל
# ==============================================================================
SHIPPING_LAUNCHER_NAME = "שליחת דוחות הובלה אוברסיז במייל.bat"
ALL_SHIPPING_LAUNCHER_NAME = "שליחת כל דוחות ההובלה במייל.bat"
SHIPPING_SUBJECT_PREFIX = "דוח הובלה עבור "
FAILED_SAVES: list[str] = []  # קבצים שלא נשמרו (למשל כי הם פתוחים באקסל)
REPORT_FILE_PREFIX = "הובלות ע״י "
OVERSEAS_DRIVER_PATTERN = r"נהג אוברסיז\s*\d+"  # "אוברסיז" + מספר


def create_overseas_shipping_launcher(
    day_dir: pathlib.Path, cfg: dict[str, Any]
) -> pathlib.Path | None:
    """
    קובץ bat ששולח במייל (כל דוח במייל נפרד) רק את דוחות ההובלה של נהגי
    "אוברסיז <מספר>" שבתיקייה, אל SEND_SHIPPING_LIST_TO_OVERSEAS_DRIVERS.
    """
    if mail_sender is None:
        return None
    recipients = cfg.get("SEND_SHIPPING_LIST_TO_OVERSEAS_DRIVERS") or []
    has_overseas = any(
        re.search(OVERSEAS_DRIVER_PATTERN, f.stem) for f in day_dir.glob("*.xlsx")
    )
    if not recipients or not has_overseas:
        return None
    return mail_sender.create_launcher(
        folder=day_dir,
        launcher_name=SHIPPING_LAUNCHER_NAME,
        title="שליחת דוחות הובלה אוברסיז במייל",
        default_to=recipients,
        default_name="Overseas drivers",
        name_filter=OVERSEAS_DRIVER_PATTERN,
        strip_prefix=REPORT_FILE_PREFIX,
        subject_prefix=SHIPPING_SUBJECT_PREFIX,
    )


def create_all_shipping_launcher(
    day_dir: pathlib.Path, cfg: dict[str, Any]
) -> pathlib.Path | None:
    """
    קובץ bat ששולח את כל דוחות ההובלה שבתיקייה אל SEND_ALL_SHIPPING_LISTS_TO:
    בלחיצה הדוחות מומרים מאקסל ל-PDF (כולל תיקונים ידניים שנשמרו) ונשלחים במייל אחד.
    """
    if mail_sender is None:
        return None
    recipients = cfg.get("SEND_ALL_SHIPPING_LISTS_TO") or []
    if not recipients:
        return None
    return mail_sender.create_launcher(
        folder=day_dir,
        launcher_name=ALL_SHIPPING_LAUNCHER_NAME,
        title="שליחת כל דוחות ההובלה במייל",
        default_to=recipients,
        default_name="All shipping lists",
        name_filter="^" + REPORT_FILE_PREFIX,
        strip_prefix=REPORT_FILE_PREFIX,
        subject_prefix=SHIPPING_SUBJECT_PREFIX,
        pdf_single_mail=True,  # מומרים ל-PDF בזמן השליחה, כולם במייל אחד
    )


PICKUP_GROUP_NAME = "איסוף עצמי"   # "נהג איסוף עצמי" בטבלה הממוינת - בלי דוח הובלה
NO_DEST_GROUP_NAME = "ללא יעד"      # "נהג ללא יעד" - דוח הובלה עם הכתובת מקובץ הגלם


def extract_driver_number(driver_label: str) -> str:
    """
    מחזיר את שם הנהג מתוך התווית:
    "נהג אברהם כץ (אזור השרון)" -> "אברהם כץ" ,  "נהג 3 (איזור גוש דן)" -> "3"
    """
    drv_str = str(driver_label or "").strip()
    m = re.match(r"^\s*נהג\s+(.+?)\s*(\(|$)", drv_str)
    name = m.group(1).strip() if m else (drv_str or "1")
    # "נהג ללא יעד 23.11.2026" (קבוצה בדוח שבועי) - התאריך נוסף לשם הקובץ בנפרד
    return re.sub(r"\s+\d{1,2}\.\d{1,2}\.\d{4}$", "", name) or name


def safe_file_part(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', " ", str(text)).strip() or "1"


def _num_or_none(row_vals: list[Any], idx: int) -> float | None:
    if not 0 <= idx < len(row_vals) or row_vals[idx] in (None, ""):
        return None
    try:
        return float(str(row_vals[idx]).replace(",", ""))
    except ValueError:
        return None


def create_single_driver_report(
    driver_num: str,
    driver_label: str,
    driver_rows: list[dict[str, Any]],
    output_file: pathlib.Path,
    cfg: dict[str, Any],
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\[\]:*?/\\]", " ", f"נהג {driver_num}")[:31]
    ws.sheet_view.rightToLeft = True

    font_family = "Arial"
    title_font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    title_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )

    sec_font = Font(name=font_family, size=11, bold=True, color="1F4E78")
    sec_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )

    wh_header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    wh_main_fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    wh_special_fill = PatternFill(
        start_color="833C0C", end_color="833C0C", fill_type="solid"
    )

    stop_banner_font = Font(
        name=font_family, size=10, bold=True, color="000000"
    )
    stop_banner_fill = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )

    tbl_header_font = Font(name=font_family, size=12, bold=True, color="1F4E78")
    tbl_header_fill = PatternFill(
        start_color="EDEDED", end_color="EDEDED", fill_type="solid"
    )

    data_font = Font(name=font_family, size=12, bold=False, color="000000")
    bold_font = Font(name=font_family, size=12, bold=True, color="000000")

    note_title_font = Font(name=font_family, size=10, bold=True, color="9C0006")
    note_font = Font(name=font_family, size=10, bold=True, color="9C0006")
    note_fill = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )

    thin_side = Side(style="thin", color="BFBFBF")
    med_navy = Side(style="medium", color="1F4E78")
    med_red = Side(style="medium", color="C00000")

    grid_border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )
    box_border = Border(
        left=med_navy, right=med_navy, top=med_navy, bottom=med_navy
    )
    note_border = Border(
        left=med_red, right=med_red, top=thin_side, bottom=thin_side
    )

    # כמות במשטח + מספר משטחים: רק כשיש נתונים (דוח PDF מפריוריטי)
    has_pallets = any(
        r.get("pallets") is not None or r.get("per_pallet") is not None for r in driver_rows
    )
    pallet_headers = ["מספר משטחים"] if has_pallets else []
    num_cols = 7 + len(pallet_headers)

    # רוחבי עמודות  ומותאמים להדפסה על דף אחד
    col_widths = {1: 34.0, 2: 11.0, 3: 30.0, 4: 10.0, 5: 14.0, 6: 30.0, 7: 22.0}
    if has_pallets:
        col_widths = {1: 34.0, 2: 11.0, 3: 30.0, 4: 10.0, 5: 9.0, 6: 14.0,
                      7: 30.0, 8: 22.0}

    def text_lines(text: Any, col: int) -> int:
        """כמה שורות יתפוס טקסט בעמודה (הערכה לפונט 12) - לחישוב גובה השורה."""
        per_line = max(4, int(col_widths.get(col, 12.0) * 0.85))
        return sum(max(1, -(-len(part) // per_line)) for part in str(text or "").split("\n"))

    def round_pallets(v: Any) -> Any:
        """מספר משטחים עם שבר (למשל 0.13) מעוגל למעלה למספר שלם (1)."""
        return math.ceil(v - 1e-9) if isinstance(v, (int, float)) else v

    def num_format(v: Any) -> str:
        """מספר שלם בלי נקודה בסוף (100 ולא 100.); שבר עם עד 2 ספרות."""
        return "#,##0" if isinstance(v, (int, float)) and float(v).is_integer() else "#,##0.##"

    def pallet_vals(item: dict[str, Any]) -> list[Any]:
        return [round_pallets(item.get("pallets"))] if has_pallets else []

    # צבע רקע לשורות של מוצרים ממחסן שאינו אוברסיז (צבע שונה לכל מחסן)
    wh_row_fills = [
        ("is_cooling", cfg["COOLING_WAREHOUSE_NAME"],
         PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")),
        ("is_hazmat", cfg["HAZMAT_WAREHOUSE_NAME"],
         PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")),
        ("is_goldbond", cfg["GOLDBOND_WAREHOUSE_NAME"],
         PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")),
    ]

    def row_fill(item: dict[str, Any]) -> PatternFill | None:
        for flag, _name, fill in wh_row_fills:
            if item.get(flag):
                return fill
        return None
    mark_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    mark_font = Font(name=font_family, size=12, bold=True, color="9C0006")

    def style_mark(cell: Any) -> None:
        if cell.value:
            cell.fill = mark_fill
            cell.font = mark_font
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # 1. שורה ראשונה: כותרת הדוח "הובלות על ידי נהג X"
    ws.merge_cells(
        start_row=1, start_column=1, end_row=1, end_column=num_cols
    )
    title_text = f"הובלות על ידי נהג {driver_num}"
    # נהג כללי (אוברסיז N / אלירן N): שם וטלפון של הנהג בפועל, כפי שהוזנו בחלון
    actual = next((r.get("actual_driver") for r in driver_rows if r.get("actual_driver")), "")
    if actual:
        title_text += f"  |  נהג: {actual}"
    title_cell = ws.cell(1, 1, title_text)
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, num_cols + 1):
        ws.cell(1, c).fill = title_fill
        ws.cell(1, c).border = box_border
    ws.row_dimensions[1].height = 30

    curr_row = 3

    # ==========================================================================
    # 2. פירוט פריקת החומרים לפי סדר היעדים במסלול
    # ==========================================================================
    ws.merge_cells(
        start_row=curr_row,
        start_column=1,
        end_row=curr_row,
        end_column=num_cols,
    )
    sec2_cell = ws.cell(
        curr_row,
        1,
        "פירוט יעדי פריקה לפי סדר הנסיעה במסלול (כתובת מדויקת וחומרים לפריקה בכל יעד)",
    )
    sec2_cell.font = sec_font
    for c in range(1, num_cols + 1):
        ws.cell(curr_row, c).fill = sec_fill
        ws.cell(curr_row, c).border = grid_border
    ws.row_dimensions[curr_row].height = 22
    curr_row += 1

    curr_row += 1

    stops_grouped: list[tuple[tuple[int, str, str], list[dict[str, Any]]]] = []
    for r in driver_rows:
        s_key = (r["stop_num"], r["dest"], r["cust_name"])
        if not stops_grouped or stops_grouped[-1][0] != s_key:
            stops_grouped.append((s_key, [r]))
        else:
            stops_grouped[-1][1].append(r)

    unload_headers = [
        "פרטי לקוח (שם, כתובת לפריקה, איש קשר)",
        "הזמנה",
        "תיאור החומר",
        "כמות לפריקה",
        *pallet_headers,  # עמודה 7 (כשיש משטחים)
        "מחסן העמסה",
        "הערות למוצר",
        "דגש מיוחד ללקוח",
    ]
    CUST_COLS = 1  # שם לקוח + כתובת + איש קשר בתא אחד, ממוזג לאורך שורות התחנה
    ORDER_COL = 2
    QTY_COL = 4
    PAL_COL = 5  # רק כשיש משטחים
    total_weight = 0.0
    total_pallets = 0.0

    for (stop_num, dest, cust_name), stop_items in stops_grouped:
        first_it = stop_items[0]
        addr_parts = [
            p
            for p in (first_it["address"], first_it["site_name"], dest)
            if p and str(p).strip()
        ]
        full_address = (
            ", ".join(dict.fromkeys(addr_parts)) if addr_parts else dest
        )
        phone_txt = (
            f"  |  טלפון: {first_it['phone']}" if first_it["phone"] else ""
        )
        contact = first_it.get("contact") or ""
        if contact:
            # איש הקשר לאספקה (מהדוח, או מקובץ הכתובות). אם יש לו טלפון - במקום
            # הטלפון הכללי של הלקוח; אחרת שניהם.
            has_phone = any(ch.isdigit() for ch in contact)
            phone_txt = f"  |  איש קשר: {contact}" + ("" if has_phone else phone_txt)
        # ערכי העמודות של הלקוח בטבלה
        stop_address = first_it.get("chosen_address") or full_address
        if contact:
            stop_contact = contact if any(ch.isdigit() for ch in contact) else (
                f"{contact} {first_it['phone']}".strip()
            )
        else:
            stop_contact = first_it["phone"] or ""
        cust_url = customer_link(cfg, first_it.get("cust_no"))

        for c_i, h_txt in enumerate(unload_headers, start=1):
            hc = ws.cell(curr_row, c_i, h_txt)
            hc.font = tbl_header_font
            hc.fill = tbl_header_fill
            hc.border = grid_border
            hc.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        ws.row_dimensions[curr_row].height = 36
        curr_row += 1

        first_item_row = curr_row
        stop_weight = 0.0
        stop_pallets_sum = 0
        # פרטי הלקוח בתא אחד: שם הלקוח בפונט גדול, הכתובת ואיש הקשר בפונט רגיל
        rest_txt = "\n".join(t for t in (stop_address, stop_contact) if t)
        cust_details = CellRichText(
            TextBlock(InlineFont(rFont=font_family, sz=12, b=True,
                                 color="0563C1" if cust_url else "000000",
                                 u="single" if cust_url else None), cust_name),
            *([TextBlock(InlineFont(rFont=font_family, sz=12, color="000000"),
                         "\n" + rest_txt)] if rest_txt else []),
        )
        for item in stop_items:
            vals = [
                cust_details,
                item["order_no"],
                item["prod_desc"],
                item["qty"],
                *pallet_vals(item),
                item["wh_short"],
                item["combined_notes"],
                item.get("client_mark") or None,
            ]
            if isinstance(item["qty"], (int, float)):
                total_weight += item["qty"]
                stop_weight += item["qty"]
            if isinstance(item.get("pallets"), (int, float)):
                total_pallets += round_pallets(item["pallets"])
                stop_pallets_sum += round_pallets(item["pallets"])
            notes_col = 6 + len(pallet_headers)
            fill = row_fill(item)
            for c_i, v in enumerate(vals, start=1):
                c = ws.cell(curr_row, c_i, v)
                c.font = bold_font if c_i in (1, QTY_COL, notes_col) else data_font
                if c_i == ORDER_COL:
                    set_link(c, order_link(cfg, v))
                elif c_i == 1:
                    set_link(c, cust_url)
                c.border = grid_border
                if fill is not None and c_i > CUST_COLS:
                    c.fill = fill
                if (c_i == QTY_COL or (has_pallets and c_i == PAL_COL)) and isinstance(v, (int, float)):
                    c.number_format = num_format(v)
                    c.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                elif c_i == ORDER_COL:
                    c.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                else:
                    c.alignment = Alignment(
                        horizontal="right", vertical="center", wrap_text=True
                    )
            style_mark(ws.cell(curr_row, num_cols))
            lines = max(text_lines(v, c_i) for c_i, v in enumerate(vals, start=1) if c_i > CUST_COLS)
            ws.row_dimensions[curr_row].height = max(24, lines * 17 + 7)
            curr_row += 1

        if curr_row - 1 > first_item_row:
            # פרטי הלקוח פעם אחת לכל התחנה (תא ממוזג לאורך שורות המוצרים)
            for c_i in range(1, CUST_COLS + 1):
                ws.merge_cells(start_row=first_item_row, start_column=c_i,
                               end_row=curr_row - 1, end_column=c_i)
        # מספיק גובה לשלוש שורות פרטי הלקוח (גם בתחנה עם מוצר אחד או שניים)
        n_rows = curr_row - first_item_row
        cust_need = 20 + 17 * (text_lines(rest_txt, 1) if rest_txt else 0) + 8
        have = sum(ws.row_dimensions[r_i].height for r_i in range(first_item_row, curr_row))
        if have < cust_need:
            extra = (cust_need - have) / n_rows
            for r_i in range(first_item_row, curr_row):
                ws.row_dimensions[r_i].height += extra

        # סיכום משקל ומשטחים לתחנה הזאת
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=QTY_COL - 1)
        sl = ws.cell(curr_row, 1, 'סה"כ לתחנה - משקל (ק"ג)' + (" / משטחים" if has_pallets else ""))
        sl.alignment = Alignment(horizontal="left", vertical="center")
        sw = ws.cell(curr_row, QTY_COL, round(stop_weight, 2))
        sw.number_format = num_format(sw.value)
        sw.alignment = Alignment(horizontal="center", vertical="center")
        if has_pallets:
            sp = ws.cell(curr_row, PAL_COL, stop_pallets_sum)
            sp.number_format = num_format(sp.value)
            sp.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, num_cols + 1):
            cc = ws.cell(curr_row, c)
            cc.font = bold_font
            cc.fill = tbl_header_fill
            cc.border = grid_border
        ws.row_dimensions[curr_row].height = 20
        curr_row += 2

    # סיכום משקל של כל ההזמנות (לפני ההערות הקבועות)
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=QTY_COL - 1)
    tl = ws.cell(curr_row, 1, 'סה"כ משקל כל ההזמנות (ק"ג)')
    tl.alignment = Alignment(horizontal="left", vertical="center")
    tv = ws.cell(curr_row, QTY_COL, round(total_weight, 2))
    tv.number_format = num_format(tv.value)
    tv.alignment = Alignment(horizontal="center", vertical="center")
    if has_pallets:
        pv = ws.cell(curr_row, PAL_COL, round(total_pallets, 2))
        pv.number_format = num_format(pv.value)
        pv.alignment = Alignment(horizontal="center", vertical="center")
        pl = ws.cell(curr_row, PAL_COL + 1, "משטחים")
        pl.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, num_cols + 1):
        cc = ws.cell(curr_row, c)
        cc.font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
        cc.fill = title_fill
        cc.border = box_border
    ws.row_dimensions[curr_row].height = 22
    curr_row += 2

    # ==========================================================================
    # 3. הערות והנחיות חובה לנהג (ללא מספור וללא סימני קריאה)
    # ==========================================================================
    ws.merge_cells(
        start_row=curr_row,
        start_column=1,
        end_row=curr_row,
        end_column=num_cols,
    )
    nh_cell = ws.cell(curr_row, 1, "הנחיות והערות חובה לנהג")
    nh_cell.font = note_title_font
    nh_cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, num_cols + 1):
        ws.cell(curr_row, c).fill = note_fill
        ws.cell(curr_row, c).border = note_border
    ws.row_dimensions[curr_row].height = 22
    curr_row += 1

    for note_text in MANDATORY_DRIVER_NOTES:
        clean_note = note_text.replace("!", "").strip()
        ws.merge_cells(
            start_row=curr_row,
            start_column=1,
            end_row=curr_row,
            end_column=num_cols,
        )
        nc = ws.cell(curr_row, 1, clean_note)
        nc.font = note_font
        nc.alignment = Alignment(horizontal="right", vertical="center")
        for c in range(1, num_cols + 1):
            ws.cell(curr_row, c).fill = note_fill
            ws.cell(curr_row, c).border = note_border
        ws.row_dimensions[curr_row].height = 20
        curr_row += 1

    for col_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = w

    # הגדרות עמוד להדפסה אוטומטית על דף A4 אחד בדיוק (רוחב וגובה בעמוד אחד)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(
        left=0.25, right=0.25, top=0.35, bottom=0.35, header=0.2, footer=0.2
    )
    ws.print_options.horizontalCentered = True

    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)


def generate_driver_reports_from_sorted_excel(
    sorted_excel_path: pathlib.Path,
    reports_dir: pathlib.Path,
    cfg: dict[str, Any],
) -> list[pathlib.Path]:
    wb = load_workbook_safe(sorted_excel_path)
    # הנתונים המלאים בגיליון "נתונים מלאים" (הגיליון הראשון הוא תצוגה מצומצמת)
    ws = wb["נתונים מלאים"] if "נתונים מלאים" in wb.sheetnames else wb.active

    headers = [
        str(ws.cell(1, col).value or "").strip()
        for col in range(1, ws.max_column + 1)
    ]

    def find_col(possible_names: list[str], default_idx: int) -> int:
        for name in possible_names:
            for idx, h in enumerate(headers):
                if name in h:
                    return idx
        return default_idx

    idx_driver = find_col(["נהג / משאית", "נהג"], 1)
    idx_route_date = find_col(["תאריך אספקה", "תאריך"], -1)
    idx_stop = find_col(["תחנה במסלול", "תחנה"], 2)
    idx_wh = find_col(["מחסן העמסה למוצר", "מחסן העמסה"], 3)
    idx_cust_name = find_col(["שם לקוח"], 5)
    idx_address = find_col(["כתובת"], 6)
    idx_dest = find_col(["יעד", "עיר"], 7)
    idx_arrival = find_col(["שעת הגעה ליעד", "שעת הגעה"], 8)
    idx_site = find_col(["שם האתר", "אתר"], 10)
    idx_phone = find_col(["מספר טלפון", "טלפון"], 11)
    idx_order = find_col(["הזמנה"], 12)
    idx_sku = find_col(["מק'ט", 'מק"ט', "מקט"], 15)
    idx_prod_desc = find_col(["תאור מוצר", "תיאור מוצר"], 16)
    idx_qty = find_col(["כמות"], 17)
    idx_unit = find_col(["יח'", "יחידות", "יח"], 18)
    idx_cust_note = find_col(["הערה למוצר ללקוח", "הערה למוצר"], 19)
    idx_internal_note = find_col(["הערה פנימית להזמנה", "הערה פנימית"], 20)
    idx_cust_no = find_cust_no_col(headers)
    idx_mark = find_col(["דגש מיוחד ללקוח"], -1)
    # הכתובת שנבחרה בחלון העריכה (מקובץ הכתובות הנכונות) - גוברת על הכתובת מההזמנה
    idx_chosen = find_col(["כתובת אספקה נבחרת"], -1)
    idx_quality = find_col(["דיוק מיקום"], -1)
    idx_contact = find_col(["איש קשר לאספקה"], -1)
    idx_actual = headers.index("נהג בפועל") if "נהג בפועל" in headers else -1
    # משטחים (רק בדוח PDF): התאמה מדויקת, כי "כמות" כבר תופס את עמודת הכמות
    idx_per_pallet = headers.index("כמות במשטח") if "כמות במשטח" in headers else -1
    idx_pallets = headers.index("מספר משטחים") if "מספר משטחים" in headers else -1

    cooling_skus, hazmat_skus, goldbond_skus = load_warehouse_sku_sets(cfg)

    # מפתח = התווית המלאה של המסלול (ייחודית לכל נהג/סבב/יום)
    drivers_map: dict[str, tuple[str, list[dict[str, Any]]]] = {}
    route_meta: dict[str, tuple[str, str]] = {}  # תווית -> (שם נהג, תאריך)

    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(v is not None and str(v).strip() != "" for v in row_vals):
            continue

        first_col_str = str(row_vals[0] or "").strip()
        if 'סה"כ' in first_col_str:
            continue

        driver_label = (
            str(row_vals[idx_driver] or "").strip()
            if idx_driver < len(row_vals)
            else "נהג 1"
        )
        driver_num = extract_driver_number(driver_label)
        if driver_num == PICKUP_GROUP_NAME:
            continue  # איסוף עצמי: רק דוח העמסות, בלי דוח הובלה
        route_date_val = (
            str(row_vals[idx_route_date] or "").strip()
            if 0 <= idx_route_date < len(row_vals)
            else ""
        )
        route_meta.setdefault(driver_label, (driver_num, route_date_val))

        try:
            stop_num = int(row_vals[idx_stop] or 1)
        except (ValueError, TypeError):
            stop_num = 1

        sku_val = (
            clean_sku_token(row_vals[idx_sku])
            if idx_sku < len(row_vals)
            else ""
        )
        wh_col_str = (
            str(row_vals[idx_wh] or "").strip()
            if idx_wh < len(row_vals)
            else ""
        )

        is_cooling = (sku_val in cooling_skus) or ("קירור" in wh_col_str)
        is_hazmat = (sku_val in hazmat_skus) or ('חומ"ס' in wh_col_str)
        is_goldbond = (sku_val in goldbond_skus) or ("גולד בונד" in wh_col_str)
        is_main = not (is_cooling or is_hazmat or is_goldbond)

        wh_names_short: list[str] = []
        if is_cooling:
            wh_names_short.append(cfg["COOLING_WAREHOUSE_NAME"])
        if is_hazmat:
            wh_names_short.append(cfg["HAZMAT_WAREHOUSE_NAME"])
        if is_goldbond:
            wh_names_short.append(cfg["GOLDBOND_WAREHOUSE_NAME"])
        if is_main:
            wh_names_short.append(cfg["MAIN_WAREHOUSE_NAME"])

        qty_val = 0.0
        if idx_qty < len(row_vals) and row_vals[idx_qty] is not None:
            try:
                qty_val = float(row_vals[idx_qty])
            except (ValueError, TypeError):
                qty_val = 0.0

        note1 = (
            str(row_vals[idx_cust_note] or "").strip()
            if 0 <= idx_cust_note < len(row_vals)
            and row_vals[idx_cust_note] is not None
            else ""
        )
        note2 = (
            str(row_vals[idx_internal_note] or "").strip()
            if 0 <= idx_internal_note < len(row_vals)
            and row_vals[idx_internal_note] is not None
            else ""
        )
        combined_notes = " | ".join(p for p in (note1, note2) if p)

        rec_dict = {
            "stop_num": stop_num,
            "cust_no": row_vals[idx_cust_no] if 0 <= idx_cust_no < len(row_vals) else None,
            "client_mark": str(row_vals[idx_mark] or "").strip()
            if 0 <= idx_mark < len(row_vals)
            else "",
            "order_no": str(row_vals[idx_order] or "").strip()
            if idx_order < len(row_vals)
            else "",
            "cust_name": str(row_vals[idx_cust_name] or "").strip()
            if idx_cust_name < len(row_vals)
            else "",
            "address": str(row_vals[idx_address] or "").strip()
            if idx_address < len(row_vals)
            else "",
            "dest": str(row_vals[idx_dest] or "").strip()
            if idx_dest < len(row_vals)
            else "",
            "site_name": str(row_vals[idx_site] or "").strip()
            if idx_site < len(row_vals)
            else "",
            "phone": str(row_vals[idx_phone] or "").strip()
            if idx_phone < len(row_vals)
            else "",
            "arrival_str": str(row_vals[idx_arrival] or "").strip()
            if idx_arrival < len(row_vals)
            else "",
            "sku": sku_val,
            "prod_desc": str(row_vals[idx_prod_desc] or "").strip()
            if idx_prod_desc < len(row_vals)
            else "",
            "qty": qty_val,
            "unit": str(row_vals[idx_unit] or "").strip()
            if idx_unit < len(row_vals)
            else "",
            "is_main": is_main,
            "is_cooling": is_cooling,
            "is_hazmat": is_hazmat,
            "is_goldbond": is_goldbond,
            "wh_short": " + ".join(wh_names_short),
            "combined_notes": combined_notes,
            "chosen_address": str(row_vals[idx_chosen] or "").strip()
            if 0 <= idx_chosen < len(row_vals)
            else "",
            "loc_quality": str(row_vals[idx_quality] or "").strip()
            if 0 <= idx_quality < len(row_vals)
            else "",
            "actual_driver": str(row_vals[idx_actual] or "").strip()
            if 0 <= idx_actual < len(row_vals)
            else "",
            "contact": str(row_vals[idx_contact] or "").strip()
            if 0 <= idx_contact < len(row_vals)
            else "",
            "per_pallet": _num_or_none(row_vals, idx_per_pallet),
            "pallets": _num_or_none(row_vals, idx_pallets),
        }

        if driver_label not in drivers_map:
            drivers_map[driver_label] = (driver_label, [])
        drivers_map[driver_label][1].append(rec_dict)

    reports_dir.mkdir(parents=True, exist_ok=True)
    created_files: list[pathlib.Path] = []

    # לפי סדר הנהגים בקובץ הממוין
    for drv_key in drivers_map.keys():
        drv_label, drv_rows = drivers_map[drv_key]
        drv_num, route_date_val = route_meta[drv_key]
        when = route_date_val.replace("-", ".").replace("/", ".")
        drv_rows.sort(key=lambda x: (x["stop_num"], x["order_no"]))
        file_who = safe_file_part(drv_num) + (f" {when}" if when else "")
        # תת-תיקייה לכל תאריך: "דוחות הובלה 2.11"
        m_day = re.match(r"^\s*(\d{1,2})[-./](\d{1,2})", route_date_val)
        day_folder = (
            f"דוחות הובלה {int(m_day.group(1))}.{int(m_day.group(2))}"
            if m_day
            else "דוחות הובלה ללא תאריך"
        )
        day_dir = reports_dir / day_folder
        day_dir.mkdir(parents=True, exist_ok=True)
        out_file = day_dir / f"הובלות ע״י נהג {file_who}.xlsx"
        try:
            create_single_driver_report(
                driver_num=f"{drv_num} - {when}" if when else drv_num,
                driver_label=drv_label,
                driver_rows=drv_rows,
                output_file=out_file,
                cfg=cfg,
            )
        except PermissionError:
            print(f"[!] לא ניתן לשמור (הקובץ פתוח באקסל?): {out_file.name}")
            FAILED_SAVES.append(out_file.name)
            continue
        created_files.append(out_file)

    return created_files


def open_folder_in_explorer(folder_path: pathlib.Path) -> None:
    try:
        resolved = folder_path.resolve()
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer.exe", str(resolved)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(resolved)])
        else:
            subprocess.Popen(["xdg-open", str(resolved)])
    except Exception as e:
        print(f"[!] לא ניתן היה לפתוח את תיקיית הפלט: {e}")


def is_excluded_support_file(p: pathlib.Path, cfg: dict[str, Any]) -> bool:
    name_low = p.name.lower()
    excluded = {
        cfg["COOLING_FILE_NAME"].lower(),
        cfg["HAZMAT_FILE_NAME"].lower(),
        cfg["GOLDBOND_FILE_NAME"].lower(),
        CONFIG_FILE_NAME.lower(),
        "drivers.xlsx",
        "clients_marks.xlsx",
        "package_sizes.xlsx",  # גודל אריזה לכל מק"ט (דוחות העמסות)
        "clients_with_correct_adresses.xlsx",
        "gold_bond_gili.xlsx",
        "gold_bond.xlsx",
        "dangerous_products.xlsx",
        "cooling_products.pdf",
        "cooling_products.xlsx",
        "gold_bond.xlsx",
        "settings.xlsx",
    }
    return (
        name_low in excluded
        or name_low.endswith("_sorted.xlsx")
        or name_low.startswith("~$")
    )


def resolve_base_list_name_and_sorted_file(
    cli_input: str | None, cfg: dict[str, Any]
) -> tuple[str, pathlib.Path]:
    out_dir: pathlib.Path = cfg["OUTPUT_DIR_PATH"]
    in_dir: pathlib.Path = cfg["INPUT_DIR_PATH"]

    if cli_input:
        p = pathlib.Path(cli_input).expanduser().resolve()
        if p.name.endswith("_sorted.xlsx") and p.exists():
            base_name = p.name[: -len("_sorted.xlsx")]
            return base_name, p
        if p.exists():
            base_name = p.stem
            cand_sorted = out_dir / f"{base_name}_sorted.xlsx"
            if cand_sorted.exists():
                return base_name, cand_sorted
            import sort_orders_by_region as sor

            full_cfg = sor.load_or_create_config()
            sor.sort_orders_excel(p, cand_sorted, full_cfg)
            return base_name, cand_sorted

    default_stem = pathlib.Path(cfg["INPUT_FILE_NAME"]).stem
    cand_default_sorted = out_dir / f"{default_stem}_sorted.xlsx"
    if cand_default_sorted.exists():
        return default_stem, cand_default_sorted

    if out_dir.exists():
        sorted_files = sorted(
            list(out_dir.glob("*_sorted.xlsx")),
            key=lambda f: f.stat().st_mtime,
            reverse=True,
        )
        if sorted_files:
            latest = sorted_files[0].resolve()
            base_name = latest.name[: -len("_sorted.xlsx")]
            return base_name, latest

    if in_dir.exists():
        raw_files = sorted(
            [
                f
                for f in in_dir.glob("*.xlsx")
                if not is_excluded_support_file(f, cfg)
            ],
            key=lambda f: f.stat().st_mtime,
            reverse=True,
        )
        if raw_files:
            raw_p = raw_files[0].resolve()
            base_name = raw_p.stem
            cand_sorted = out_dir / f"{base_name}_sorted.xlsx"
            import sort_orders_by_region as sor

            full_cfg = sor.load_or_create_config()
            sor.sort_orders_excel(raw_p, cand_sorted, full_cfg)
            return base_name, cand_sorted

    raise FileNotFoundError(
        f"לא נמצא קובץ ממוין (*_sorted.xlsx) בתיקייה: {cfg['OUTPUT_FOLDER']}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="ייצור דוחות הובלה לנהגים - דורמקו"
    )
    parser.add_argument(
        "input_pos", nargs="?", default=None, help="נתיב לקובץ אקסל"
    )
    parser.add_argument("-i", "--input", default=None, help="נתיב לקובץ קלט")
    parser.add_argument(
        "--config", default=None, help="נתיב לקובץ config_parameters.txt"
    )
    parser.add_argument(
        "--no-open", action="store_true", help="אל תפתח תיקייה בסיום"
    )
    args = parser.parse_args()

    cfg_path = pathlib.Path(args.config) if args.config else None
    cfg = load_config(cfg_path)

    input_arg = args.input or args.input_pos
    list_name, sorted_excel_path = resolve_base_list_name_and_sorted_file(
        input_arg, cfg
    )

    out_dir: pathlib.Path = cfg["OUTPUT_DIR_PATH"]
    reports_dir = out_dir / f"דוחות הובלה {list_name}"
    # מחיקת דוחות ישנים מהרצה קודמת (שמות/מספר הנהגים יכולים להשתנות בין הרצות)
    if reports_dir.exists():
        old_files = (
            list(reports_dir.rglob("הובלות ע״י נהג *.xlsx"))
            + list(reports_dir.rglob(SHIPPING_LAUNCHER_NAME))
            + list(reports_dir.rglob(ALL_SHIPPING_LAUNCHER_NAME))
        )
        for old in old_files:
            try:
                old.unlink()
            except OSError as exc:
                print(f"[!] לא ניתן למחוק דוח ישן (אולי פתוח באקסל): {old.name} ({exc})")
        for sub in reports_dir.iterdir():
            if sub.is_dir() and not any(sub.iterdir()):
                try:
                    sub.rmdir()  # תת-תיקיית תאריך שהתרוקנה
                except OSError:
                    pass  # תיקייה פתוחה בסייר / נעולה ע"י OneDrive - לא קריטי

    print("=" * 75)
    print(f"[*] קובץ ממוין שנקרא:  {sorted_excel_path}")
    print(f"[*] תיקיית דוחות הובלה: {reports_dir}")
    print("=" * 75)

    created = generate_driver_reports_from_sorted_excel(
        sorted_excel_path=sorted_excel_path,
        reports_dir=reports_dir,
        cfg=cfg,
    )

    print(f"\n[V] נוצרו {len(created)} דוחות הובלה לנהגים בהצלחה:")
    for f_path in created:
        print(f"  - {f_path.name}")

    # קובץ שליחה במייל לדוחות של נהגי אוברסיז (בכל תיקיית תאריך שיש בה כאלה)
    launchers = [
        create_overseas_shipping_launcher(d, cfg)
        for d in sorted({f.parent for f in created})
    ]
    launchers = [x for x in launchers if x]
    if launchers:
        print(
            f"\n[V] נוצר קובץ שליחה של דוחות נהגי אוברסיז ({SHIPPING_LAUNCHER_NAME}) ב:"
        )
        for x in launchers:
            print(f"      - {x.parent.name}")
        print("    נמענים: " + ", ".join(cfg["SEND_SHIPPING_LIST_TO_OVERSEAS_DRIVERS"]))

    # קובץ שליחה של כל דוחות ההובלה (בכל תיקיית תאריך)
    all_launchers = [
        create_all_shipping_launcher(d, cfg) for d in sorted({f.parent for f in created})
    ]
    if any(all_launchers):
        print(f"\n[V] נוצר קובץ שליחה של כל דוחות ההובלה ({ALL_SHIPPING_LAUNCHER_NAME})")
        print("    נמענים: " + ", ".join(cfg["SEND_ALL_SHIPPING_LISTS_TO"]))

    if FAILED_SAVES:
        # שאר הקבצים נוצרו; מדווחים על מה שלא נשמר (קוד יציאה 4 = נוצר חלקית)
        print(
            "\n[!] הקבצים הבאים לא נשמרו כי הם פתוחים באקסל - סגרו אותם והריצו שוב:\n"
            + "\n".join(f"      - {n}" for n in FAILED_SAVES),
            file=sys.stderr,
        )
        sys.exit(4)

    if cfg.get("AUTO_OPEN_OUTPUT_FOLDER", True) and not args.no_open:
        open_folder_in_explorer(out_dir)


if __name__ == "__main__":
    main()