#!/usr/bin/env python3
"""
מערכת אוטומטית להפקת רשימות ליקוט (Picking Lists) נפרדות לכל נהג ולכל מחסן (Python 3.11/3.12/3.13+)
======================================================================================================
תכונות המערכת:
1. קוראת ישירות את הקובץ הממוין (<ListName>_sorted.xlsx) שנוצר על ידי sort_orders_by_region.py
   בתוך התיקייה BASE_DIR\\output_files (ואם טרם הורץ, מריצה אותו אוטומטית).
2. מפצלת את רשימות הליקוט לפי נהגים (נהג 1, נהג 2 וכו') וגם לפי מחסנים:
   - אוברסיז (מבוא הספנים 1 אשדוד): <ListName>_Driver_<N>_<Region>.xlsx -> שורה ראשונה: "העמסה באוברסיז"
   - מחסן קירור (מבוא הספנים 6 אשדוד): <ListName>_Driver_<N>_<Region>_Cooling.xlsx -> שורה ראשונה: "העמסה במחסן קירור"
   - מחסן חומ"ס (פארק נמלי ישראל, אשדוד): <ListName>_Driver_<N>_<Region>_Hazmat.xlsx -> שורה ראשונה: "העמסה במחסן חומ\"ס"
   - גולד בונד (הקידמה 1 אשדוד): <ListName>_Driver_<N>_<Region>_Gold_Bond.xlsx -> שורה ראשונה: "העמסה בגולד בונד"
3. שומרת את כל רשימות הליקוט בתיקייה: output_files\\רשימות ליקוט <ListName>\\
4. בסיום הריצה פותחת את התיקייה הראשית: BASE_DIR\\output_files.

דרישות התקנה:
    pip install openpyxl requests pandas pypdf
"""

from __future__ import annotations

import argparse
import io
import os
import pathlib
import re
import subprocess
import sys
import zipfile
from typing import Any

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 0. הגדרת נתיבי ברירת מחדל ופרמטרים דינמיים (נטענים מ-config_parameters.txt)
# ==============================================================================
BASE_DIR = r"C:\Users\itaik\OneDrive\מסמכים\dormeco-sorting-script"
INPUT_FOLDER = BASE_DIR + r"\input_files"
OUTPUT_FOLDER = BASE_DIR + r"\output_files"

CONFIG_FILE_NAME = "config_parameters.txt"
INPUT_FILE_NAME = "list.xlsx"
COOLING_FILE_NAME = "Cooling_products.pdf"
HAZMAT_FILE_NAME = "dangerous_products.xlsx"
GOLDBOND_FILE_NAME = "Gold_bond.xlsx"

MAX_TRUCK_WEIGHT_KG: float = 7500.0
AUTO_OPEN_OUTPUT_FOLDER: bool = True

REGION_TO_ENGLISH: dict[str, str] = {
    "צפון מזרח": "North_east",
    "צפון מערב": "North_west",
    "איזור ירושלים": "Jerusalem",
    "אזור השרון": "Sharon",
    "איזור גוש דן": "Gush_Dan",
    "דרום הארץ": "South",
}

DEFAULT_COOLING_SKUS: set[str] = {
    "BANKBIOPRO10L25",
    "CCCPPP0000",
    "JSCURVE2025",
}

DEFAULT_HAZMAT_SKUS: set[str] = {
    "BPSA4001",
    "CCPAGE000",
    "CCPBE11400",
    "FUJNH4F40IBC",
    "KECCMIT0151000",
    "KECCMIT1425",
    "KECCMIT15200",
    "KECCTAC30200",
    "KECGLUT50220F",
    "OCIH2O2351100",
    "OCIH2O235FG1100",
    "OCIH2O2501150",
    "PCCRKNGT5950",
    "POIEPA1234",
    "POIEPA1356",
    "POIEPA1534",
    "POIPC11000",
    "ROTFORIBC",
    "SEMSMK6001",
    "SQM100125",
    "UNIV5005",
    "UNIV5005G",
    "UNIVH20235TECH",
    "UNIVHCLTE30",
    "UNIVNIACC",
    "UNIVNIACD",
    "UNIVPHOSA85DR",
    "UNIVPOHYDTEDR",
    "VTRAAFG0030",
}

DEFAULT_GOLDBOND_SKUS: set[str] = {
    "BANKBIOPR3025",
    "BANKBIOPRO10L25",
    "CCCIVC040",
    "CCCPPP0000",
    "INFSHEBUO25",
}

EXTRA_COOLING_SKUS: set[str] = set()
EXTRA_HAZMAT_SKUS: set[str] = set()
EXTRA_GOLDBOND_SKUS: set[str] = set()


# ==============================================================================
# 1. טעינת הפרמטרים מקובץ config_parameters.txt
# ==============================================================================
def parse_bool(val_str: str, default: bool = False) -> bool:
    v = val_str.strip().lower()
    if v in ("1", "true", "yes", "y", "כן", "אמת", "on"):
        return True
    if v in ("0", "false", "no", "n", "לא", "שקר", "off"):
        return False
    return default


def load_dynamic_config(custom_config_path: str | pathlib.Path | None = None) -> pathlib.Path | None:
    global BASE_DIR, INPUT_FOLDER, OUTPUT_FOLDER
    global INPUT_FILE_NAME, COOLING_FILE_NAME, HAZMAT_FILE_NAME, GOLDBOND_FILE_NAME
    global MAX_TRUCK_WEIGHT_KG, AUTO_OPEN_OUTPUT_FOLDER
    global EXTRA_COOLING_SKUS, EXTRA_HAZMAT_SKUS, EXTRA_GOLDBOND_SKUS

    candidates: list[pathlib.Path] = []
    if custom_config_path:
        candidates.append(pathlib.Path(custom_config_path))
    candidates.append(pathlib.Path(INPUT_FOLDER) / CONFIG_FILE_NAME)
    candidates.append(pathlib.Path(BASE_DIR) / CONFIG_FILE_NAME)
    candidates.append(pathlib.Path(__file__).resolve().parent / "input_files" / CONFIG_FILE_NAME)
    candidates.append(pathlib.Path(__file__).resolve().parent / CONFIG_FILE_NAME)

    cfg_path: pathlib.Path | None = None
    for cand in candidates:
        if cand.exists() and cand.is_file():
            cfg_path = cand
            break

    if cfg_path is None:
        return None

    try:
        raw_text = cfg_path.read_text(encoding="utf-8-sig")
    except Exception:
        try:
            raw_text = cfg_path.read_text(encoding="utf-8")
        except Exception:
            return cfg_path

    kv: dict[str, str] = {}
    for line in raw_text.splitlines():
        line_s = line.strip()
        if not line_s or line_s.startswith("#") or line_s.startswith(";"):
            continue
        if "=" in line_s:
            k, v = line_s.split("=", 1)
            kv[k.strip().upper()] = v.strip().strip('"').strip("'")

    if "BASE_DIR" in kv and kv["BASE_DIR"]:
        BASE_DIR = kv["BASE_DIR"].rstrip("\\/")
        INPUT_FOLDER = BASE_DIR + r"\input_files"
        OUTPUT_FOLDER = BASE_DIR + r"\output_files"

    if "INPUT_FILE_NAME" in kv and kv["INPUT_FILE_NAME"]:
        INPUT_FILE_NAME = kv["INPUT_FILE_NAME"]
    if "COOLING_FILE_NAME" in kv and kv["COOLING_FILE_NAME"]:
        COOLING_FILE_NAME = kv["COOLING_FILE_NAME"]
    if "HAZMAT_FILE_NAME" in kv and kv["HAZMAT_FILE_NAME"]:
        HAZMAT_FILE_NAME = kv["HAZMAT_FILE_NAME"]
    if "GOLDBOND_FILE_NAME" in kv and kv["GOLDBOND_FILE_NAME"]:
        GOLDBOND_FILE_NAME = kv["GOLDBOND_FILE_NAME"]

    if "MAX_TRUCK_WEIGHT_KG" in kv and kv["MAX_TRUCK_WEIGHT_KG"]:
        try:
            MAX_TRUCK_WEIGHT_KG = float(kv["MAX_TRUCK_WEIGHT_KG"])
        except ValueError:
            pass

    if "AUTO_OPEN_OUTPUT_FOLDER" in kv:
        AUTO_OPEN_OUTPUT_FOLDER = parse_bool(
            kv["AUTO_OPEN_OUTPUT_FOLDER"], AUTO_OPEN_OUTPUT_FOLDER
        )

    if "EXTRA_COOLING_SKUS" in kv and kv["EXTRA_COOLING_SKUS"]:
        EXTRA_COOLING_SKUS = {
            s.strip().upper()
            for s in kv["EXTRA_COOLING_SKUS"].split(",")
            if s.strip()
        }
    if "EXTRA_HAZMAT_SKUS" in kv and kv["EXTRA_HAZMAT_SKUS"]:
        EXTRA_HAZMAT_SKUS = {
            s.strip().upper()
            for s in kv["EXTRA_HAZMAT_SKUS"].split(",")
            if s.strip()
        }
    if "EXTRA_GOLDBOND_SKUS" in kv and kv["EXTRA_GOLDBOND_SKUS"]:
        EXTRA_GOLDBOND_SKUS = {
            s.strip().upper()
            for s in kv["EXTRA_GOLDBOND_SKUS"].split(",")
            if s.strip()
        }

    return cfg_path


# ==============================================================================
# 2. פונקציות עזר לקריאה בטוחה של אקסל, סיווג מחסנים וחישוב כמות אריזות
# ==============================================================================
def clean_sku_token(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r'^="?|"?$', "", s).strip()
    return s.upper()


def load_workbook_safe(excel_path: str | pathlib.Path) -> openpyxl.Workbook:
    excel_path = pathlib.Path(excel_path)
    try:
        return openpyxl.load_workbook(excel_path, data_only=True)
    except TypeError:
        bio = io.BytesIO()
        with zipfile.ZipFile(excel_path, "r") as zin, zipfile.ZipFile(bio, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    text = data.decode("utf-8")
                    text = re.sub(
                        r"<([a-zA-Z0-9_]+:)?fill\s*/>",
                        lambda m: (
                            f"<{m.group(1) or ''}fill>"
                            f"<{m.group(1) or ''}patternFill patternType=\"none\"/>"
                            f"</{m.group(1) or ''}fill>"
                        ),
                        text,
                    )
                    data = text.encode("utf-8")
                zout.writestr(item, data)
        bio.seek(0)
        return openpyxl.load_workbook(bio, data_only=True)


def extract_skus_from_excel_file(xlsx_path: pathlib.Path) -> set[str]:
    skus: set[str] = set()
    try:
        wb = load_workbook_safe(xlsx_path)
    except Exception:
        return skus

    header_keywords = {"פריט", 'מק"ט', "מק'ט", "מקט", "SKU", "ITEM", "PART", "קוד פריט"}
    ignored_words = header_keywords | {"סה\"כ", "סהכ", "TOTAL", "GOODS", "מחסן", "איתור", "שם המחסן"}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sku_col_idx: int | None = None
        start_row = 2

        for r in range(1, min(12, ws.max_row + 1)):
            for c in range(1, min(15, ws.max_column + 1)):
                cell_val = clean_sku_token(ws.cell(r, c).value)
                if cell_val in header_keywords:
                    sku_col_idx = c
                    start_row = r + 1
                    break
            if sku_col_idx is not None:
                break

        if sku_col_idx is not None:
            for r in range(start_row, ws.max_row + 1):
                v = clean_sku_token(ws.cell(r, sku_col_idx).value)
                if (
                    v
                    and v not in ignored_words
                    and re.match(r"^[A-Z0-9][A-Z0-9_.-]{3,}$", v)
                    and not v.isdigit()
                ):
                    skus.add(v)
        else:
            for r in range(1, ws.max_row + 1):
                v = clean_sku_token(ws.cell(r, 1).value)
                if (
                    v
                    and v not in ignored_words
                    and re.match(r"^[A-Z][A-Z0-9_.-]{4,}$", v)
                ):
                    skus.add(v)

    return skus


def extract_skus_from_pdf_file(pdf_path: pathlib.Path) -> set[str]:
    skus: set[str] = set()
    pages_text: list[str] = []

    for mod_name in ("pypdf", "PyPDF2"):
        try:
            pdf_mod = __import__(mod_name)
            reader = pdf_mod.PdfReader(str(pdf_path))
            for page in reader.pages:
                pages_text.append(page.extract_text() or "")
            break
        except Exception:
            continue

    if not pages_text:
        try:
            import pdfplumber  # type: ignore

            with pdfplumber.open(str(pdf_path)) as pdf:
                for page in pdf.pages:
                    pages_text.append(page.extract_text() or "")
        except Exception:
            pass

    ignored_tokens = {
        "DORMECO",
        "LTD",
        "GOODS",
        "DAN",
        "HTTP",
        "HTTPS",
        "NETFILES",
        "PRI",
        "HTM",
    }

    for text in pages_text:
        for line in text.splitlines():
            line_s = line.strip()
            if not line_s or "http://" in line_s.lower() or "https://" in line_s.lower():
                continue

            if "Goods" in line_s or re.search(r"\b\d{2}/\d{2}/\d{4}\b", line_s):
                tokens = re.findall(r"\b[A-Z][A-Z0-9_-]{4,}\b", line_s)
                valid_tokens = [
                    t.upper() for t in tokens if t.upper() not in ignored_tokens
                ]
                if valid_tokens:
                    skus.add(valid_tokens[-1])

    return skus


def load_warehouse_catalogs(
    input_folder: str | pathlib.Path,
) -> tuple[set[str], set[str], set[str]]:
    in_dir = pathlib.Path(input_folder)
    cooling_skus: set[str] = set(EXTRA_COOLING_SKUS)
    hazmat_skus: set[str] = set(EXTRA_HAZMAT_SKUS)
    goldbond_skus: set[str] = set(EXTRA_GOLDBOND_SKUS)

    if not in_dir.exists() or not in_dir.is_dir():
        return (
            cooling_skus | DEFAULT_COOLING_SKUS,
            hazmat_skus | DEFAULT_HAZMAT_SKUS,
            goldbond_skus | DEFAULT_GOLDBOND_SKUS,
        )

    cooling_candidates: list[pathlib.Path] = []
    exact_cooling = in_dir / COOLING_FILE_NAME
    if exact_cooling.exists():
        cooling_candidates.append(exact_cooling)
    for f in in_dir.iterdir():
        if f.is_file() and f not in cooling_candidates:
            name_low = f.name.lower()
            if "cooling" in name_low or "קירור" in name_low:
                cooling_candidates.append(f)

    loaded_cooling: set[str] = set()
    for c_file in cooling_candidates:
        suf = c_file.suffix.lower()
        if suf == ".pdf":
            loaded_cooling |= extract_skus_from_pdf_file(c_file)
        elif suf in (".xlsx", ".xlsm"):
            loaded_cooling |= extract_skus_from_excel_file(c_file)
    cooling_skus |= loaded_cooling if loaded_cooling else DEFAULT_COOLING_SKUS

    hazmat_candidates: list[pathlib.Path] = []
    exact_hazmat = in_dir / HAZMAT_FILE_NAME
    if exact_hazmat.exists():
        hazmat_candidates.append(exact_hazmat)
    for f in in_dir.iterdir():
        if f.is_file() and f not in hazmat_candidates:
            name_low = f.name.lower()
            if (
                "dangerous" in name_low
                or "hazmat" in name_low
                or "חומס" in name_low
                or 'חומ"ס' in name_low
            ):
                hazmat_candidates.append(f)

    loaded_hazmat: set[str] = set()
    for h_file in hazmat_candidates:
        suf = h_file.suffix.lower()
        if suf in (".xlsx", ".xlsm"):
            loaded_hazmat |= extract_skus_from_excel_file(h_file)
        elif suf == ".pdf":
            loaded_hazmat |= extract_skus_from_pdf_file(h_file)
    hazmat_skus |= loaded_hazmat if loaded_hazmat else DEFAULT_HAZMAT_SKUS

    goldbond_candidates: list[pathlib.Path] = []
    exact_goldbond = in_dir / GOLDBOND_FILE_NAME
    if exact_goldbond.exists():
        goldbond_candidates.append(exact_goldbond)
    for f in in_dir.iterdir():
        if f.is_file() and f not in goldbond_candidates:
            name_low = f.name.lower()
            if "gold" in name_low or "bond" in name_low or "גולד" in name_low:
                goldbond_candidates.append(f)

    loaded_goldbond: set[str] = set()
    for g_file in goldbond_candidates:
        suf = g_file.suffix.lower()
        if suf in (".xlsx", ".xlsm"):
            loaded_goldbond |= extract_skus_from_excel_file(g_file)
        elif suf == ".pdf":
            loaded_goldbond |= extract_skus_from_pdf_file(g_file)
    goldbond_skus |= loaded_goldbond if loaded_goldbond else DEFAULT_GOLDBOND_SKUS

    return cooling_skus, hazmat_skus, goldbond_skus


def extract_package_weight_kg(
    sku: str | None,
    description: str | None,
    customer_note: str | None = "",
    qty: float = 0.0,
) -> float:
    sku_s = str(sku or "").strip().upper()
    desc_s = str(description or "").strip()

    m_mult = re.search(r"(\d+)X0?(\d+)", sku_s)
    if m_mult:
        count_in_box = int(m_mult.group(1))
        m_desc_dec = re.search(r'(\d+(?:\.\d+)?)\s*(?:ק"ג|ק\'ג|קג|kg|KG)', desc_s)
        if m_desc_dec:
            return count_in_box * float(m_desc_dec.group(1))

    m_desc = re.search(
        r'(\d+(?:\.\d+)?)\s*(?:ק"ג|ק\'ג|קג|kg|KG)', desc_s, re.IGNORECASE
    )
    if m_desc:
        val = float(m_desc.group(1))
        if val >= 1.0:
            return val

    if "SINPRO510A" in sku_s or "SINOPRO 510A" in desc_s.upper():
        return 20.0
    if "UNIVNIACD" in sku_s or "חומצה חנקתית" in desc_s:
        return 290.0

    m_sku_tail = re.search(r"(200|215|220|240|250|300|20|25)[A-Z]?$", sku_s)
    if m_sku_tail:
        return float(m_sku_tail.group(1))

    if qty > 0 and qty % 25 != 0 and qty % 20 == 0:
        return 20.0

    return 25.0


def calculate_packages_count(
    sku: str | None,
    description: str | None,
    qty: float,
    customer_note: str | None = "",
) -> int | float:
    pkg_weight = extract_package_weight_kg(sku, description, customer_note, qty)
    if pkg_weight <= 0:
        return 0
    packages = qty / pkg_weight
    if abs(packages - round(packages)) < 1e-4:
        return int(round(packages))
    return round(packages, 2)


def combine_notes(cust_note: Any, internal_note: Any) -> str:
    n1 = str(cust_note or "").strip()
    n2 = str(internal_note or "").strip()
    if n1 and n2:
        if n1 == n2 or n1 in n2:
            return n2
        if n2 in n1:
            return n1
        return f"{n1} | {n2}"
    return n1 or n2


def parse_driver_num_and_english_region(
    driver_name: str, region_heb: str
) -> tuple[int, str]:
    """
    מחלץ את מספר הנהג (למשל 'נהג 2 (איזור גוש דן)' -> 2)
    ואת שם האזור באנגלית עבור שם הקובץ.
    """
    m_num = re.search(r"נהג\s*(\d+)", str(driver_name))
    d_num = int(m_num.group(1)) if m_num else 1

    m_parens = re.search(r"\(([^)]+)\)", str(driver_name))
    if m_parens:
        inside = m_parens.group(1).strip()
        if "+" in inside:
            parts = [p.strip() for p in inside.split("+")]
            eng_parts = [REGION_TO_ENGLISH.get(p, p) for p in parts]
            return d_num, "_".join(eng_parts)
        if inside in REGION_TO_ENGLISH:
            return d_num, REGION_TO_ENGLISH[inside]

    return d_num, REGION_TO_ENGLISH.get(region_heb, region_heb or "Region")


# ==============================================================================
# 3. יצירת קובץ אקסל של רשימת ליקוט לנהג ולמחסן ספציפי
# ==============================================================================
def write_driver_warehouse_picking_list_excel(
    driver_label: str,
    warehouse_banner_text: str,
    driver_records: list[dict[str, Any]],
    output_file_path: str | pathlib.Path,
    max_truck_weight: float = MAX_TRUCK_WEIGHT_KG,
    sheet_title: str = "Picking_List",
) -> None:
    """
    מייצרת קובץ אקסל של רשימת ליקוט עבור נהג ומחסן ספציפי:
    - שורה 1: "העמסה באוברסיז" / "העמסה במחסן קירור" / "העמסה במחסן חומ\"ס" / "העמסה בגולד בונד"
    - שורה 2: כותרות הטבלה
    - שורה 3 והלאה: פריטי הליקוט ושורת סה"כ
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.sheet_view.rightToLeft = True
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    NAVY_HEADER = "1F4E78"
    BANNER_FILL_HEX = "D9E1F2"
    ZEBRA_FILL = "F9FAFC"
    TOTAL_FILL = "EAEEF7"
    BORDER_GRAY = "D9D9D9"
    OVERWEIGHT_FILL_HEX = "FFC7CE"
    OVERWEIGHT_FONT_HEX = "9C0006"

    banner_fill = PatternFill(
        start_color=BANNER_FILL_HEX, end_color=BANNER_FILL_HEX, fill_type="solid"
    )
    banner_font = Font(name="Arial", size=12, bold=True, color="1F4E78")

    header_fill = PatternFill(
        start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid"
    )
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    overweight_fill = PatternFill(
        start_color=OVERWEIGHT_FILL_HEX, end_color=OVERWEIGHT_FILL_HEX, fill_type="solid"
    )
    overweight_font = Font(name="Arial", size=10, bold=True, color=OVERWEIGHT_FONT_HEX)

    thin_border = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )
    total_border = Border(
        top=Side(style="thin", color=NAVY_HEADER),
        bottom=Side(style="double", color=NAVY_HEADER),
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
    )

    headers = [
        "הזמנה",
        "שם לקוח",
        "מק'ט",
        "תאור מוצר",
        "כמות",
        "אריזות",
        "הערות",
    ]
    numeric_cols = {5, 6}

    # שורה 1: "העמסה בXYZ" (למשל "העמסה באוברסיז")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(row=1, column=1, value=warehouse_banner_text)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = banner_fill
        c.font = banner_font
        c.border = thin_border
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # שורה 2: כותרות הטבלה
    header_row = 2
    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.border = thin_border
        c.alignment = Alignment(
            horizontal="right" if col_idx in numeric_cols else "left",
            vertical="center",
        )
    ws.row_dimensions[header_row].height = 25
    ws.freeze_panes = "A3"

    total_weight = sum(float(r["qty"]) for r in driver_records)
    is_overweight = total_weight > max_truck_weight

    # שורה 3 והלאה: נתוני הליקוט
    start_row = 3
    for idx, rec in enumerate(driver_records):
        r_idx = start_row + idx
        row_values = [
            rec["order_no"],
            rec["cust_name"],
            rec["sku"],
            rec["product_desc"],
            rec["qty"],
            rec["packages"],
            rec["notes"] if rec["notes"] else None,
        ]
        zebra = PatternFill(
            start_color=ZEBRA_FILL if r_idx % 2 == 1 else "FFFFFF",
            end_color=ZEBRA_FILL if r_idx % 2 == 1 else "FFFFFF",
            fill_type="solid",
        )

        for col_idx, val in enumerate(row_values, start=1):
            c = ws.cell(row=r_idx, column=col_idx, value=val)
            c.border = thin_border
            if is_overweight:
                c.fill = overweight_fill
                c.font = overweight_font
            else:
                c.fill = zebra
                c.font = Font(name="Arial", size=10, bold=(col_idx in (1, 5, 6)))

            if col_idx in numeric_cols:
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0"
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

        ws.row_dimensions[r_idx].height = 20

    if driver_records:
        end_row = start_row + len(driver_records) - 1
        total_row = end_row + 1
        ws.cell(
            row=total_row,
            column=1,
            value=f'סה"כ ({driver_label})',
        )
        ws.cell(row=total_row, column=5, value=f"=SUM(E{start_row}:E{end_row})")
        ws.cell(row=total_row, column=6, value=f"=SUM(F{start_row}:F{end_row})")

        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=total_row, column=col_idx)
            c.border = total_border
            if is_overweight:
                c.fill = overweight_fill
                c.font = overweight_font
            else:
                c.fill = PatternFill(
                    start_color=TOTAL_FILL, end_color=TOTAL_FILL, fill_type="solid"
                )
                c.font = Font(name="Arial", size=10, bold=True, color="1F4E78")

            if col_idx in numeric_cols:
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0"
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[total_row].height = 23

        ws.conditional_formatting.add(
            f"A{start_row}:G{total_row}",
            FormulaRule(
                formula=[f"$E${total_row}>{int(max_truck_weight)}"],
                stopIfTrue=True,
                fill=overweight_fill,
                font=overweight_font,
            ),
        )
        ws.conditional_formatting.add(
            f"E{total_row}:E{total_row}",
            CellIsRule(
                operator="greaterThan",
                formula=[str(int(max_truck_weight))],
                stopIfTrue=True,
                fill=overweight_fill,
                font=overweight_font,
            ),
        )

    max_check_row = start_row + len(driver_records) - 1
    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(headers[col_idx - 1]))
        for r_idx in range(start_row, max_check_row + 1):
            val = ws.cell(row=r_idx, column=col_idx).value
            if val is not None:
                val_str = str(val)
                if not val_str.startswith("="):
                    max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 42)

    out_path = pathlib.Path(output_file_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ==============================================================================
# 4. פונקציה ראשית: קריאת הקובץ הממוין מ-output_files והפקת רשימות ליקוט לכל נהג ומחסן
# ==============================================================================
def generate_picking_lists_from_sorted_excel(
    sorted_excel_path: str | pathlib.Path,
    output_dir: str | pathlib.Path = OUTPUT_FOLDER,
    base_list_name: str | None = None,
    max_truck_weight: float = MAX_TRUCK_WEIGHT_KG,
    use_subfolder: bool = True,
) -> tuple[dict[str, pathlib.Path], pathlib.Path]:
    """
    קוראת את הקובץ הממוין (<ListName>_sorted.xlsx) מתוך output_files,
    ומייצרת בתוך output_files\\רשימות ליקוט <ListName>\\ קובץ אקסל נפרד
    לכל נהג (Driver 1, Driver 2...) ולכל מחסן (אוברסיז, קירור, חומ"ס, גולד בונד).
    """
    cooling_skus, hazmat_skus, goldbond_skus = load_warehouse_catalogs(INPUT_FOLDER)

    sorted_path = pathlib.Path(sorted_excel_path)
    wb = load_workbook_safe(sorted_path)
    ws = wb.active

    raw_headers = [
        str(ws.cell(1, col).value or "").strip()
        for col in range(1, ws.max_column + 1)
    ]

    def find_col(possible_names: list[str], default_idx: int) -> int:
        for name in possible_names:
            for idx, h in enumerate(raw_headers):
                if name in h:
                    return idx
        return default_idx

    idx_region = find_col(["אזור חלוקה"], 0)
    idx_driver = find_col(["נהג / משאית", "נהג", "משאית"], 1)
    idx_wh_col = find_col(["מחסן העמסה למוצר", "מחסן העמסה"], 3)
    idx_cust_name = find_col(["שם לקוח"], 5)
    idx_order = find_col(["הזמנה"], 12)
    idx_sku = find_col(["מק'ט", 'מק"ט', "מקט"], 15)
    idx_desc = find_col(["תאור מוצר", "תיאור מוצר"], 16)
    idx_qty = find_col(["כמות"], 17)
    idx_note_cust = find_col(["הערה למוצר ללקוח", "הערה ללקוח"], 19)
    idx_note_int = find_col(["הערה פנימית להזמנה", "הערה פנימית"], 20)

    # קיבוץ שורות לפי נהג בסדר הופעתם בקובץ הממוין
    drivers_order: list[str] = []
    records_by_driver: dict[str, list[dict[str, Any]]] = {}

    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(v is not None and str(v).strip() != "" for v in row_vals):
            continue

        first_col_str = str(row_vals[0] or "").strip()
        if first_col_str.startswith('סה"כ'):
            continue

        region_heb = str(row_vals[idx_region] or "").strip() if idx_region < len(row_vals) else ""
        driver_name = (
            str(row_vals[idx_driver] or "").strip()
            if idx_driver < len(row_vals)
            else ""
        )
        if not driver_name:
            driver_name = f"נהג 1 ({region_heb})"

        order_no = str(row_vals[idx_order] or "").strip() if idx_order < len(row_vals) else ""
        cust_name = str(row_vals[idx_cust_name] or "").strip() if idx_cust_name < len(row_vals) else ""
        sku_raw = str(row_vals[idx_sku] or "").strip() if idx_sku < len(row_vals) else ""
        sku_clean = clean_sku_token(sku_raw)
        product_desc = str(row_vals[idx_desc] or "").strip() if idx_desc < len(row_vals) else ""

        qty_val = 0.0
        if idx_qty < len(row_vals) and row_vals[idx_qty] is not None:
            try:
                qty_val = float(row_vals[idx_qty])
            except (ValueError, TypeError):
                qty_val = 0.0

        qty_clean: int | float = (
            int(round(qty_val)) if abs(qty_val - round(qty_val)) < 1e-4 else qty_val
        )

        cust_note = row_vals[idx_note_cust] if idx_note_cust < len(row_vals) else ""
        int_note = row_vals[idx_note_int] if idx_note_int < len(row_vals) else ""
        notes_combined = combine_notes(cust_note, int_note)

        packages_count = calculate_packages_count(
            sku=sku_clean,
            description=product_desc,
            qty=qty_val,
            customer_note=str(cust_note or ""),
        )

        # זיהוי המחסן של הפריט (לפי עמודת המחסן בקובץ הממוין או לפי המק"ט)
        wh_col_text = (
            str(row_vals[idx_wh_col] or "").strip()
            if idx_wh_col < len(row_vals)
            else ""
        )
        if "קירור" in wh_col_text or sku_clean in cooling_skus:
            wh_key = "cooling"
        elif "גולד בונד" in wh_col_text or sku_clean in goldbond_skus:
            wh_key = "goldbond"
        elif 'חומ"ס' in wh_col_text or "חומס" in wh_col_text or sku_clean in hazmat_skus:
            wh_key = "hazmat"
        else:
            wh_key = "main"

        if driver_name not in records_by_driver:
            records_by_driver[driver_name] = []
            drivers_order.append(driver_name)

        records_by_driver[driver_name].append(
            {
                "region": region_heb,
                "driver_name": driver_name,
                "wh_key": wh_key,
                "order_no": order_no,
                "cust_name": cust_name,
                "sku": sku_clean,
                "product_desc": product_desc,
                "qty": qty_clean,
                "packages": packages_count,
                "notes": notes_combined,
            }
        )

    # חילוץ שם קובץ הבסיס (למשל list מתוך list_sorted.xlsx)
    if base_list_name:
        file_prefix = base_list_name
    else:
        stem = sorted_path.stem
        if stem.lower().endswith("_sorted"):
            file_prefix = stem[:-7]
        else:
            file_prefix = stem

    base_out_dir = pathlib.Path(str(output_dir).rstrip("\\/"))
    if use_subfolder:
        target_out_dir = base_out_dir / f"רשימות ליקוט {file_prefix}"
    else:
        target_out_dir = base_out_dir

    target_out_dir.mkdir(parents=True, exist_ok=True)
    created_files: dict[str, pathlib.Path] = {}

    # הגדרת 4 המחסנים והכיתוב המדויק בשורה הראשונה:
    # (wh_key, file_suffix, row1_banner_text)
    warehouse_configs = [
        ("main", "", "העמסה באוברסיז"),
        ("cooling", "_Cooling", "העמסה במחסן קירור"),
        ("hazmat", "_Hazmat", 'העמסה במחסן חומ"ס'),
        ("goldbond", "_Gold_Bond", "העמסה בגולד בונד"),
    ]

    for driver_name in drivers_order:
        d_records = records_by_driver[driver_name]
        if not d_records:
            continue

        first_reg = d_records[0]["region"]
        d_num, eng_reg = parse_driver_num_and_english_region(driver_name, first_reg)

        for wh_key, file_suffix, banner_text in warehouse_configs:
            wh_records = [r for r in d_records if r["wh_key"] == wh_key]
            if not wh_records:
                continue

            out_filename = f"{file_prefix}_Driver_{d_num}_{eng_reg}{file_suffix}.xlsx"
            out_file_path = target_out_dir / out_filename
            sheet_title = f"D{d_num}_{eng_reg}{file_suffix}"[:31]

            write_driver_warehouse_picking_list_excel(
                driver_label=driver_name,
                warehouse_banner_text=banner_text,
                driver_records=wh_records,
                output_file_path=out_file_path,
                max_truck_weight=max_truck_weight,
                sheet_title=sheet_title,
            )
            created_files[f"Driver_{d_num}_{eng_reg}{file_suffix}"] = out_file_path

    return created_files, target_out_dir


# ==============================================================================
# 5. איתור הקובץ הממוין ב-output_files ופתיחת התיקייה ב-Windows
# ==============================================================================
def open_folder_in_windows_explorer(folder_path: pathlib.Path | str) -> None:
    folder_str = str(pathlib.Path(folder_path))
    if os.name == "nt":
        try:
            os.startfile(folder_str)
            return
        except Exception:
            pass
        try:
            subprocess.run(["explorer", folder_str], check=False)
        except Exception:
            pass


def resolve_sorted_excel_file(
    cli_input: str | None,
) -> tuple[pathlib.Path | None, str | None]:
    """
    מאתרת את הקובץ הממוין (<ListName>_sorted.xlsx) בתוך תיקיית output_files:
    1. אם הועבר קובץ בשורת הפקודה (למשל list.xlsx או list_sorted.xlsx):
       - אם זה כבר קובץ _sorted.xlsx קיים, משתמשת בו ישירות.
       - אם הועבר קובץ המקור (למשל input_files\\list.xlsx), מחפשת את output_files\\list_sorted.xlsx.
    2. אם לא הועבר פרמטר, לוקחת מתיקיית output_files את קובץ ה-*_sorted.xlsx העדכני ביותר
       (או את <INPUT_FILE_NAME[:-5]>_sorted.xlsx).
    3. אם הקובץ הממוין טרם נוצר ב-output_files, מריצה אוטומטית את sort_orders_by_region.py כדי ליצור אותו.
    """
    out_dir = pathlib.Path(OUTPUT_FOLDER)

    if cli_input:
        p = pathlib.Path(cli_input.strip('"'))
        stem = p.stem
        if stem.lower().endswith("_sorted") and p.exists():
            return p, stem[:-7]

        base_prefix = stem[:-7] if stem.lower().endswith("_sorted") else stem
        sorted_candidate = out_dir / f"{base_prefix}_sorted.xlsx"
        if sorted_candidate.exists():
            return sorted_candidate, base_prefix

        # אם הועבר קובץ מקור ועדיין אין לו קובץ _sorted ב-output_files, נריץ את המיון אוטומטית
        if p.exists():
            try:
                import sort_orders_by_region as sor

                sor.load_dynamic_config()
                out_dir.mkdir(parents=True, exist_ok=True)
                sor.sort_orders_excel(input_path=p, output_path=sorted_candidate)
                if sorted_candidate.exists():
                    return sorted_candidate, base_prefix
            except Exception:
                pass

    # בדיקה לפי INPUT_FILE_NAME מתוך config_parameters.txt
    default_prefix = (
        INPUT_FILE_NAME[:-5]
        if INPUT_FILE_NAME.lower().endswith(".xlsx")
        else pathlib.Path(INPUT_FILE_NAME).stem
    )
    default_sorted = out_dir / f"{default_prefix}_sorted.xlsx"
    if default_sorted.exists():
        return default_sorted, default_prefix

    # חיפוש הקובץ הממוין (*_sorted.xlsx) החדש ביותר בתיקיית output_files
    if out_dir.exists() and out_dir.is_dir():
        sorted_files = sorted(
            (
                f
                for f in out_dir.glob("*_sorted.xlsx")
                if not f.name.startswith("~$")
            ),
            key=lambda f: f.stat().st_mtime,
            reverse=True,
        )
        if sorted_files:
            newest = sorted_files[0]
            return newest, newest.stem[:-7]

    # אם לא נמצא קובץ ממוין ב-output_files, ננסה להריץ את sort_orders_by_region על קובץ הקלט
    try:
        import sort_orders_by_region as sor

        sor.load_dynamic_config()
        in_path, in_name = sor.resolve_input_file(cli_input)
        if in_path and in_path.exists():
            prefix = (
                in_name[:-5]
                if in_name and in_name.lower().endswith(".xlsx")
                else in_path.stem
            )
            generated_sorted = out_dir / f"{prefix}_sorted.xlsx"
            out_dir.mkdir(parents=True, exist_ok=True)
            sor.sort_orders_excel(input_path=in_path, output_path=generated_sorted)
            if generated_sorted.exists():
                return generated_sorted, prefix
    except Exception:
        pass

    return None, None


def main() -> None:
    load_dynamic_config()

    parser = argparse.ArgumentParser(
        description="הפקת רשימות ליקוט לפי נהגים ולפי מחסנים מתוך הקובץ הממוין ב-output_files"
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default=None,
        help="נתיב לקובץ ההזמנות או לקובץ הממוין (אופציונלי)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        help="תיקיית הפלט הראשית (ברירת מחדל: BASE_DIR\\output_files)",
    )
    parser.add_argument(
        "-c",
        "--config",
        default=None,
        help="נתיב לקובץ הגדרות מותאם אישית (ברירת מחדל: input_files\\config_parameters.txt)",
    )
    parser.add_argument(
        "--no-subfolder",
        action="store_true",
        help="שמירה ישירות בתיקיית הפלט ללא תיקיית משנה 'רשימות ליקוט'",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="ביטול פתיחה אוטומטית של תיקיית הפלט בסיום",
    )
    args = parser.parse_args()

    if args.config:
        load_dynamic_config(args.config)

    eff_output_dir = args.output_dir if args.output_dir else OUTPUT_FOLDER

    sorted_path, base_prefix = resolve_sorted_excel_file(args.input_file)
    if not sorted_path or not sorted_path.exists():
        print("לא נמצא קובץ הזמנות ממוין בתיקיית output_files.")
        print("אנא הרץ תחילה את sort_orders_by_region.py או את run_all.bat.")
        return

    created, target_dir = generate_picking_lists_from_sorted_excel(
        sorted_excel_path=sorted_path,
        output_dir=eff_output_dir,
        base_list_name=base_prefix,
        max_truck_weight=MAX_TRUCK_WEIGHT_KG,
        use_subfolder=not args.no_subfolder,
    )

    print(f"\n=======================================================")
    print(f"הפקת רשימות הליקוט לפי נהגים ומחסנים מתוך '{sorted_path.name}' הושלמה בהצלחה!")
    print(f"הקבצים נשמרו בתיקייה:\n  {target_dir}\n")
    for key_label, fpath in created.items():
        print(f"  - [{key_label}]: {fpath.name}")
    print(f"=======================================================\n")

    # פתיחת התיקייה הראשית BASE_DIR\output_files בסיום
    if AUTO_OPEN_OUTPUT_FOLDER and not args.no_open:
        open_folder_in_windows_explorer(eff_output_dir)


if __name__ == "__main__":
    main()