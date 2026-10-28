#!/usr/bin/env python3
"""
מערכת אוטומטית למיון הזמנות, העמסה ממחסנים (רגיל / קירור / חומ"ס) ותכנון מסלולי נהגים (Python 3.11/3.12/3.13+)
=================================================================================================================
תכונות המערכת:
1. טעינה דינמית של כל הפרמטרים מקובץ הגדרות בתיקיית input_files (config_parameters.txt),
   כולל זמני העמסה נפרדים לכל מחסן:
   - מחסן רגיל (מבוא הספנים 1 אשדוד): 60 דקות כברירת מחדל (MAIN_WAREHOUSE_LOAD_MINUTES).
   - מחסן קירור (מבוא הספנים 6 אשדוד): 30 דקות כברירת מחדל (COOLING_WAREHOUSE_LOAD_MINUTES).
   - מחסן חומ"ס (פארק נמלי ישראל, אשדוד): 30 דקות כברירת מחדל (HAZMAT_WAREHOUSE_LOAD_MINUTES).
   - פריקה אצל כל לקוח במסלול: 30 דקות כברירת מחדל (UNLOAD_MINUTES).
2. זיהוי אוטומטי של מוצרי קירור מתוך ה-PDF בתיקיית input_files (Cooling_products.pdf)
   וזיהוי מוצרי חומ"ס מתוך קובץ האקסל בתיקיית input_files (dangerous_products.xlsx) לפי מק"ט.
3. חישוב זמני נסיעה אמיתיים בכבישי ישראל ב-3 שכבות:
   - שכבה 1: Google Maps Distance Matrix API (אם הוזן GOOGLE_MAPS_API_KEY).
   - שכבה 2: OSRM Real Israeli Road Network API (חינמי, מבוסס על מיפוי כבישי ישראל ב-OpenStreetMap).
   - שכבה 3: מודל כבישי ישראל מכויל למשאיות (לעבודה ללא אינטרנט / גיבוי).
4. הגבלת זמן משמרת נהג ל-12 שעות (מיציאה מנתניה ב-05:30 ועד חזרה לנתניה) ומשקל משאית ל-7,500 ק"ג:
   חריגה באזור כלשהו מפצלת אותו אוטומטית ל-2 נהגים/משאיות (במקום סימון באדום).
5. איחוד חכם של "צפון מזרח + אזור השרון" או "צפון מערב + אזור השרון" בדרך חזרה לנתניה (ללא שילוב שני הצפונים יחד).
6. סידור עמודות בטבלה הסופית:
   - עמודת "שעת הגעה ליעד" צמודה לעמודת ה"יעד".
   - עמודת "יחידות" צמודה לעמודת ה"כמות", ורק אחריהן העמודה "הערה למוצר ללקוח".

דרישות התקנה:
    pip install openpyxl requests pandas pypdf
"""

from __future__ import annotations

import argparse
import datetime
import io
import itertools
import json
import math
import os
import pathlib
import re
import subprocess
import sys
import time
import zipfile
from typing import Any

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests

# ==============================================================================
# 0. הגדרת נתיבי ברירת מחדל ופרמטרים דינמיים (ניתנים לעדכון ב-config_parameters.txt)
# ==============================================================================
BASE_DIR = r"C:\Users\itaik\OneDrive\מסמכים\dormeco-sorting-script"
INPUT_FOLDER = BASE_DIR + r"\input_files"
OUTPUT_FOLDER = BASE_DIR + r"\output_files"

CONFIG_FILE_NAME = "config_parameters.txt"
INPUT_FILE_NAME = "list.xlsx"
COOLING_FILE_NAME = "Cooling_products.pdf"
HAZMAT_FILE_NAME = "dangerous_products.xlsx"

# זמני העמסה נפרדים לכל מחסן באשדוד וזמן פריקה אצל לקוח (בדקות)
MAIN_WAREHOUSE_LOAD_MINUTES: float = 60.0     # שעה העמסה ראשונית במבוא הספנים 1
COOLING_WAREHOUSE_LOAD_MINUTES: float = 30.0  # 30 דקות העמסה במחסן קירור (מבוא הספנים 6)
HAZMAT_WAREHOUSE_LOAD_MINUTES: float = 30.0   # 30 דקות העמסה במחסן חומ"ס (פארק נמלי ישראל)
UNLOAD_MINUTES: float = 30.0                  # 30 דקות פריקה בכל עצירה אצל לקוח
ALWAYS_LOAD_AT_MAIN_WAREHOUSE: bool = True

MAX_TRUCK_WEIGHT_KG: float = 7500.0
MAX_DRIVER_SHIFT_HOURS: float = 12.0

DRIVER_HOME_NAME: str = "מרכז נתניה"
DRIVER_HOME_COORDS: tuple[float, float] = (32.3215, 34.8532)
DRIVER_DEPARTURE_TIME_STR: str = "05:30"
DRIVER_START_MINUTES: float = 5 * 60 + 30  # 05:30 = 330 דקות מחצות

MAIN_WAREHOUSE_ADDRESS: str = "מבוא הספנים 1 אשדוד"
MAIN_WAREHOUSE_COORDS: tuple[float, float] = (31.8315, 34.6465)

COOLING_WAREHOUSE_ADDRESS: str = "מבוא הספנים 6 אשדוד"
COOLING_WAREHOUSE_COORDS: tuple[float, float] = (31.8322, 34.6482)

HAZMAT_WAREHOUSE_ADDRESS: str = "פארק נמלי ישראל, אשדוד"
HAZMAT_WAREHOUSE_COORDS: tuple[float, float] = (31.8245, 34.6610)

OSRM_TRUCK_SPEED_FACTOR: float = 1.20
ENABLE_NORTH_SHARON_MERGE: bool = True
USE_EXTERNAL_ROUTING_API: bool = True
PREFER_SITE_NAME_CITY: bool = False
AUTO_OPEN_OUTPUT_FOLDER: bool = True

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

EXTRA_COOLING_SKUS: set[str] = set()
EXTRA_HAZMAT_SKUS: set[str] = set()

# ==============================================================================
# 1. טעינה ויצירה אוטומטית של קובץ הפרמטרים הדינמיים ב-input_files
# ==============================================================================
DEFAULT_CONFIG_TEMPLATE = """# ==============================================================================
# קובץ הגדרות ופרמטרים דינמיים - מערכת מיון הזמנות ותכנון מסלולים (דורמקו)
# יש לשמור קובץ זה בתוך תיקיית input_files בשם: config_parameters.txt
# ניתן לשנות כל ערך אחרי סימן השווה (=) ולשמור.
# ==============================================================================

# --- 1. נתיבי תיקיות ושמות קבצים ---
BASE_DIR = C:\\Users\\itaik\\OneDrive\\מסמכים\\dormeco-sorting-script
INPUT_FILE_NAME = list.xlsx
COOLING_FILE_NAME = Cooling_products.pdf
HAZMAT_FILE_NAME = dangerous_products.xlsx

# --- 2. זמני העמסה בכל מחסן בנפרד וזמן פריקה אצל לקוחות (בדקות) ---
# זמן העמסה ראשונית במחסן הרגיל (מבוא הספנים 1 אשדוד) - ברירת מחדל: 60 דקות (שעה)
MAIN_WAREHOUSE_LOAD_MINUTES = 60
# זמן העמסה במחסן קירור (מבוא הספנים 6 אשדוד) - ברירת מחדל: 30 דקות
COOLING_WAREHOUSE_LOAD_MINUTES = 30
# זמן העמסה במחסן חומ"ס (פארק נמלי ישראל, אשדוד) - ברירת מחדל: 30 דקות
HAZMAT_WAREHOUSE_LOAD_MINUTES = 30
# זמן פריקת סחורה בכל עצירה אצל לקוח במסלול - ברירת מחדל: 30 דקות
UNLOAD_MINUTES = 30
# האם הנהג תמיד עוצר תחילה במחסן הראשי (מבוא הספנים 1) להעמסה ראשונית לפני קירור/חומ"ס
ALWAYS_LOAD_AT_MAIN_WAREHOUSE = true

# --- 3. מגבלות משאית ונהג (חריגה תגרום לפיצול אוטומטי ל-2 נהגים) ---
MAX_TRUCK_WEIGHT_KG = 7500
MAX_DRIVER_SHIFT_HOURS = 12.0

# --- 4. שעת יציאה ומיקום בית הנהגים ---
DRIVER_HOME_NAME = מרכז נתניה
DRIVER_HOME_LAT = 32.3215
DRIVER_HOME_LON = 34.8532
DRIVER_DEPARTURE_TIME = 05:30

# --- 5. כתובות וקואורדינטות של המחסנים באשדוד ---
MAIN_WAREHOUSE_ADDRESS = מבוא הספנים 1 אשדוד
MAIN_WAREHOUSE_LAT = 31.8315
MAIN_WAREHOUSE_LON = 34.6465

COOLING_WAREHOUSE_ADDRESS = מבוא הספנים 6 אשדוד
COOLING_WAREHOUSE_LAT = 31.8322
COOLING_WAREHOUSE_LON = 34.6482

HAZMAT_WAREHOUSE_ADDRESS = פארק נמלי ישראל, אשדוד
HAZMAT_WAREHOUSE_LAT = 31.8245
HAZMAT_WAREHOUSE_LON = 34.6610

# --- 6. פרמטרים לחישוב כבישים ואופטימיזציית מסלולים ---
OSRM_TRUCK_SPEED_FACTOR = 1.20
ENABLE_NORTH_SHARON_MERGE = true
USE_EXTERNAL_ROUTING_API = true
PREFER_SITE_NAME_CITY = false
AUTO_OPEN_OUTPUT_FOLDER = true
GOOGLE_MAPS_API_KEY = 

# --- 7. הוספת מק"טים ידנית לקירור או חומ"ס (אופציונלי, מופרדים בפסיק) ---
EXTRA_COOLING_SKUS = 
EXTRA_HAZMAT_SKUS = 
"""


def parse_bool(val_str: str, default: bool = False) -> bool:
    v = val_str.strip().lower()
    if v in ("1", "true", "yes", "y", "כן", "אמת", "on"):
        return True
    if v in ("0", "false", "no", "n", "לא", "שקר", "off"):
        return False
    return default


def parse_time_to_minutes(time_str: str, default_minutes: float = 330.0) -> float:
    m = re.match(r"^(\d{1,2}):(\d{2})$", time_str.strip())
    if m:
        return float(int(m.group(1)) * 60 + int(m.group(2)))
    return default_minutes


def load_dynamic_config(custom_config_path: str | pathlib.Path | None = None) -> pathlib.Path | None:
    """
    טוענת את קובץ ההגדרות מתוך input_files\\config_parameters.txt.
    אם הקובץ עדיין לא קיים בתיקיית input_files, יוצרת אותו אוטומטית כדי שיהיה מוכן לעריכה.
    """
    global BASE_DIR, INPUT_FOLDER, OUTPUT_FOLDER
    global INPUT_FILE_NAME, COOLING_FILE_NAME, HAZMAT_FILE_NAME
    global MAIN_WAREHOUSE_LOAD_MINUTES, COOLING_WAREHOUSE_LOAD_MINUTES, HAZMAT_WAREHOUSE_LOAD_MINUTES
    global UNLOAD_MINUTES, ALWAYS_LOAD_AT_MAIN_WAREHOUSE
    global MAX_TRUCK_WEIGHT_KG, MAX_DRIVER_SHIFT_HOURS
    global DRIVER_HOME_NAME, DRIVER_HOME_COORDS, DRIVER_DEPARTURE_TIME_STR, DRIVER_START_MINUTES
    global MAIN_WAREHOUSE_ADDRESS, MAIN_WAREHOUSE_COORDS
    global COOLING_WAREHOUSE_ADDRESS, COOLING_WAREHOUSE_COORDS
    global HAZMAT_WAREHOUSE_ADDRESS, HAZMAT_WAREHOUSE_COORDS
    global OSRM_TRUCK_SPEED_FACTOR, ENABLE_NORTH_SHARON_MERGE
    global USE_EXTERNAL_ROUTING_API, PREFER_SITE_NAME_CITY, AUTO_OPEN_OUTPUT_FOLDER
    global EXTRA_COOLING_SKUS, EXTRA_HAZMAT_SKUS

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
        target_dir = pathlib.Path(INPUT_FOLDER)
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            cfg_path = target_dir / CONFIG_FILE_NAME
            cfg_path.write_text(DEFAULT_CONFIG_TEMPLATE, encoding="utf-8-sig")
        except Exception:
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

    # תמיכה גם בפרמטר הישן LOAD_MINUTES וגם ב-3 הפרמטרים הנפרדים לכל מחסן
    if "LOAD_MINUTES" in kv and kv["LOAD_MINUTES"]:
        try:
            legacy_load = float(kv["LOAD_MINUTES"])
            COOLING_WAREHOUSE_LOAD_MINUTES = legacy_load
            HAZMAT_WAREHOUSE_LOAD_MINUTES = legacy_load
        except ValueError:
            pass

    if "MAIN_WAREHOUSE_LOAD_MINUTES" in kv and kv["MAIN_WAREHOUSE_LOAD_MINUTES"]:
        try:
            MAIN_WAREHOUSE_LOAD_MINUTES = float(kv["MAIN_WAREHOUSE_LOAD_MINUTES"])
        except ValueError:
            pass
    if "COOLING_WAREHOUSE_LOAD_MINUTES" in kv and kv["COOLING_WAREHOUSE_LOAD_MINUTES"]:
        try:
            COOLING_WAREHOUSE_LOAD_MINUTES = float(kv["COOLING_WAREHOUSE_LOAD_MINUTES"])
        except ValueError:
            pass
    if "HAZMAT_WAREHOUSE_LOAD_MINUTES" in kv and kv["HAZMAT_WAREHOUSE_LOAD_MINUTES"]:
        try:
            HAZMAT_WAREHOUSE_LOAD_MINUTES = float(kv["HAZMAT_WAREHOUSE_LOAD_MINUTES"])
        except ValueError:
            pass
    if "UNLOAD_MINUTES" in kv and kv["UNLOAD_MINUTES"]:
        try:
            UNLOAD_MINUTES = float(kv["UNLOAD_MINUTES"])
        except ValueError:
            pass
    if "ALWAYS_LOAD_AT_MAIN_WAREHOUSE" in kv:
        ALWAYS_LOAD_AT_MAIN_WAREHOUSE = parse_bool(
            kv["ALWAYS_LOAD_AT_MAIN_WAREHOUSE"], ALWAYS_LOAD_AT_MAIN_WAREHOUSE
        )

    if "MAX_TRUCK_WEIGHT_KG" in kv and kv["MAX_TRUCK_WEIGHT_KG"]:
        try:
            MAX_TRUCK_WEIGHT_KG = float(kv["MAX_TRUCK_WEIGHT_KG"])
        except ValueError:
            pass
    if "MAX_DRIVER_SHIFT_HOURS" in kv and kv["MAX_DRIVER_SHIFT_HOURS"]:
        try:
            MAX_DRIVER_SHIFT_HOURS = float(kv["MAX_DRIVER_SHIFT_HOURS"])
        except ValueError:
            pass

    if "DRIVER_HOME_NAME" in kv and kv["DRIVER_HOME_NAME"]:
        DRIVER_HOME_NAME = kv["DRIVER_HOME_NAME"]
    home_lat, home_lon = DRIVER_HOME_COORDS
    if "DRIVER_HOME_LAT" in kv and kv["DRIVER_HOME_LAT"]:
        try:
            home_lat = float(kv["DRIVER_HOME_LAT"])
        except ValueError:
            pass
    if "DRIVER_HOME_LON" in kv and kv["DRIVER_HOME_LON"]:
        try:
            home_lon = float(kv["DRIVER_HOME_LON"])
        except ValueError:
            pass
    DRIVER_HOME_COORDS = (home_lat, home_lon)

    if "DRIVER_DEPARTURE_TIME" in kv and kv["DRIVER_DEPARTURE_TIME"]:
        DRIVER_DEPARTURE_TIME_STR = kv["DRIVER_DEPARTURE_TIME"]
        DRIVER_START_MINUTES = parse_time_to_minutes(
            DRIVER_DEPARTURE_TIME_STR, DRIVER_START_MINUTES
        )

    if "MAIN_WAREHOUSE_ADDRESS" in kv and kv["MAIN_WAREHOUSE_ADDRESS"]:
        MAIN_WAREHOUSE_ADDRESS = kv["MAIN_WAREHOUSE_ADDRESS"]
    m_lat, m_lon = MAIN_WAREHOUSE_COORDS
    if "MAIN_WAREHOUSE_LAT" in kv and kv["MAIN_WAREHOUSE_LAT"]:
        try:
            m_lat = float(kv["MAIN_WAREHOUSE_LAT"])
        except ValueError:
            pass
    if "MAIN_WAREHOUSE_LON" in kv and kv["MAIN_WAREHOUSE_LON"]:
        try:
            m_lon = float(kv["MAIN_WAREHOUSE_LON"])
        except ValueError:
            pass
    MAIN_WAREHOUSE_COORDS = (m_lat, m_lon)

    if "COOLING_WAREHOUSE_ADDRESS" in kv and kv["COOLING_WAREHOUSE_ADDRESS"]:
        COOLING_WAREHOUSE_ADDRESS = kv["COOLING_WAREHOUSE_ADDRESS"]
    c_lat, c_lon = COOLING_WAREHOUSE_COORDS
    if "COOLING_WAREHOUSE_LAT" in kv and kv["COOLING_WAREHOUSE_LAT"]:
        try:
            c_lat = float(kv["COOLING_WAREHOUSE_LAT"])
        except ValueError:
            pass
    if "COOLING_WAREHOUSE_LON" in kv and kv["COOLING_WAREHOUSE_LON"]:
        try:
            c_lon = float(kv["COOLING_WAREHOUSE_LON"])
        except ValueError:
            pass
    COOLING_WAREHOUSE_COORDS = (c_lat, c_lon)

    if "HAZMAT_WAREHOUSE_ADDRESS" in kv and kv["HAZMAT_WAREHOUSE_ADDRESS"]:
        HAZMAT_WAREHOUSE_ADDRESS = kv["HAZMAT_WAREHOUSE_ADDRESS"]
    h_lat, h_lon = HAZMAT_WAREHOUSE_COORDS
    if "HAZMAT_WAREHOUSE_LAT" in kv and kv["HAZMAT_WAREHOUSE_LAT"]:
        try:
            h_lat = float(kv["HAZMAT_WAREHOUSE_LAT"])
        except ValueError:
            pass
    if "HAZMAT_WAREHOUSE_LON" in kv and kv["HAZMAT_WAREHOUSE_LON"]:
        try:
            h_lon = float(kv["HAZMAT_WAREHOUSE_LON"])
        except ValueError:
            pass
    HAZMAT_WAREHOUSE_COORDS = (h_lat, h_lon)

    if "OSRM_TRUCK_SPEED_FACTOR" in kv and kv["OSRM_TRUCK_SPEED_FACTOR"]:
        try:
            OSRM_TRUCK_SPEED_FACTOR = float(kv["OSRM_TRUCK_SPEED_FACTOR"])
        except ValueError:
            pass
    if "ENABLE_NORTH_SHARON_MERGE" in kv:
        ENABLE_NORTH_SHARON_MERGE = parse_bool(
            kv["ENABLE_NORTH_SHARON_MERGE"], ENABLE_NORTH_SHARON_MERGE
        )
    if "USE_EXTERNAL_ROUTING_API" in kv:
        USE_EXTERNAL_ROUTING_API = parse_bool(
            kv["USE_EXTERNAL_ROUTING_API"], USE_EXTERNAL_ROUTING_API
        )
    if "PREFER_SITE_NAME_CITY" in kv:
        PREFER_SITE_NAME_CITY = parse_bool(
            kv["PREFER_SITE_NAME_CITY"], PREFER_SITE_NAME_CITY
        )
    if "AUTO_OPEN_OUTPUT_FOLDER" in kv:
        AUTO_OPEN_OUTPUT_FOLDER = parse_bool(
            kv["AUTO_OPEN_OUTPUT_FOLDER"], AUTO_OPEN_OUTPUT_FOLDER
        )
    if "GOOGLE_MAPS_API_KEY" in kv and kv["GOOGLE_MAPS_API_KEY"]:
        os.environ["GOOGLE_MAPS_API_KEY"] = kv["GOOGLE_MAPS_API_KEY"]

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

    return cfg_path


# ==============================================================================
# 2. קריאה אוטומטית של קובץ מוצרי הקירור (PDF/Excel) וקובץ החומ"ס (Excel/CSV)
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
    ignored_words = header_keywords | {"סה\"כ", "סהכ", "TOTAL", "GOODS", "מחסן", "איתור"}

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


def load_cooling_and_hazmat_catalogs(
    input_folder: str | pathlib.Path,
) -> tuple[set[str], set[str]]:
    in_dir = pathlib.Path(input_folder)
    cooling_skus: set[str] = set(DEFAULT_COOLING_SKUS) | set(EXTRA_COOLING_SKUS)
    hazmat_skus: set[str] = set(DEFAULT_HAZMAT_SKUS) | set(EXTRA_HAZMAT_SKUS)

    if not in_dir.exists() or not in_dir.is_dir():
        return cooling_skus, hazmat_skus

    cooling_candidates: list[pathlib.Path] = []
    exact_cooling = in_dir / COOLING_FILE_NAME
    if exact_cooling.exists():
        cooling_candidates.append(exact_cooling)
    for f in in_dir.iterdir():
        if f.is_file() and f not in cooling_candidates:
            name_low = f.name.lower()
            if "cooling" in name_low or "קירור" in name_low:
                cooling_candidates.append(f)

    for c_file in cooling_candidates:
        suf = c_file.suffix.lower()
        if suf == ".pdf":
            cooling_skus |= extract_skus_from_pdf_file(c_file)
        elif suf in (".xlsx", ".xlsm"):
            cooling_skus |= extract_skus_from_excel_file(c_file)

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

    for h_file in hazmat_candidates:
        suf = h_file.suffix.lower()
        if suf in (".xlsx", ".xlsm"):
            hazmat_skus |= extract_skus_from_excel_file(h_file)
        elif suf == ".pdf":
            hazmat_skus |= extract_skus_from_pdf_file(h_file)

    return cooling_skus, hazmat_skus


# ==============================================================================
# 3. הגדרת סדר אזורי החלוקה ומילון יישובים
# ==============================================================================
REGION_ORDER: list[str] = [
    "צפון מזרח",
    "צפון מערב",
    "איזור ירושלים",
    "אזור השרון",
    "איזור גוש דן",
    "דרום הארץ",
]

REGION_RANK: dict[str, int] = {name: idx for idx, name in enumerate(REGION_ORDER)}
CACHE_FILE = pathlib.Path("geocoding_cache.json")

KNOWN_LOCALITIES: dict[str, tuple[str, float, float]] = {
    # --- 1. צפון מזרח ---
    "קריית שמונה": ("צפון מזרח", 33.2073, 35.5702),
    "קרית שמונה": ("צפון מזרח", 33.2073, 35.5702),
    "צפת": ("צפון מזרח", 32.9646, 35.4960),
    "טבריה": ("צפון מזרח", 32.7922, 35.5312),
    "עפולה": ("צפון מזרח", 32.6078, 35.2897),
    "בית שאן": ("צפון מזרח", 32.4971, 35.4974),
    "נצרת": ("צפון מזרח", 32.6996, 35.3035),
    "נוף הגליל": ("צפון מזרח", 32.7058, 35.3268),
    "נצרת עילית": ("צפון מזרח", 32.7058, 35.3268),
    "מגדל העמק": ("צפון מזרח", 32.6750, 35.2414),
    "קצרין": ("צפון מזרח", 32.9926, 35.6915),
    "חצור הגלילית": ("צפון מזרח", 32.9814, 35.5458),
    "ראש פינה": ("צפון מזרח", 32.9689, 35.5425),
    "מטולה": ("צפון מזרח", 33.2775, 35.5739),
    "צמח": ("צפון מזרח", 32.7042, 35.5861),
    "עמק הירדן": ("צפון מזרח", 32.7042, 35.5861),
    "עמק המעיינות": ("צפון מזרח", 32.4900, 35.5100),
    "כפר תבור": ("צפון מזרח", 32.6869, 35.4214),
    "יבנאל": ("צפון מזרח", 32.7044, 35.5036),
    "מגדל": ("צפון מזרח", 32.8392, 35.5161),
    "רמת הגולן": ("צפון מזרח", 32.9926, 35.6915),
    "גולן": ("צפון מזרח", 32.9926, 35.6915),
    "אלון תבור": ("צפון מזרח", 32.6414, 35.3536),
    "ציפורית": ("צפון מזרח", 32.7480, 35.3440),
    "תל חי": ("צפון מזרח", 33.2356, 35.5769),
    "קדמת גליל": ("צפון מזרח", 32.7830, 35.4650),

    # --- 2. צפון מערב ---
    "בקה אלגרבייה": ("צפון מערב", 32.4189, 35.0408),
    "באקה אל-גרבייה": ("צפון מערב", 32.4189, 35.0408),
    "באקה אל גרביה": ("צפון מערב", 32.4189, 35.0408),
    "חדרה": ("צפון מערב", 32.4340, 34.9197),
    "אור עקיבא": ("צפון מערב", 32.5064, 34.9200),
    "קיסריה": ("צפון מערב", 32.5089, 34.9092),
    "פרדס חנה": ("צפון מערב", 32.4739, 34.9706),
    "פרדס חנה כרכור": ("צפון מערב", 32.4739, 34.9706),
    "בנימינה": ("צפון מערב", 32.5214, 34.9469),
    "זכרון יעקב": ("צפון מערב", 32.5708, 34.9517),
    "אום אל פחם": ("צפון מערב", 32.5186, 35.1536),
    "אום אל-פחם": ("צפון מערב", 32.5186, 35.1536),
    "יקנעם": ("צפון מערב", 32.6594, 35.1100),
    "יקנעם עילית": ("צפון מערב", 32.6594, 35.1100),
    "עתלית": ("צפון מערב", 32.6872, 34.9383),
    "דלית אל כרמל": ("צפון מערב", 32.6936, 35.0469),
    "עספיא": ("צפון מערב", 32.7194, 35.0650),
    "טירת כרמל": ("צפון מערב", 32.7603, 34.9719),
    "חיפה": ("צפון מערב", 32.7940, 34.9896),
    "נשר": ("צפון מערב", 32.7708, 35.0394),
    "מפרץ חיפה": ("צפון מערב", 32.8120, 35.0520),
    "קריית טבעון": ("צפון מערב", 32.7236, 35.1275),
    "קרית טבעון": ("צפון מערב", 32.7236, 35.1275),
    "קריית אתא": ("צפון מערב", 32.8064, 35.1050),
    "קרית אתא": ("צפון מערב", 32.8064, 35.1050),
    "קריית ביאליק": ("צפון מערב", 32.8275, 35.0858),
    "קרית ביאליק": ("צפון מערב", 32.8275, 35.0858),
    "קריית מוצקין": ("צפון מערב", 32.8331, 35.0778),
    "קרית מוצקין": ("צפון מערב", 32.8331, 35.0778),
    "קריית ים": ("צפון מערב", 32.8492, 35.0689),
    "קרית ים": ("צפון מערב", 32.8492, 35.0689),
    "שפרעם": ("צפון מערב", 32.8056, 35.1694),
    "טמרה": ("צפון מערב", 32.8536, 35.1978),
    "סחנין": ("צפון מערב", 32.8639, 35.2972),
    "סח'נין": ("צפון מערב", 32.8639, 35.2972),
    "עראבה": ("צפון מערב", 32.8519, 35.3375),
    "בר לב": ("צפון מערב", 32.9056, 35.1889),
    "כרמיאל": ("צפון מערב", 32.9136, 35.2961),
    "עכו": ("צפון מערב", 32.9278, 35.0817),
    "מילואות": ("צפון מערב", 33.0250, 35.1050),
    "נהריה": ("צפון מערב", 33.0058, 35.0989),
    "תפן": ("צפון מערב", 32.9772, 35.2747),
    "מעלות": ("צפון מערב", 33.0167, 35.2708),
    "מעלות תרשיחא": ("צפון מערב", 33.0167, 35.2708),
    "קורן": ("צפון מערב", 33.0200, 35.2800),
    "שלומי": ("צפון מערב", 33.0722, 35.1436),

    # --- 3. איזור ירושלים ---
    "מודיעין": ("איזור ירושלים", 31.8969, 35.0104),
    "מודיעין מכבים רעות": ("איזור ירושלים", 31.8969, 35.0104),
    "מודיעין עילית": ("איזור ירושלים", 31.9333, 35.0419),
    "בית שמש": ("איזור ירושלים", 31.7456, 34.9867),
    "הר טוב": ("איזור ירושלים", 31.7610, 34.9940),
    "צור הדסה": ("איזור ירושלים", 31.7192, 35.0972),
    "אבו גוש": ("איזור ירושלים", 31.8058, 35.1092),
    "בית נקופה": ("איזור ירושלים", 31.8028, 35.1258),
    "מבשרת ציון": ("איזור ירושלים", 31.8019, 35.1525),
    "גבעת זאב": ("איזור ירושלים", 31.8614, 35.1681),
    "עטרות": ("איזור ירושלים", 31.8550, 35.2150),
    "ירושלים": ("איזור ירושלים", 31.7683, 35.2137),
    "ביתר עילית": ("איזור ירושלים", 31.6978, 35.1156),
    "אפרת": ("איזור ירושלים", 31.6547, 35.1508),
    "גוש עציון": ("איזור ירושלים", 31.6547, 35.1300),
    "מעלה אדומים": ("איזור ירושלים", 31.7772, 35.2992),
    "מישור אדומים": ("איזור ירושלים", 31.7970, 35.3360),
    "שער בנימין": ("איזור ירושלים", 31.8780, 35.2570),

    # --- 4. אזור השרון ---
    "כפר קאסם": ("אזור השרון", 32.1142, 34.9764),
    "רמת השרון": ("אזור השרון", 32.1461, 34.8394),
    "הרצליה": ("אזור השרון", 32.1624, 34.8447),
    "ג'לג'וליה": ("אזור השרון", 32.1547, 34.9536),
    "הוד השרון": ("אזור השרון", 32.1558, 34.8947),
    "רעננה": ("אזור השרון", 32.1848, 34.8713),
    "כפר סבא": ("אזור השרון", 32.1780, 34.9065),
    "בצרה": ("אזור השרון", 32.2150, 34.8780),
    "שפיים": ("אזור השרון", 32.2194, 34.8286),
    "געש": ("אזור השרון", 32.2281, 34.8322),
    "יקום": ("אזור השרון", 32.2478, 34.8419),
    "כוכב יאיר": ("אזור השרון", 32.2269, 34.9964),
    "צור יגאל": ("אזור השרון", 32.2220, 34.9960),
    "טירה": ("אזור השרון", 32.2342, 34.9503),
    "תל מונד": ("אזור השרון", 32.2536, 34.9175),
    "בני דרור": ("אזור השרון", 32.2611, 34.9006),
    "אבן יהודה": ("אזור השרון", 32.2697, 34.8875),
    "קדימה": ("אזור השרון", 32.2792, 34.9156),
    "קדימה צורן": ("אזור השרון", 32.2792, 34.9156),
    "טייבה": ("אזור השרון", 32.2661, 35.0089),
    "קלנסווה": ("אזור השרון", 32.2850, 34.9811),
    "נתניה": ("אזור השרון", 32.2770, 34.8640),
    "תנובות": ("אזור השרון", 32.3092, 34.9686),
    "ניצני עוז": ("אזור השרון", 32.3080, 35.0060),
    "כפר יונה": ("אזור השרון", 32.3167, 34.9358),
    "עמק חפר": ("אזור השרון", 32.3820, 34.9070),

    # --- 5. איזור גוש דן והמרכז ---
    "יבנה": ("איזור גוש דן", 31.8781, 34.7394),
    "מזכרת בתיה": ("איזור גוש דן", 31.8536, 34.8414),
    "קריית עקרון": ("איזור גוש דן", 31.8608, 34.8236),
    "קרית עקרון": ("איזור גוש דן", 31.8608, 34.8236),
    "רחובות": ("איזור גוש דן", 31.8928, 34.8113),
    "נס ציונה": ("איזור גוש דן", 31.9292, 34.7989),
    "באר יעקב": ("איזור גוש דן", 31.9406, 34.8353),
    "רמלה": ("איזור גוש דן", 31.9275, 34.8625),
    "לוד": ("איזור גוש דן", 31.9514, 34.8881),
    "ראשון לציון": ("איזור גוש דן", 31.9730, 34.7925),
    "ראשלצ": ("איזור גוש דן", 31.9730, 34.7925),
    'ראשל"צ': ("איזור גוש דן", 31.9730, 34.7925),
    "איירפורט סיטי": ("איזור גוש דן", 31.9917, 34.9189),
    "קריית שדה התעופה": ("איזור גוש דן", 31.9917, 34.9189),
    "שוהם": ("איזור גוש דן", 31.9986, 34.9456),
    "קרית אסם": ("איזור גוש דן", 31.9920, 34.9380),
    "בית דגן": ("איזור גוש דן", 32.0019, 34.8297),
    "חולון": ("איזור גוש דן", 32.0158, 34.7874),
    'אזה"ת חולון': ("איזור גוש דן", 32.0110, 34.8050),
    "בת ים": ("איזור גוש דן", 32.0132, 34.7480),
    "אזור": ("איזור גוש דן", 32.0242, 34.8053),
    "אור יהודה": ("איזור גוש דן", 32.0292, 34.8572),
    "יהוד": ("איזור גוש דן", 32.0331, 34.8881),
    "יהוד מונוסון": ("איזור גוש דן", 32.0331, 34.8881),
    "סביון": ("איזור גוש דן", 32.0469, 34.8772),
    "גני תקווה": ("איזור גוש דן", 32.0600, 34.8736),
    "קריית אונו": ("איזור גוש דן", 32.0622, 34.8553),
    "קרית אונו": ("איזור גוש דן", 32.0622, 34.8553),
    "גבעת שמואל": ("איזור גוש דן", 32.0781, 34.8475),
    "תל אביב": ("איזור גוש דן", 32.0853, 34.7818),
    "תל אביב יפו": ("איזור גוש דן", 32.0853, 34.7818),
    "תל-אביב": ("איזור גוש דן", 32.0853, 34.7818),
    "גבעתיים": ("איזור גוש דן", 32.0722, 34.8089),
    "רמת גן": ("איזור גוש דן", 32.0823, 34.8106),
    "בני ברק": ("איזור גוש דן", 32.0849, 34.8352),
    "אלעד": ("איזור גוש דן", 32.0519, 34.9514),
    "פתח תקווה": ("איזור גוש דן", 32.0889, 34.8864),
    "פתח-תקווה": ("איזור גוש דן", 32.0889, 34.8864),
    "ראש העין": ("איזור גוש דן", 32.0958, 34.9567),
    "ברקן": ("איזור גוש דן", 32.1060, 35.1060),
    "אריאל": ("איזור גוש דן", 32.1047, 35.1844),

    # --- 6. דרום הארץ ---
    "אשדוד": ("דרום הארץ", 31.8044, 34.6553),
    "עד הלום": ("דרום הארץ", 31.7750, 34.6660),
    "גן יבנה": ("דרום הארץ", 31.7872, 34.7064),
    "כנות": ("דרום הארץ", 31.7764, 34.7564),
    "גדרה": ("דרום הארץ", 31.8147, 34.7797),
    "באר טוביה": ("דרום הארץ", 31.7370, 34.7450),
    "קריית מלאכי": ("דרום הארץ", 31.7372, 34.7480),
    "קרית מלאכי": ("דרום הארץ", 31.7372, 34.7480),
    "אשקלון": ("דרום הארץ", 31.6440, 34.5800),
    "קריית גת": ("דרום הארץ", 31.6020, 34.7820),
    "קרית גת": ("דרום הארץ", 31.6020, 34.7820),
    "שדרות": ("דרום הארץ", 31.5250, 34.5956),
    "נתיבות": ("דרום הארץ", 31.4231, 34.5886),
    "רהט": ("דרום הארץ", 31.3953, 34.7544),
    "להבים": ("דרום הארץ", 31.3728, 34.8169),
    "לקיה": ("דרום הארץ", 31.3283, 34.8586),
    "מיתר": ("דרום הארץ", 31.3264, 34.9389),
    "חורה": ("דרום הארץ", 31.2961, 34.9367),
    "אופקים": ("דרום הארץ", 31.3142, 34.6203),
    "עומר": ("דרום הארץ", 31.2672, 34.8494),
    "באר שבע": ("דרום הארץ", 31.2530, 34.7915),
    "תל שבע": ("דרום הארץ", 31.2469, 34.8292),
    "מושב עמיעוז": ("דרום הארץ", 31.2461, 34.4125),
    "עמיעוז": ("דרום הארץ", 31.2461, 34.4125),
    "אשכול": ("דרום הארץ", 31.2461, 34.4125),
    "כסייפה": ("דרום הארץ", 31.2439, 35.0919),
    "ערד": ("דרום הארץ", 31.2589, 35.2128),
    "רמת חובב": ("דרום הארץ", 31.1336, 34.7889),
    "נאות חובב": ("דרום הארץ", 31.1336, 34.7889),
    "דימונה": ("דרום הארץ", 31.0694, 35.0333),
    "ירוחם": ("דרום הארץ", 30.9881, 34.9314),
    "מצפה רמון": ("דרום הארץ", 30.6103, 34.8019),
    "אילת": ("דרום הארץ", 29.5577, 34.9519),
}

SPECIFIC_ADDRESS_COORDS: list[tuple[str, str, float, float]] = [
    ("ראשון לציון", "אליהו איתן", 31.9868, 34.7695),
    ("ראשון לציון", "שמוטקין", 31.9702, 34.8065),
    ("חיפה", "חלוצי התעשיה", 32.8055, 35.0480),
    ("מפרץ חיפה", "הקיטור", 32.8135, 35.0540),
    ("אשדוד", "מבוא הספנים 1", 31.8315, 34.6465),
    ("אשדוד", "מבוא הספנים 6", 31.8322, 34.6482),
    ("אשדוד", "פארק נמלי ישראל", 31.8245, 34.6610),
]

CBS_SUBDISTRICT_TO_REGION: dict[str, tuple[str, float, float]] = {
    "גולן": ("צפון מזרח", 32.9926, 35.6915),
    "צפת": ("צפון מזרח", 32.9646, 35.4960),
    "כנרת": ("צפון מזרח", 32.7922, 35.5312),
    "יזרעאל": ("צפון מזרח", 32.6078, 35.2897),
    "עפולה": ("צפון מזרח", 32.6078, 35.2897),
    "בית שאן": ("צפון מזרח", 32.4971, 35.4974),
    "נצרת": ("צפון מזרח", 32.6996, 35.3035),
    "עכו": ("צפון מערב", 32.9278, 35.0817),
    "חיפה": ("צפון מערב", 32.7940, 34.9896),
    "חדרה": ("צפון מערב", 32.4340, 34.9197),
    "ירושלים": ("איזור ירושלים", 31.7683, 35.2137),
    "בית לחם": ("איזור ירושלים", 31.6547, 35.1508),
    "ראמאללה": ("איזור ירושלים", 31.8969, 35.0104),
    "השרון": ("אזור השרון", 32.2770, 34.8640),
    "תל אביב": ("איזור גוש דן", 32.0853, 34.7818),
    "רמת גן": ("איזור גוש דן", 32.0823, 34.8106),
    "חולון": ("איזור גוש דן", 32.0158, 34.7874),
    "פתח תקווה": ("איזור גוש דן", 32.0889, 34.8864),
    "רמלה": ("איזור גוש דן", 31.9275, 34.8625),
    "רחובות": ("איזור גוש דן", 31.8928, 34.8113),
    "אשקלון": ("דרום הארץ", 31.6440, 34.5800),
    "באר שבע": ("דרום הארץ", 31.2530, 34.7915),
}


# ==============================================================================
# 4. סיווג גיאוגרפי ופענוח כתובות
# ==============================================================================
def classify_coordinates_to_region(lat: float, lon: float) -> str:
    if 31.62 <= lat <= 31.95 and lon >= 34.93:
        return "איזור ירושלים"
    if lat < 31.83:
        return "דרום הארץ"
    if 31.83 <= lat < 32.13:
        return "איזור גוש דן"
    if 32.13 <= lat < 32.42 and lon < 35.12:
        return "אזור השרון"
    if lat >= 32.35:
        if lon >= 35.22:
            return "צפון מזרח"
        return "צפון מערב"
    return "איזור גוש דן"


def clean_location_text(raw_text: str | None) -> str:
    if not raw_text:
        return ""
    text = str(raw_text).strip()
    text = re.sub(r"ת\.?ד\.?\s*\d+\s*,?", " ", text)
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r'אזה"?ת|אזור תעשייה|איזור תעשיה|תעשיות|מרלו"?ג', " ", text)
    text = re.sub(r"\s+", " ", text).strip(" ,.-")
    return text


def load_cache() -> dict[str, Any]:
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_cache(cache: dict[str, Any]) -> None:
    try:
        CACHE_FILE.write_text(
            json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def query_external_geocoding_coords(
    query: str, timeout: int = 6
) -> tuple[str, float, float] | None:
    search_query = f"{query}, ישראל"
    google_api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()

    if google_api_key:
        try:
            g_url = "https://maps.googleapis.com/maps/api/geocode/json"
            resp = requests.get(
                g_url,
                params={"address": search_query, "key": google_api_key, "language": "iw"},
                timeout=timeout,
            )
            if resp.status_code == 200:
                results = resp.json().get("results", [])
                if results:
                    loc = results[0]["geometry"]["location"]
                    lat, lon = float(loc["lat"]), float(loc["lng"])
                    return classify_coordinates_to_region(lat, lon), lat, lon
        except Exception:
            pass

    try:
        nom_url = "https://nominatim.openstreetmap.org/search"
        headers = {"User-Agent": "DormacoOrderRegionSorter/5.1 (Python/3.12)"}
        resp = requests.get(
            nom_url,
            params={
                "q": search_query,
                "format": "json",
                "limit": 1,
                "countrycodes": "il",
            },
            headers=headers,
            timeout=timeout,
        )
        if resp.status_code == 200:
            data = resp.json()
            if data:
                lat = float(data[0]["lat"])
                lon = float(data[0]["lon"])
                return classify_coordinates_to_region(lat, lon), lat, lon
        time.sleep(1.0)
    except Exception:
        pass

    return None


def query_data_gov_il_api(
    locality_query: str, timeout: int = 5
) -> tuple[str, float, float] | None:
    url = "https://data.gov.il/api/3/action/datastore_search"
    params = {
        "resource_id": "5c78e9fa-c2e2-4771-93ff-7f400a12f7ba",
        "q": locality_query,
        "limit": 5,
    }
    try:
        resp = requests.get(url, params=params, timeout=timeout)
        if resp.status_code == 200:
            records = resp.json().get("result", {}).get("records", [])
            for rec in records:
                nafa = str(rec.get("שם_נפה", "")).strip()
                for key, info in CBS_SUBDISTRICT_TO_REGION.items():
                    if key in nafa:
                        return info
    except Exception:
        pass
    return None


def resolve_region_and_coords(
    destination: str | None,
    address: str | None,
    site_name: str | None = None,
    phone: str | None = None,
    cache: dict[str, Any] | None = None,
    use_external_api: bool = True,
    prefer_site_name_city: bool = False,
) -> tuple[str, float, float]:
    if cache is None:
        cache = {}

    dest_str = str(destination or "").strip()
    addr_str = str(address or "").strip()
    site_str = str(site_name or "").strip()

    for city_kw, street_kw, s_lat, s_lon in SPECIFIC_ADDRESS_COORDS:
        if (city_kw in dest_str or city_kw in addr_str) and street_kw in addr_str:
            reg = KNOWN_LOCALITIES[city_kw][0]
            return reg, s_lat, s_lon

    candidates: list[str] = []
    if prefer_site_name_city and site_str:
        candidates.append(site_str)
    if dest_str:
        candidates.append(dest_str)
    if addr_str:
        candidates.append(addr_str)
    if not prefer_site_name_city and site_str:
        candidates.append(site_str)

    for cand in candidates:
        if not cand:
            continue
        if cand in KNOWN_LOCALITIES:
            return KNOWN_LOCALITIES[cand]
        cleaned = clean_location_text(cand)
        if cleaned in KNOWN_LOCALITIES:
            return KNOWN_LOCALITIES[cleaned]
        if cand in cache and isinstance(cache[cand], list) and len(cache[cand]) == 3:
            reg, lat, lon = cache[cand]
            return str(reg), float(lat), float(lon)
        for loc_name, info in sorted(
            KNOWN_LOCALITIES.items(), key=lambda x: len(x[0]), reverse=True
        ):
            if loc_name in cand or loc_name in cleaned:
                return info

    if use_external_api:
        if dest_str and addr_str:
            combined = f"{clean_location_text(addr_str)}, {clean_location_text(dest_str)}"
            res = query_external_geocoding_coords(combined)
            if res:
                cache[combined] = list(res)
                save_cache(cache)
                return res

        for cand in candidates:
            cleaned = clean_location_text(cand)
            if not cleaned:
                continue
            res = query_external_geocoding_coords(cleaned)
            if res:
                cache[cand] = list(res)
                save_cache(cache)
                return res
            res_gov = query_data_gov_il_api(cleaned)
            if res_gov:
                cache[cand] = list(res_gov)
                save_cache(cache)
                return res_gov

    phone_str = str(phone or "").strip()
    if phone_str.startswith("02"):
        return ("איזור ירושלים", 31.7683, 35.2137)
    if phone_str.startswith("09"):
        return ("אזור השרון", 32.2770, 34.8640)
    if phone_str.startswith("03"):
        return ("איזור גוש דן", 32.0158, 34.7874)
    if phone_str.startswith("08"):
        return ("דרום הארץ", 31.6440, 34.5800)
    if phone_str.startswith("04"):
        return ("צפון מערב", 32.7940, 34.9896)

    return ("איזור גוש דן", 32.0158, 34.7874)


# ==============================================================================
# 5. חישוב זמני נסיעה בכבישי ישראל + העמסה רב-מחסנית (רגיל / קירור / חומ"ס)
# ==============================================================================
NODE_HOME = 0
NODE_MAIN_WH = 1
NODE_COOLING_WH = 2
NODE_HAZMAT_WH = 3


def haversine_air_km(c1: tuple[float, float], c2: tuple[float, float]) -> float:
    lat1, lon1 = math.radians(c1[0]), math.radians(c1[1])
    lat2, lon2 = math.radians(c2[0]), math.radians(c2[1])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    )
    return 6371.0 * (2 * math.asin(math.sqrt(a)))


def estimate_israel_road_drive_minutes(
    c1: tuple[float, float], c2: tuple[float, float]
) -> float:
    """
    שכבה 3 (גיבוי Offline): חישוב זמן נסיעה ריאלי למשאית בכבישי ישראל (בדקות).
    """
    air_km = haversine_air_km(c1, c2)
    if air_km < 0.05:
        return 1.0
    if air_km < 0.6:
        return 3.0

    avg_lon = (c1[1] + c2[1]) / 2.0
    avg_lat = (c1[0] + c2[0]) / 2.0
    if avg_lon > 35.05 or (31.65 <= avg_lat <= 31.92 and avg_lon > 34.95):
        circuity = 1.35
    else:
        circuity = 1.24

    road_km = air_km * circuity

    if road_km < 8.0:
        speed_kmh = 32.0
    elif road_km < 25.0:
        speed_kmh = 50.0
    elif road_km < 65.0:
        speed_kmh = 64.0
    else:
        speed_kmh = 70.0

    return round((road_km / speed_kmh) * 60.0 + 4.0, 1)


def fetch_google_distance_matrix(
    points: list[tuple[float, float]],
    api_key: str,
    truck_speed_factor: float,
    fallback_matrix: list[list[float]],
) -> list[list[float]] | None:
    """
    שכבה 1 (אם הוגדר GOOGLE_MAPS_API_KEY):
    חישוב זמני נסיעה בכבישי ישראל באמצעות Google Maps Distance Matrix API.
    """
    n = len(points)
    if n > 25:
        return None
    try:
        locs = "|".join(f"{lat:.5f},{lon:.5f}" for lat, lon in points)
        url = "https://maps.googleapis.com/maps/api/distancematrix/json"
        Matrix = [[fallback_matrix[i][j] for j in range(n)] for i in range(n)]
        # Google Distance Matrix מאפשר עד 100 אלמנטים בבקשה אחת, לכן נחלק לשורות
        batch_size = max(1, 100 // n)
        for start_i in range(0, n, batch_size):
            end_i = min(n, start_i + batch_size)
            origins_str = "|".join(
                f"{points[i][0]:.5f},{points[i][1]:.5f}" for i in range(start_i, end_i)
            )
            resp = requests.get(
                url,
                params={
                    "origins": origins_str,
                    "destinations": locs,
                    "mode": "driving",
                    "language": "iw",
                    "region": "il",
                    "key": api_key,
                },
                timeout=6,
            )
            if resp.status_code != 200:
                return None
            data = resp.json()
            if data.get("status") != "OK":
                return None
            rows = data.get("rows", [])
            for r_offset, row_obj in enumerate(rows):
                i = start_i + r_offset
                elements = row_obj.get("elements", [])
                for j, el in enumerate(elements):
                    if i == j:
                        Matrix[i][j] = 0.0
                    elif el.get("status") == "OK" and "duration" in el:
                        sec = float(el["duration"]["value"])
                        g_min = (sec / 60.0) * truck_speed_factor
                        Matrix[i][j] = round(max(g_min, fallback_matrix[i][j] * 0.8), 1)
        return Matrix
    except Exception:
        return None


def fetch_real_road_minutes_matrix(
    points: list[tuple[float, float]],
    use_external_api: bool = True,
    truck_speed_factor: float = OSRM_TRUCK_SPEED_FACTOR,
) -> list[list[float]]:
    """
    מחשב מטריצת זמני נסיעה בדקות בין כל הנקודות לפי 3 שכבות:
    1. Google Maps Distance Matrix API (אם הוזן GOOGLE_MAPS_API_KEY).
    2. OSRM Real Israeli Road Network API (מבוסס כבישי ישראל ב-OpenStreetMap).
    3. מודל כבישי ישראל מכויל למשאיות (גיבוי Offline).
    """
    n = len(points)
    fallback_matrix = [
        [
            0.0 if i == j else estimate_israel_road_drive_minutes(points[i], points[j])
            for j in range(n)
        ]
        for i in range(n)
    ]

    if not use_external_api or n <= 1 or n > 65:
        return fallback_matrix

    # שכבה 1: Google Maps Distance Matrix (אם הוזן מפתח API)
    google_api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if google_api_key:
        g_matrix = fetch_google_distance_matrix(
            points, google_api_key, truck_speed_factor, fallback_matrix
        )
        if g_matrix is not None:
            return g_matrix

    # שכבה 2: OSRM Real Road Network API
    coords_str = ";".join(f"{lon:.5f},{lat:.5f}" for lat, lon in points)
    url = f"https://router.project-osrm.org/table/v1/driving/{coords_str}?annotations=duration"
    try:
        resp = requests.get(url, timeout=6)
        if resp.status_code == 200:
            data = resp.json()
            durations = data.get("durations")
            if durations and len(durations) == n:
                matrix: list[list[float]] = []
                for i in range(n):
                    row_vals: list[float] = []
                    for j in range(n):
                        if i == j:
                            row_vals.append(0.0)
                        elif durations[i][j] is not None and durations[i][j] > 0:
                            osrm_min = (float(durations[i][j]) / 60.0) * truck_speed_factor
                            row_vals.append(
                                round(max(osrm_min, fallback_matrix[i][j] * 0.8), 1)
                            )
                        else:
                            row_vals.append(fallback_matrix[i][j])
                    matrix.append(row_vals)
                return matrix
    except Exception:
        pass

    # שכבה 3: גיבוי מודל כבישי ישראל
    return fallback_matrix


def format_hhmm(minutes_from_midnight: float) -> str:
    total_min = int(round(minutes_from_midnight))
    hh = (total_min // 60) % 24
    mm = total_min % 60
    return f"{hh:02d}:{mm:02d}"


def format_duration_hhmm(duration_minutes: float) -> str:
    total_min = max(0, int(round(duration_minutes)))
    hh = total_min // 60
    mm = total_min % 60
    return f"{hh:02d}:{mm:02d}"


def optimize_open_tsp_order(
    stop_indices: list[int],
    start_node: int,
    end_node: int,
    drive_matrix: list[list[float]],
) -> list[int]:
    k = len(stop_indices)
    if k <= 1:
        return list(stop_indices)

    if k <= 8:
        best_perm = tuple(stop_indices)
        best_cost = float("inf")
        for perm in itertools.permutations(stop_indices):
            cost = drive_matrix[start_node][perm[0]]
            for idx in range(k - 1):
                cost += drive_matrix[perm[idx]][perm[idx + 1]]
                if cost >= best_cost:
                    break
            cost += drive_matrix[perm[-1]][end_node]
            if cost < best_cost:
                best_cost = cost
                best_perm = perm
        return list(best_perm)

    unvisited = set(stop_indices)
    current = start_node
    tour: list[int] = []
    while unvisited:
        nxt = min(unvisited, key=lambda node: drive_matrix[current][node])
        tour.append(nxt)
        unvisited.remove(nxt)
        current = nxt

    improved = True
    while improved:
        improved = False
        for i in range(k - 1):
            for j in range(i + 1, k):
                prev_node = start_node if i == 0 else tour[i - 1]
                next_node = end_node if j + 1 == k else tour[j + 1]
                old_edge = drive_matrix[prev_node][tour[i]] + drive_matrix[tour[j]][next_node]
                new_edge = drive_matrix[prev_node][tour[j]] + drive_matrix[tour[i]][next_node]
                if new_edge + 1e-6 < old_edge:
                    tour[i : j + 1] = reversed(tour[i : j + 1])
                    improved = True

    return tour


def evaluate_truck_route(
    ordered_stops: list[dict[str, Any]],
    drive_matrix: list[list[float]],
    main_load_minutes: float = MAIN_WAREHOUSE_LOAD_MINUTES,
    cooling_load_minutes: float = COOLING_WAREHOUSE_LOAD_MINUTES,
    hazmat_load_minutes: float = HAZMAT_WAREHOUSE_LOAD_MINUTES,
    unload_minutes: float = UNLOAD_MINUTES,
    start_minutes: float = DRIVER_START_MINUTES,
    always_load_main_wh: bool = ALWAYS_LOAD_AT_MAIN_WAREHOUSE,
) -> dict[str, Any]:
    """
    מחשב לוח זמנים מלא למשאית עם זמני העמסה נפרדים לכל מחסן:
    1. יציאה ממרכז נתניה (NODE_HOME = 0) בשעה 05:30.
    2. מחסן ראשי "מבוא הספנים 1 אשדוד" (NODE_MAIN_WH = 1) -> העמסה במשך main_load_minutes (60 דק' כברירת מחדל).
    3. אם יש מוצר קירור: מחסן קירור "מבוא הספנים 6 אשדוד" (NODE_COOLING_WH = 2) -> העמסה במשך cooling_load_minutes (30 דק').
    4. אם יש מוצר חומ"ס: מחסן חומ"ס "פארק נמלי ישראל, אשדוד" (NODE_HAZMAT_WH = 3) -> העמסה במשך hazmat_load_minutes (30 דק').
    5. נסיעה מהמחסן האחרון לכל תחנות הלקוחות במסלול (כולל unload_minutes = 30 דק' פריקה בכל עצירה).
    6. נסיעה מהלקוח האחרון חזרה למרכז נתניה (NODE_HOME = 0).
    """
    has_regular = False
    has_cooling = False
    has_hazmat = False

    for st in ordered_stops:
        for r in st["records"]:
            w_type = r.get("warehouse_type", "רגיל")
            if w_type == "קירור":
                has_cooling = True
            elif w_type == 'חומ"ס':
                has_hazmat = True
            else:
                has_regular = True

    warehouse_visits: list[dict[str, Any]] = []
    current_time = start_minutes
    current_node = NODE_HOME
    total_drive_min = 0.0
    total_load_min = 0.0

    # שלב א': מחסן ראשי - מבוא הספנים 1 אשדוד (60 דקות כברירת מחדל)
    if always_load_main_wh or has_regular or (not has_cooling and not has_hazmat):
        leg = drive_matrix[current_node][NODE_MAIN_WH]
        arr = current_time + leg
        dep = arr + main_load_minutes
        total_drive_min += leg
        total_load_min += main_load_minutes
        warehouse_visits.append(
            {
                "name": f"מחסן ראשי ({MAIN_WAREHOUSE_ADDRESS})",
                "short_name": f"מבוא הספנים 1 ({int(main_load_minutes)} דק')",
                "node": NODE_MAIN_WH,
                "arrival_min": arr,
                "arrival_str": format_hhmm(arr),
                "departure_min": dep,
                "departure_str": format_hhmm(dep),
            }
        )
        current_time = dep
        current_node = NODE_MAIN_WH

    # שלב ב': מחסן קירור - מבוא הספנים 6 אשדוד (30 דקות כברירת מחדל)
    if has_cooling:
        leg = drive_matrix[current_node][NODE_COOLING_WH]
        arr = current_time + leg
        dep = arr + cooling_load_minutes
        total_drive_min += leg
        total_load_min += cooling_load_minutes
        warehouse_visits.append(
            {
                "name": f"מחסן קירור ({COOLING_WAREHOUSE_ADDRESS})",
                "short_name": f"מבוא הספנים 6 קירור ({int(cooling_load_minutes)} דק')",
                "node": NODE_COOLING_WH,
                "arrival_min": arr,
                "arrival_str": format_hhmm(arr),
                "departure_min": dep,
                "departure_str": format_hhmm(dep),
            }
        )
        current_time = dep
        current_node = NODE_COOLING_WH

    # שלב ג': מחסן חומ"ס - פארק נמלי ישראל, אשדוד (30 דקות כברירת מחדל)
    if has_hazmat:
        leg = drive_matrix[current_node][NODE_HAZMAT_WH]
        arr = current_time + leg
        dep = arr + hazmat_load_minutes
        total_drive_min += leg
        total_load_min += hazmat_load_minutes
        warehouse_visits.append(
            {
                "name": f'מחסן חומ"ס ({HAZMAT_WAREHOUSE_ADDRESS})',
                "short_name": f'פארק נמלי ישראל חומ"ס ({int(hazmat_load_minutes)} דק\')',
                "node": NODE_HAZMAT_WH,
                "arrival_min": arr,
                "arrival_str": format_hhmm(arr),
                "departure_min": dep,
                "departure_str": format_hhmm(dep),
            }
        )
        current_time = dep
        current_node = NODE_HAZMAT_WH

    # שלב ד': מעבר בין תחנות החלוקה אצל הלקוחות
    stop_schedules: list[dict[str, Any]] = []
    total_unload_min = 0.0
    total_weight = 0.0

    for stop_idx, stop in enumerate(ordered_stops, start=1):
        mat_idx = stop["matrix_idx"]
        leg_drive = drive_matrix[current_node][mat_idx]
        arrival_min = current_time + leg_drive
        departure_min = arrival_min + unload_minutes

        total_drive_min += leg_drive
        total_unload_min += unload_minutes
        total_weight += stop["weight"]

        stop_schedules.append(
            {
                "stop": stop,
                "route_stop_num": stop_idx,
                "leg_drive_min": leg_drive,
                "arrival_min": arrival_min,
                "arrival_str": format_hhmm(arrival_min),
                "departure_min": departure_min,
                "departure_str": format_hhmm(departure_min),
            }
        )

        current_time = departure_min
        current_node = mat_idx

    # שלב ה': חזרה מהלקוח האחרון לבית הנהג במרכז נתניה
    last_to_home_min = drive_matrix[current_node][NODE_HOME] if ordered_stops else 0.0
    total_drive_min += last_to_home_min
    return_home_min = current_time + last_to_home_min
    total_shift_min = return_home_min - start_minutes

    wh_summary_str = " -> ".join(
        f"{w['short_name']} [{w['arrival_str']}-{w['departure_str']}]"
        for w in warehouse_visits
    )

    return {
        "ordered_stops": ordered_stops,
        "warehouse_visits": warehouse_visits,
        "warehouses_summary_str": wh_summary_str,
        "has_cooling": has_cooling,
        "has_hazmat": has_hazmat,
        "stop_schedules": stop_schedules,
        "return_home_min": return_home_min,
        "return_home_str": format_hhmm(return_home_min),
        "total_drive_min": total_drive_min,
        "total_load_min": total_load_min,
        "total_unload_min": total_unload_min,
        "total_shift_min": total_shift_min,
        "total_shift_hours": total_shift_min / 60.0,
        "total_weight": total_weight,
    }


def build_region_stops(
    reg_records: list[dict[str, Any]],
    max_truck_weight: float = MAX_TRUCK_WEIGHT_KG,
) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    order_of_keys: list[str] = []
    for rec in reg_records:
        s_key = rec["stop_key"]
        if s_key not in grouped:
            grouped[s_key] = []
            order_of_keys.append(s_key)
        grouped[s_key].append(rec)

    stops: list[dict[str, Any]] = []
    for s_key in order_of_keys:
        recs = grouped[s_key]
        total_w = sum(r["qty"] for r in recs)
        if total_w <= max_truck_weight or len(recs) == 1:
            stops.append(
                {
                    "stop_key": s_key,
                    "region": recs[0]["region"],
                    "effective_dest": recs[0]["effective_dest"],
                    "cust_name": recs[0]["cust_name"],
                    "lat": recs[0]["lat"],
                    "lon": recs[0]["lon"],
                    "weight": total_w,
                    "records": recs,
                }
            )
        else:
            current_batch: list[dict[str, Any]] = []
            current_w = 0.0
            part_num = 1
            for r in recs:
                if current_batch and (current_w + r["qty"] > max_truck_weight):
                    stops.append(
                        {
                            "stop_key": f"{s_key}#part{part_num}",
                            "region": current_batch[0]["region"],
                            "effective_dest": current_batch[0]["effective_dest"],
                            "cust_name": current_batch[0]["cust_name"],
                            "lat": current_batch[0]["lat"],
                            "lon": current_batch[0]["lon"],
                            "weight": current_w,
                            "records": current_batch,
                        }
                    )
                    part_num += 1
                    current_batch = [r]
                    current_w = r["qty"]
                else:
                    current_batch.append(r)
                    current_w += r["qty"]
            if current_batch:
                stops.append(
                    {
                        "stop_key": f"{s_key}#part{part_num}",
                        "region": current_batch[0]["region"],
                        "effective_dest": current_batch[0]["effective_dest"],
                        "cust_name": current_batch[0]["cust_name"],
                        "lat": current_batch[0]["lat"],
                        "lon": current_batch[0]["lon"],
                        "weight": current_w,
                        "records": current_batch,
                    }
                )
    return stops


def get_last_warehouse_node_for_stops(stops_slice: list[dict[str, Any]]) -> int:
    for st in stops_slice:
        for r in st["records"]:
            if r.get("warehouse_type") == 'חומ"ס':
                return NODE_HAZMAT_WH
    for st in stops_slice:
        for r in st["records"]:
            if r.get("warehouse_type") == "קירור":
                return NODE_COOLING_WH
    return NODE_MAIN_WH


def split_region_into_trucks(
    region_name: str,
    region_stops: list[dict[str, Any]],
    drive_matrix: list[list[float]],
    max_truck_weight: float = MAX_TRUCK_WEIGHT_KG,
    max_shift_hours: float = MAX_DRIVER_SHIFT_HOURS,
    main_load_minutes: float = MAIN_WAREHOUSE_LOAD_MINUTES,
    cooling_load_minutes: float = COOLING_WAREHOUSE_LOAD_MINUTES,
    hazmat_load_minutes: float = HAZMAT_WAREHOUSE_LOAD_MINUTES,
    unload_minutes: float = UNLOAD_MINUTES,
) -> list[dict[str, Any]]:
    if not region_stops:
        return []

    max_shift_min = max_shift_hours * 60.0
    start_wh_node = get_last_warehouse_node_for_stops(region_stops)

    stop_idxs = [s["matrix_idx"] for s in region_stops]
    idx_to_stop_list: dict[int, list[dict[str, Any]]] = {}
    for s in region_stops:
        idx_to_stop_list.setdefault(s["matrix_idx"], []).append(s)

    unique_mat_idxs = list(dict.fromkeys(stop_idxs))
    ordered_mat_idxs = optimize_open_tsp_order(
        unique_mat_idxs,
        start_node=start_wh_node,
        end_node=NODE_HOME,
        drive_matrix=drive_matrix,
    )

    master_stops: list[dict[str, Any]] = []
    for m_idx in ordered_mat_idxs:
        master_stops.extend(idx_to_stop_list[m_idx])

    k = len(master_stops)

    single_eval = evaluate_truck_route(
        master_stops,
        drive_matrix=drive_matrix,
        main_load_minutes=main_load_minutes,
        cooling_load_minutes=cooling_load_minutes,
        hazmat_load_minutes=hazmat_load_minutes,
        unload_minutes=unload_minutes,
    )
    if (
        single_eval["total_weight"] <= max_truck_weight + 1e-6
        and single_eval["total_shift_min"] <= max_shift_min + 1e-6
    ):
        single_eval["regions"] = [region_name]
        single_eval["region_label"] = region_name
        single_eval["was_split"] = False
        return [single_eval]

    dp: list[tuple[int, float, int, dict[str, Any] | None]] = [
        (999999, float("inf"), -1, None) for _ in range(k + 1)
    ]
    dp[0] = (0, 0.0, -1, None)

    for i in range(k):
        if dp[i][0] >= 999999:
            continue
        for j in range(i + 1, k + 1):
            sub_slice = master_stops[i:j]
            sub_w = sum(s["weight"] for s in sub_slice)
            if j > i + 1 and sub_w > max_truck_weight + 1e-6:
                break

            sub_wh_node = get_last_warehouse_node_for_stops(sub_slice)
            sub_mat_idxs = list(dict.fromkeys(s["matrix_idx"] for s in sub_slice))
            opt_sub_mat = optimize_open_tsp_order(
                sub_mat_idxs,
                start_node=sub_wh_node,
                end_node=NODE_HOME,
                drive_matrix=drive_matrix,
            )
            sub_map: dict[int, list[dict[str, Any]]] = {}
            for s in sub_slice:
                sub_map.setdefault(s["matrix_idx"], []).append(s)
            opt_sub_stops: list[dict[str, Any]] = []
            for m_idx in opt_sub_mat:
                opt_sub_stops.extend(sub_map[m_idx])

            ev = evaluate_truck_route(
                opt_sub_stops,
                drive_matrix=drive_matrix,
                main_load_minutes=main_load_minutes,
                cooling_load_minutes=cooling_load_minutes,
                hazmat_load_minutes=hazmat_load_minutes,
                unload_minutes=unload_minutes,
            )
            if j > i + 1 and ev["total_shift_min"] > max_shift_min + 1e-6:
                continue

            cand_trucks = dp[i][0] + 1
            cand_shift = dp[i][1] + ev["total_shift_min"]
            if (cand_trucks < dp[j][0]) or (
                cand_trucks == dp[j][0] and cand_shift + 1e-6 < dp[j][1]
            ):
                dp[j] = (cand_trucks, cand_shift, i, ev)

    trucks: list[dict[str, Any]] = []
    curr = k
    while curr > 0:
        _, _, prev_i, ev = dp[curr]
        if ev is None or prev_i < 0:
            break
        trucks.append(ev)
        curr = prev_i
    trucks.reverse()

    total_parts = len(trucks)
    for idx, tr in enumerate(trucks, start=1):
        tr["regions"] = [region_name]
        if total_parts > 1:
            tr["region_label"] = f"{region_name} (נהג {idx} מתוך {total_parts})"
            tr["was_split"] = True
        else:
            tr["region_label"] = region_name
            tr["was_split"] = False

    return trucks


def try_consolidate_north_and_sharon_trucks(
    trucks_by_region: dict[str, list[dict[str, Any]]],
    drive_matrix: list[list[float]],
    max_truck_weight: float = MAX_TRUCK_WEIGHT_KG,
    max_shift_hours: float = MAX_DRIVER_SHIFT_HOURS,
    main_load_minutes: float = MAIN_WAREHOUSE_LOAD_MINUTES,
    cooling_load_minutes: float = COOLING_WAREHOUSE_LOAD_MINUTES,
    hazmat_load_minutes: float = HAZMAT_WAREHOUSE_LOAD_MINUTES,
    unload_minutes: float = UNLOAD_MINUTES,
    enable_merge: bool = ENABLE_NORTH_SHARON_MERGE,
) -> list[dict[str, Any]]:
    max_shift_min = max_shift_hours * 60.0
    sharon_trucks = trucks_by_region.get("אזור השרון", [])

    if enable_merge and sharon_trucks:
        best_merge: tuple[str, int, int, dict[str, Any], float] | None = None

        for north_reg in ("צפון מזרח", "צפון מערב"):
            n_trucks = trucks_by_region.get(north_reg, [])
            for n_idx, n_tr in enumerate(n_trucks):
                for s_idx, s_tr in enumerate(sharon_trucks):
                    combined_weight = n_tr["total_weight"] + s_tr["total_weight"]
                    if combined_weight > max_truck_weight + 1e-6:
                        continue

                    north_stops = list(n_tr["ordered_stops"])
                    sharon_stops = list(s_tr["ordered_stops"])

                    last_north_mat_idx = north_stops[-1]["matrix_idx"]
                    sharon_mat_idxs = list(
                        dict.fromkeys(s["matrix_idx"] for s in sharon_stops)
                    )
                    opt_sharon_mat = optimize_open_tsp_order(
                        sharon_mat_idxs,
                        start_node=last_north_mat_idx,
                        end_node=NODE_HOME,
                        drive_matrix=drive_matrix,
                    )
                    s_map: dict[int, list[dict[str, Any]]] = {}
                    for s in sharon_stops:
                        s_map.setdefault(s["matrix_idx"], []).append(s)
                    opt_sharon_stops: list[dict[str, Any]] = []
                    for m_idx in opt_sharon_mat:
                        opt_sharon_stops.extend(s_map[m_idx])

                    combined_stops = north_stops + opt_sharon_stops
                    merged_eval = evaluate_truck_route(
                        combined_stops,
                        drive_matrix=drive_matrix,
                        main_load_minutes=main_load_minutes,
                        cooling_load_minutes=cooling_load_minutes,
                        hazmat_load_minutes=hazmat_load_minutes,
                        unload_minutes=unload_minutes,
                    )

                    if merged_eval["total_shift_min"] <= max_shift_min + 1e-6:
                        saved_min = (
                            n_tr["total_shift_min"] + s_tr["total_shift_min"]
                        ) - merged_eval["total_shift_min"]
                        if best_merge is None or saved_min > best_merge[4]:
                            merged_eval["regions"] = [north_reg, "אזור השרון"]
                            merged_eval["region_label"] = f"{north_reg} + אזור השרון"
                            merged_eval["was_split"] = False
                            merged_eval["is_merged_north_sharon"] = True
                            best_merge = (
                                north_reg,
                                n_idx,
                                s_idx,
                                merged_eval,
                                saved_min,
                            )

        if best_merge is not None:
            north_reg, n_idx, s_idx, merged_eval, _ = best_merge
            trucks_by_region[north_reg][n_idx] = merged_eval
            del trucks_by_region["אזור השרון"][s_idx]

    final_trucks: list[dict[str, Any]] = []
    for reg in REGION_ORDER:
        for tr in trucks_by_region.get(reg, []):
            final_trucks.append(tr)

    for d_idx, tr in enumerate(final_trucks, start=1):
        tr["driver_num"] = d_idx
        tr["driver_name"] = f"נהג {d_idx} ({tr['region_label']})"

    return final_trucks


# ==============================================================================
# 6. מיון, תכנון זמני העמסה/הגעה וייצוא לקובץ אקסל מעוצב
# ==============================================================================
def sort_orders_excel(
    input_path: str | pathlib.Path,
    output_path: str | pathlib.Path,
    use_external_api: bool = True,
    prefer_site_name_city: bool = False,
    max_truck_weight: float = MAX_TRUCK_WEIGHT_KG,
    max_shift_hours: float = MAX_DRIVER_SHIFT_HOURS,
    main_load_minutes: float = MAIN_WAREHOUSE_LOAD_MINUTES,
    cooling_load_minutes: float = COOLING_WAREHOUSE_LOAD_MINUTES,
    hazmat_load_minutes: float = HAZMAT_WAREHOUSE_LOAD_MINUTES,
    unload_minutes: float = UNLOAD_MINUTES,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    cooling_skus, hazmat_skus = load_cooling_and_hazmat_catalogs(INPUT_FOLDER)

    wb_in = load_workbook_safe(input_path)
    ws_in = wb_in.active

    raw_headers = [
        str(ws_in.cell(1, col).value or "").strip()
        for col in range(1, ws_in.max_column + 1)
    ]

    def find_col(possible_names: list[str], default_idx: int) -> int:
        for name in possible_names:
            for idx, h in enumerate(raw_headers):
                if name in h:
                    return idx
        return default_idx

    idx_cust_id = find_col(["מס. לקוח", "מספר לקוח"], 0)
    idx_cust_name = find_col(["שם לקוח"], 1)
    idx_address = find_col(["כתובת"], 2)
    idx_dest = find_col(["יעד", "עיר"], 3)
    idx_site = find_col(["שם האתר", "אתר"], 4)
    idx_phone = find_col(["מספר טלפון", "טלפון"], 5)
    idx_order = find_col(["הזמנה"], 6)
    idx_date = find_col(["תאריך אספקה", "תאריך"], 7)
    idx_line = find_col(["שורה"], 8)
    idx_sku = find_col(["מק'ט", 'מק"ט', "מקט"], 9)
    idx_qty = find_col(["כמות"], 11)
    idx_note_cust = find_col(["הערה למוצר ללקוח", "הערה ללקוח"], 12)
    idx_units = find_col(["יחידות", "יח'", "יחידה"], 13)

    cache = load_cache()
    records: list[dict[str, Any]] = []

    for r in range(2, ws_in.max_row + 1):
        row_vals = [ws_in.cell(r, c).value for c in range(1, ws_in.max_column + 1)]
        if not any(v is not None and str(v).strip() != "" for v in row_vals):
            continue

        dest_val = row_vals[idx_dest] if idx_dest < len(row_vals) else None
        addr_val = row_vals[idx_address] if idx_address < len(row_vals) else None
        site_val = row_vals[idx_site] if idx_site < len(row_vals) else None
        phone_val = row_vals[idx_phone] if idx_phone < len(row_vals) else None

        region, lat, lon = resolve_region_and_coords(
            destination=dest_val,
            address=addr_val,
            site_name=site_val,
            phone=phone_val,
            cache=cache,
            use_external_api=use_external_api,
            prefer_site_name_city=prefer_site_name_city,
        )

        effective_dest = str(dest_val or "").strip()
        if not effective_dest and addr_val:
            effective_dest = str(addr_val).strip()

        qty_val = 0.0
        if idx_qty < len(row_vals) and row_vals[idx_qty] is not None:
            try:
                qty_val = float(row_vals[idx_qty])
            except (ValueError, TypeError):
                qty_val = 0.0

        sku_raw = row_vals[idx_sku] if idx_sku < len(row_vals) else ""
        sku_clean = clean_sku_token(sku_raw)

        if sku_clean in cooling_skus:
            warehouse_type = "קירור"
            warehouse_label = f"קירור ({COOLING_WAREHOUSE_ADDRESS})"
        elif sku_clean in hazmat_skus:
            warehouse_type = 'חומ"ס'
            warehouse_label = f'חומ"ס ({HAZMAT_WAREHOUSE_ADDRESS})'
        else:
            warehouse_type = "רגיל"
            warehouse_label = f"רגיל ({MAIN_WAREHOUSE_ADDRESS})"

        cust_name = str(row_vals[idx_cust_name] or "").strip()
        stop_key = f"{effective_dest}||{cust_name}||{lat:.5f},{lon:.5f}"

        records.append(
            {
                "region": region,
                "region_rank": REGION_RANK.get(region, 99),
                "effective_dest": effective_dest,
                "cust_name": cust_name,
                "order_no": str(row_vals[idx_order] or "").strip(),
                "line_no": int(row_vals[idx_line] or 0)
                if row_vals[idx_line] is not None
                else 0,
                "sku": sku_clean,
                "warehouse_type": warehouse_type,
                "warehouse_label": warehouse_label,
                "qty": qty_val,
                "lat": lat,
                "lon": lon,
                "stop_key": stop_key,
                "raw_values": row_vals,
            }
        )

    all_coords: list[tuple[float, float]] = [
        DRIVER_HOME_COORDS,
        MAIN_WAREHOUSE_COORDS,
        COOLING_WAREHOUSE_COORDS,
        HAZMAT_WAREHOUSE_COORDS,
    ]
    coord_to_mat_idx: dict[tuple[float, float], int] = {
        DRIVER_HOME_COORDS: NODE_HOME,
        MAIN_WAREHOUSE_COORDS: NODE_MAIN_WH,
        COOLING_WAREHOUSE_COORDS: NODE_COOLING_WH,
        HAZMAT_WAREHOUSE_COORDS: NODE_HAZMAT_WH,
    }

    stops_by_region: dict[str, list[dict[str, Any]]] = {}
    for reg in REGION_ORDER:
        reg_recs = [rec for rec in records if rec["region"] == reg]
        reg_stops = build_region_stops(reg_recs, max_truck_weight=max_truck_weight)
        for st in reg_stops:
            pt = (round(st["lat"], 5), round(st["lon"], 5))
            if pt not in coord_to_mat_idx:
                coord_to_mat_idx[pt] = len(all_coords)
                all_coords.append(pt)
            st["matrix_idx"] = coord_to_mat_idx[pt]
        stops_by_region[reg] = reg_stops

    drive_matrix = fetch_real_road_minutes_matrix(
        all_coords,
        use_external_api=use_external_api,
        truck_speed_factor=OSRM_TRUCK_SPEED_FACTOR,
    )

    trucks_by_region: dict[str, list[dict[str, Any]]] = {}
    for reg in REGION_ORDER:
        trucks_by_region[reg] = split_region_into_trucks(
            region_name=reg,
            region_stops=stops_by_region[reg],
            drive_matrix=drive_matrix,
            max_truck_weight=max_truck_weight,
            max_shift_hours=max_shift_hours,
            main_load_minutes=main_load_minutes,
            cooling_load_minutes=cooling_load_minutes,
            hazmat_load_minutes=hazmat_load_minutes,
            unload_minutes=unload_minutes,
        )

    final_trucks = try_consolidate_north_and_sharon_trucks(
        trucks_by_region=trucks_by_region,
        drive_matrix=drive_matrix,
        max_truck_weight=max_truck_weight,
        max_shift_hours=max_shift_hours,
        main_load_minutes=main_load_minutes,
        cooling_load_minutes=cooling_load_minutes,
        hazmat_load_minutes=hazmat_load_minutes,
        unload_minutes=unload_minutes,
        enable_merge=ENABLE_NORTH_SHARON_MERGE,
    )

    sorted_records: list[dict[str, Any]] = []
    for tr in final_trucks:
        for sched in tr["stop_schedules"]:
            st = sched["stop"]
            st_recs = list(st["records"])
            st_recs.sort(key=lambda item: (item["order_no"], item["line_no"]))
            for rec in st_recs:
                rec["driver_num"] = tr["driver_num"]
                rec["driver_name"] = tr["driver_name"]
                rec["route_stop_num"] = sched["route_stop_num"]
                rec["arrival_str"] = sched["arrival_str"]
                rec["departure_str"] = sched["departure_str"]
                rec["warehouses_summary_str"] = tr["warehouses_summary_str"]
                rec["return_home_str"] = tr["return_home_str"]
                sorted_records.append(rec)

    # ==========================================================================
    # סידור העמודות בטבלה הסופית:
    # - ליד "יעד": "שעת הגעה ליעד" ו"שעת יציאה מיעד"
    # - "יחידות" צמודה ל"כמות", ורק אחריהן "הערה למוצר ללקוח"
    # ==========================================================================
    reordered_orig_cols: list[int] = []
    for c_i in range(len(raw_headers)):
        if c_i in (idx_qty, idx_units, idx_note_cust):
            continue
        reordered_orig_cols.append(c_i)

    insert_pos = min(idx_qty, len(reordered_orig_cols))
    reordered_orig_cols[insert_pos:insert_pos] = [idx_qty, idx_units, idx_note_cust]

    prefix_headers = [
        "אזור חלוקה",
        "נהג / משאית",
        "תחנה במסלול",
        "מחסן העמסה למוצר",
    ]

    out_headers: list[str] = list(prefix_headers)
    col_arrival_idx = -1
    col_depart_idx = -1
    orig_to_out_col: dict[int, int] = {}

    for orig_idx in reordered_orig_cols:
        out_headers.append(raw_headers[orig_idx])
        orig_to_out_col[orig_idx] = len(out_headers)
        if orig_idx == idx_dest:
            out_headers.append("שעת הגעה ליעד")
            col_arrival_idx = len(out_headers)
            out_headers.append(f"שעת יציאה (+{int(unload_minutes)} דק' פריקה)")
            col_depart_idx = len(out_headers)

    total_cols = len(out_headers)
    out_cust_id_col = orig_to_out_col.get(idx_cust_id, -1)
    out_date_col = orig_to_out_col.get(idx_date, -1)
    out_line_col = orig_to_out_col.get(idx_line, -1)
    out_qty_col = orig_to_out_col.get(idx_qty, -1)

    numeric_col_indices = {3, out_cust_id_col, out_line_col, out_qty_col}
    center_col_indices = {col_arrival_idx, col_depart_idx}

    wb_out = openpyxl.Workbook()
    ws_orders = wb_out.active
    ws_orders.title = "הזמנות ממוינות לפי אזור"
    ws_orders.sheet_view.rightToLeft = True
    ws_orders.views.sheetView[0].showGridLines = True
    ws_orders.page_setup.orientation = "landscape"
    ws_orders.sheet_properties.pageSetUpPr.fitToPage = True

    NAVY_HEADER = "1F4E78"
    ZEBRA_FILL = "F9FAFC"
    TOTAL_FILL = "EAEEF7"
    BORDER_GRAY = "D9D9D9"

    ARRIVAL_BG = "E8F5E9"
    ARRIVAL_FONT = "1B5E20"
    COOLING_BG = "E1F5FE"
    COOLING_FONT = "01579B"
    HAZMAT_BG = "FFF3E0"
    HAZMAT_FONT = "E65100"

    REGION_COLORS: dict[str, str] = {
        "צפון מזרח": "E2EFDA",
        "צפון מערב": "D9E1F2",
        "איזור ירושלים": "FFF2CC",
        "אזור השרון": "EDEDED",
        "איזור גוש דן": "FCE4D6",
        "דרום הארץ": "F8CBAD",
    }

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

    header_fill = PatternFill(
        start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid"
    )
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")

    for col_idx, h_text in enumerate(out_headers, start=1):
        cell = ws_orders.cell(row=1, column=col_idx, value=h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        if col_idx in numeric_col_indices:
            horiz = "right"
        elif col_idx in center_col_indices:
            horiz = "center"
        else:
            horiz = "left"
        cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)

    ws_orders.row_dimensions[1].height = 28
    ws_orders.freeze_panes = "A2"

    arrival_fill = PatternFill(
        start_color=ARRIVAL_BG, end_color=ARRIVAL_BG, fill_type="solid"
    )
    arrival_font = Font(name="Arial", size=10, bold=True, color=ARRIVAL_FONT)
    cooling_fill = PatternFill(
        start_color=COOLING_BG, end_color=COOLING_BG, fill_type="solid"
    )
    cooling_font = Font(name="Arial", size=10, bold=True, color=COOLING_FONT)
    hazmat_fill = PatternFill(
        start_color=HAZMAT_BG, end_color=HAZMAT_BG, fill_type="solid"
    )
    hazmat_font = Font(name="Arial", size=10, bold=True, color=HAZMAT_FONT)

    start_data_row = 2
    for offset, rec in enumerate(sorted_records):
        r_idx = start_data_row + offset
        raw = list(rec["raw_values"])
        if not raw[idx_dest] and rec["effective_dest"]:
            raw[idx_dest] = rec["effective_dest"]

        reg_name = rec["region"]
        row_out: list[Any] = [
            reg_name,
            rec["driver_name"],
            rec["route_stop_num"],
            rec["warehouse_label"],
        ]

        for orig_idx in reordered_orig_cols:
            val = raw[orig_idx] if orig_idx < len(raw) else None
            row_out.append(val)
            if orig_idx == idx_dest:
                row_out.append(rec["arrival_str"])
                row_out.append(rec["departure_str"])

        reg_color = REGION_COLORS.get(reg_name, "FFFFFF")
        reg_fill = PatternFill(
            start_color=reg_color, end_color=reg_color, fill_type="solid"
        )
        zebra_fill = PatternFill(
            start_color=ZEBRA_FILL if r_idx % 2 == 0 else "FFFFFF",
            end_color=ZEBRA_FILL if r_idx % 2 == 0 else "FFFFFF",
            fill_type="solid",
        )

        for col_idx, val in enumerate(row_out, start=1):
            if isinstance(val, str):
                val = val.strip()

            if col_idx == out_cust_id_col and val is not None and str(val).isdigit():
                val = int(str(val))
            elif col_idx == out_line_col and val is not None:
                try:
                    val = int(float(val))
                except (ValueError, TypeError):
                    pass
            elif col_idx == out_qty_col and val is not None:
                try:
                    val = float(val)
                except (ValueError, TypeError):
                    pass
            elif col_idx == out_date_col and isinstance(val, datetime.datetime):
                val = val.strftime("%Y-%m-%d")

            cell = ws_orders.cell(row=r_idx, column=col_idx, value=val)
            cell.border = thin_border

            if col_idx == 1:
                cell.fill = reg_fill
                cell.font = Font(name="Arial", size=10, bold=True)
            elif col_idx == 4:
                if rec["warehouse_type"] == "קירור":
                    cell.fill = cooling_fill
                    cell.font = cooling_font
                elif rec["warehouse_type"] == 'חומ"ס':
                    cell.fill = hazmat_fill
                    cell.font = hazmat_font
                else:
                    cell.fill = zebra_fill
                    cell.font = Font(name="Arial", size=10)
            elif col_idx == col_arrival_idx:
                cell.fill = arrival_fill
                cell.font = arrival_font
            else:
                cell.fill = zebra_fill
                cell.font = Font(name="Arial", size=10, bold=(col_idx == 2))

            if col_idx in numeric_col_indices:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if col_idx in (3, out_cust_id_col):
                    cell.number_format = "0"
                else:
                    cell.number_format = "#,##0"
            elif col_idx in center_col_indices:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        ws_orders.row_dimensions[r_idx].height = 20

    end_data_row = start_data_row + len(sorted_records) - 1
    total_row_idx = end_data_row + 1
    qty_col_letter = get_column_letter(out_qty_col)

    ws_orders.cell(row=total_row_idx, column=1, value='סה"כ כללי')
    if end_data_row >= start_data_row:
        ws_orders.cell(
            row=total_row_idx,
            column=out_qty_col,
            value=f"=SUM({qty_col_letter}{start_data_row}:{qty_col_letter}{end_data_row})",
        )
        ws_orders.cell(
            row=total_row_idx,
            column=out_line_col,
            value=f"=COUNT({qty_col_letter}{start_data_row}:{qty_col_letter}{end_data_row})",
        )

    for col_idx in range(1, total_cols + 1):
        c = ws_orders.cell(row=total_row_idx, column=col_idx)
        c.font = Font(name="Arial", size=10, bold=True, color="1F4E78")
        c.fill = PatternFill(
            start_color=TOTAL_FILL, end_color=TOTAL_FILL, fill_type="solid"
        )
        c.border = total_border
        if col_idx in numeric_col_indices:
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = "#,##0"
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
    ws_orders.row_dimensions[total_row_idx].height = 23

    for col_idx in range(1, total_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(out_headers[col_idx - 1]))
        for r_idx in range(start_data_row, end_data_row + 1):
            val = ws_orders.cell(row=r_idx, column=col_idx).value
            if val is not None:
                val_str = str(val)
                if not val_str.startswith("="):
                    max_len = max(max_len, len(val_str))
        ws_orders.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 38)

    # ==========================================================================
    # גיליון 2: סיכום מסלולי נהגים, מחסני העמסה ושעות הגעה
    # ==========================================================================
    ws_trucks = wb_out.create_sheet(title="סיכום מסלולי נהגים ומשאיות")
    ws_trucks.sheet_view.rightToLeft = True
    ws_trucks.views.sheetView[0].showGridLines = True
    ws_trucks.page_setup.orientation = "landscape"
    ws_trucks.sheet_properties.pageSetUpPr.fitToPage = True

    truck_headers = [
        "נהג / משאית",
        "אזורי חלוקה במסלול",
        "יציאה מנתניה",
        "מחסני העמסה באשדוד (כולל זמני הגעה והעמסה)",
        "מספר עצירות",
        'סה"כ משקל במשאית (ק"ג)',
        "חזרה למרכז נתניה",
        "סה\"כ זמן משמרת (מתוך 12:00 שעות)",
        "סטטוס מסלול",
        "פירוט מלא של מסלול הנסיעה ושעות ההגעה",
    ]

    for col_idx, h_text in enumerate(truck_headers, start=1):
        c = ws_trucks.cell(row=1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.border = thin_border
        c.alignment = Alignment(
            horizontal="right" if col_idx in (5, 6) else "left",
            vertical="center",
            wrap_text=True,
        )
    ws_trucks.row_dimensions[1].height = 28
    ws_trucks.freeze_panes = "A2"

    ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    ok_font = Font(name="Arial", size=10, bold=True, color="274E13")

    orders_sheet_ref = "'הזמנות ממוינות לפי אזור'"
    driver_range_ref = f"{orders_sheet_ref}!$B${start_data_row}:$B${end_data_row}"
    qty_range_ref = (
        f"{orders_sheet_ref}!${qty_col_letter}${start_data_row}:"
        f"${qty_col_letter}${end_data_row}"
    )

    t_start_row = 2
    for idx, tr in enumerate(final_trucks):
        r_idx = t_start_row + idx
        stops_summary = " -> ".join(
            f"{s['stop']['effective_dest']} ({s['arrival_str']})"
            for s in tr["stop_schedules"]
        )
        full_route_text = (
            f"{DRIVER_HOME_NAME} ({format_hhmm(DRIVER_START_MINUTES)}) -> "
            f"{tr['warehouses_summary_str']} -> "
            f"{stops_summary} -> {DRIVER_HOME_NAME} ({tr['return_home_str']})"
        )

        status_notes: list[str] = []
        if tr.get("has_cooling"):
            status_notes.append("כולל העמסת קירור")
        if tr.get("has_hazmat"):
            status_notes.append('כולל העמסת חומ"ס')
        if tr.get("is_merged_north_sharon"):
            status_notes.append("אוחד נהג (צפון + השרון)")
        elif tr.get("was_split"):
            status_notes.append("פוצל אוטומטית לעמידה במגבלות")
        else:
            status_notes.append("תקין")

        ws_trucks.cell(row=r_idx, column=1, value=tr["driver_name"])
        ws_trucks.cell(row=r_idx, column=2, value=tr["region_label"])
        ws_trucks.cell(row=r_idx, column=3, value=format_hhmm(DRIVER_START_MINUTES))
        ws_trucks.cell(row=r_idx, column=4, value=tr["warehouses_summary_str"])
        ws_trucks.cell(row=r_idx, column=5, value=len(tr["stop_schedules"]))
        ws_trucks.cell(
            row=r_idx,
            column=6,
            value=f"=SUMIF({driver_range_ref}, A{r_idx}, {qty_range_ref})",
        )
        ws_trucks.cell(row=r_idx, column=7, value=tr["return_home_str"])
        ws_trucks.cell(
            row=r_idx, column=8, value=format_duration_hhmm(tr["total_shift_min"])
        )
        ws_trucks.cell(row=r_idx, column=9, value=" | ".join(status_notes))
        ws_trucks.cell(row=r_idx, column=10, value=full_route_text)

        for col_idx in range(1, 11):
            c = ws_trucks.cell(row=r_idx, column=col_idx)
            c.border = thin_border
            c.font = Font(name="Arial", size=10, bold=(col_idx == 1))
            if col_idx == 9:
                c.fill = ok_fill
                c.font = ok_font
            if col_idx in (5, 6):
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0"
            elif col_idx in (3, 7, 8):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

        ws_trucks.row_dimensions[r_idx].height = 22

    t_end_row = t_start_row + len(final_trucks) - 1
    t_total_row = t_end_row + 1
    ws_trucks.cell(row=t_total_row, column=1, value=f'סה"כ ({len(final_trucks)} נהגים)')
    if t_end_row >= t_start_row:
        ws_trucks.cell(
            row=t_total_row, column=5, value=f"=SUM(E{t_start_row}:E{t_end_row})"
        )
        ws_trucks.cell(
            row=t_total_row, column=6, value=f"=SUM(F{t_start_row}:F{t_end_row})"
        )

    for col_idx in range(1, 11):
        c = ws_trucks.cell(row=t_total_row, column=col_idx)
        c.font = Font(name="Arial", size=10, bold=True, color="1F4E78")
        c.fill = PatternFill(
            start_color=TOTAL_FILL, end_color=TOTAL_FILL, fill_type="solid"
        )
        c.border = total_border
        if col_idx in (5, 6):
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = "#,##0"
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")
    ws_trucks.row_dimensions[t_total_row].height = 24

    for col_idx, width in enumerate([28, 24, 15, 52, 14, 22, 18, 26, 32, 70], start=1):
        ws_trucks.column_dimensions[get_column_letter(col_idx)].width = width

    if t_end_row >= t_start_row:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = 'סה"כ משקל העמסה (ק"ג) לפי נהג / משאית'
        chart.y_axis.title = 'משקל (ק"ג)'
        chart.x_axis.title = "נהג / משאית"
        chart.width = 18
        chart.height = 10

        data_ref = Reference(ws_trucks, min_col=6, min_row=1, max_row=t_end_row)
        cats_ref = Reference(ws_trucks, min_col=1, min_row=t_start_row, max_row=t_end_row)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.legend = None
        ws_trucks.add_chart(chart, f"A{t_total_row + 2}")

    output_path = pathlib.Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb_out.save(output_path)
    return sorted_records, final_trucks


# ==============================================================================
# 7. פתיחת חלון התיקייה ב-Windows ובחירת קובץ קלט (תוך דילוג על קובץ החומ"ס)
# ==============================================================================
def open_folder_in_windows_explorer(target_path: pathlib.Path | str) -> None:
    p = pathlib.Path(target_path)
    target_str = str(p)
    if os.name == "nt":
        if p.is_file() and p.exists():
            try:
                subprocess.run(
                    f'explorer /select,"{target_str}"', shell=True, check=False
                )
                return
            except Exception:
                pass
            target_str = str(p.parent)
        elif p.is_dir() and p.exists():
            target_str = str(p)
        else:
            target_str = str(p.parent)

        try:
            os.startfile(target_str)
            return
        except Exception:
            pass
        try:
            subprocess.run(["explorer", target_str], check=False)
        except Exception:
            pass


def prompt_gui_file_picker() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="בחר קובץ הזמנות למיון (Excel)",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")],
        )
        root.destroy()
        return selected if selected else None
    except Exception:
        return None


def resolve_input_file(cli_input: str | None) -> tuple[pathlib.Path | None, str | None]:
    if cli_input:
        p = pathlib.Path(cli_input.strip('"'))
        if p.exists():
            return p, p.name
        candidate = pathlib.Path(INPUT_FOLDER) / cli_input.strip('"')
        if candidate.exists():
            return candidate, candidate.name

    default_path = pathlib.Path(INPUT_FOLDER) / INPUT_FILE_NAME
    if default_path.exists():
        return default_path, INPUT_FILE_NAME

    excluded_names = {
        HAZMAT_FILE_NAME.lower(),
        COOLING_FILE_NAME.lower(),
        CONFIG_FILE_NAME.lower(),
        "dangerous_products.xlsx",
        "cooling_products.xlsx",
        "settings.xlsx",
        "config.xlsx",
    }

    in_dir = pathlib.Path(INPUT_FOLDER)
    if in_dir.exists() and in_dir.is_dir():
        xlsx_files = sorted(
            f
            for f in in_dir.glob("*.xlsx")
            if not f.name.startswith("~$")
            and f.name.lower() not in excluded_names
            and "dangerous" not in f.name.lower()
            and "cooling" not in f.name.lower()
        )
        if xlsx_files:
            return xlsx_files[0], xlsx_files[0].name

    gui_selected = prompt_gui_file_picker()
    if gui_selected:
        gp = pathlib.Path(gui_selected)
        return gp, gp.name

    return None, None


def main() -> None:
    cfg_loaded = load_dynamic_config()

    parser = argparse.ArgumentParser(
        description="מיון הזמנות לפי אזורי חלוקה, העמסה ממחסנים (רגיל/קירור/חומ\"ס) וחישוב שעות הגעה"
    )
    parser.add_argument(
        "input_excel",
        nargs="?",
        default=None,
        help="נתיב לקובץ האקסל המקורי (או גרירת קובץ Drag & Drop)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="נתיב לקובץ האקסל הממוין שייווצר (ברירת מחדל: בתוך תיקיית output_files)",
    )
    parser.add_argument(
        "-c",
        "--config",
        default=None,
        help="נתיב לקובץ הגדרות מותאם אישית (ברירת מחדל: input_files\\config_parameters.txt)",
    )
    parser.add_argument(
        "--main-load-minutes",
        type=float,
        default=None,
        help="זמן העמסה בדקות במחסן הרגיל במבוא הספנים 1 (ברירת מחדל: 60)",
    )
    parser.add_argument(
        "--cooling-load-minutes",
        type=float,
        default=None,
        help="זמן העמסה בדקות במחסן הקירור במבוא הספנים 6 (ברירת מחדל: 30)",
    )
    parser.add_argument(
        "--hazmat-load-minutes",
        type=float,
        default=None,
        help='זמן העמסה בדקות במחסן חומ"ס בפארק נמלי ישראל (ברירת מחדל: 30)',
    )
    parser.add_argument(
        "-u",
        "--unload-minutes",
        type=float,
        default=None,
        help="זמן פריקה בדקות בכל עצירה אצל לקוח (ברירת מחדל: 30)",
    )
    parser.add_argument(
        "--max-weight",
        type=float,
        default=None,
        help='משקל מקסימלי למשאית בק"ג לפני פיצול ל-2 נהגים (ברירת מחדל: 7500)',
    )
    parser.add_argument(
        "--max-hours",
        type=float,
        default=None,
        help="זמן משמרת מקסימלי לנהג בשעות מנתניה ועד החזרה לנתניה (ברירת מחדל: 12)",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="הרצה ללא פניות ל-API חיצוני",
    )
    parser.add_argument(
        "--prefer-site-name",
        action="store_true",
        help="מתן עדיפות לעיר המופיעה בעמודת 'שם האתר' על פני עמודת 'יעד'",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="ביטול פתיחה אוטומטית של תיקיית הפלט בסיום",
    )
    args = parser.parse_args()

    if args.config:
        cfg_loaded = load_dynamic_config(args.config)

    eff_main_load = (
        args.main_load_minutes
        if args.main_load_minutes is not None
        else MAIN_WAREHOUSE_LOAD_MINUTES
    )
    eff_cooling_load = (
        args.cooling_load_minutes
        if args.cooling_load_minutes is not None
        else COOLING_WAREHOUSE_LOAD_MINUTES
    )
    eff_hazmat_load = (
        args.hazmat_load_minutes
        if args.hazmat_load_minutes is not None
        else HAZMAT_WAREHOUSE_LOAD_MINUTES
    )
    eff_unload_min = (
        args.unload_minutes if args.unload_minutes is not None else UNLOAD_MINUTES
    )
    eff_max_weight = (
        args.max_weight if args.max_weight is not None else MAX_TRUCK_WEIGHT_KG
    )
    eff_max_hours = (
        args.max_hours if args.max_hours is not None else MAX_DRIVER_SHIFT_HOURS
    )
    eff_use_api = USE_EXTERNAL_ROUTING_API and (not args.offline)
    eff_prefer_site = PREFER_SITE_NAME_CITY or args.prefer_site_name

    input_path, input_file_name = resolve_input_file(args.input_excel)
    if not input_path or not input_path.exists():
        print("לא נבחר קובץ אקסל להזמנות או שהקובץ אינו קיים.")
        print(f"אנא שים את קובץ ההזמנות בתיקייה:\n  {INPUT_FOLDER}")
        return

    if args.output:
        out_path = pathlib.Path(args.output)
    else:
        file_prefix = (
            input_file_name[:-5]
            if input_file_name.lower().endswith(".xlsx")
            else input_file_name
        )
        out_path = pathlib.Path(OUTPUT_FOLDER) / f"{file_prefix}_sorted.xlsx"

    out_path.parent.mkdir(parents=True, exist_ok=True)

    records, trucks = sort_orders_excel(
        input_path=input_path,
        output_path=out_path,
        use_external_api=eff_use_api,
        prefer_site_name_city=eff_prefer_site,
        max_truck_weight=eff_max_weight,
        max_shift_hours=eff_max_hours,
        main_load_minutes=eff_main_load,
        cooling_load_minutes=eff_cooling_load,
        hazmat_load_minutes=eff_hazmat_load,
        unload_minutes=eff_unload_min,
    )

    print(f"\n=======================================================")
    print(f"תכנון המסלולים, העמסת המחסנים וחישוב שעות ההגעה הושלמו בהצלחה!")
    if cfg_loaded:
        print(f"קובץ הגדרות פעיל: {cfg_loaded}")
    print(
        f"זמני העמסה: רגיל={int(eff_main_load)} דק' | קירור={int(eff_cooling_load)} דק' | "
        f'חומ"ס={int(eff_hazmat_load)} דק\' | פריקה אצל לקוח={int(eff_unload_min)} דק\''
    )
    print(f"קובץ הפלט נשמר בנתיב:\n  {out_path}")
    print(f"סה\"כ שורות הזמנה: {len(records)} | סה\"כ נהגים/משאיות: {len(trucks)}")
    for tr in trucks:
        print(
            f"  * {tr['driver_name']}: העמסה [{tr['warehouses_summary_str']}] | "
            f"{len(tr['stop_schedules'])} עצירות | {tr['total_weight']:,.0f} ק\"ג | "
            f"חזרה לנתניה: {tr['return_home_str']} (משמרת: {format_duration_hhmm(tr['total_shift_min'])} שעות)"
        )
    print(f"=======================================================\n")

    if AUTO_OPEN_OUTPUT_FOLDER and not args.no_open:
        open_folder_in_windows_explorer(out_path)


if __name__ == "__main__":
    main()