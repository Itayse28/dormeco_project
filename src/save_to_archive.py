#!/usr/bin/env python3
r"""
שמירת הדוחות בתיקיות הארכיון (כונן G) - save_to_archive.py
=========================================================
כפתור "שמור בשם" בחלון העריכה מעתיק כל דוח הובלה ודוח העמסות לתיקייה שלו,
מתחת ל-BASE_DIR_DORMECO (מקובץ ההגדרות). תיקייה שלא קיימת נוצרת.

  נהגי דורמקו:   <BASE>\<DORMECO_DRIVERS>\<שם הנהג>\<שנה>\<חודש>\הובלות|העמסות
  אלירן:          <BASE>\<ELIRAN_DIR>\<שנה>\<חודש>\הובלות|העמסות
  אוברסיז:        <BASE>\<OVERSEAS_DIR>\<שנה>\<חודש>\הובלות|העמסות
  איסוף עצמי:     <BASE>\<TAKE_AWAY>\איסוף עצמי <שנה>\<חודש>

השנה והחודש נלקחים מתאריך המשלוח שבשם הקובץ (למשל 23.11.2026 -> 2026\נובמבר),
ואם אין תאריך בשם - מהתאריך של היום.
"""

from __future__ import annotations

import datetime
import pathlib
import re
import shutil
from typing import Any

HEBREW_MONTHS = [
    "ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
    "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר",
]
SHIPPING_PREFIX = "הובלות ע״י "
PICKING_PREFIX = "דוח העמסות עבור "
SHIPPING_SUB = "הובלות"
PICKING_SUB = "העמסות"
PICKUP_NAME = "איסוף עצמי"
ELIRAN_KEY = "אלירן"
OVERSEAS_KEY = "אוברסיז"
CONFIG_KEYS = ("BASE_DIR_DORMECO", "DORMECO_DRIVERS", "TAKE_AWAY", "ELIRAN_DIR", "OVERSEAS_DIR")

_DATE_RE = re.compile(r"\s+(\d{1,2})\.(\d{1,2})\.(\d{4})")


def read_archive_config(raw_kv: dict[str, str]) -> dict[str, str]:
    """הפרמטרים של תיקיות הארכיון מקובץ ההגדרות (בלי מרכאות)."""
    return {k: str(raw_kv.get(k, "") or "").strip().strip('"').strip() for k in CONFIG_KEYS}


def _parse_report(path: pathlib.Path) -> tuple[str, str, datetime.date] | None:
    """(סוג: הובלות/העמסות, שם הנהג, תאריך) מתוך שם הקובץ; None אם זה לא דוח."""
    stem = path.stem
    if stem.startswith(SHIPPING_PREFIX):
        kind, rest = SHIPPING_SUB, stem[len(SHIPPING_PREFIX):]
    elif stem.startswith(PICKING_PREFIX):
        kind, rest = PICKING_SUB, stem[len(PICKING_PREFIX):]
    else:
        return None
    m = _DATE_RE.search(rest)
    if m:
        try:
            when = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            when = datetime.date.today()
        who = rest[: m.start()]
    else:
        when = datetime.date.today()
        who = rest.split(" - ")[0]  # בלי סיומת המחסן ("... - קירור")
    who = who.strip()
    # "נהג יוסי" -> "יוסי"; נהג בלי שם נשאר "נהג 3"
    m_name = re.match(r"^נהג\s+(.+)$", who)
    if m_name and not m_name.group(1).strip().isdigit():
        who = m_name.group(1).strip()
    return kind, who or "ללא שם", when


def target_dir(path: pathlib.Path, acfg: dict[str, str]) -> pathlib.Path | None:
    """תיקיית היעד של דוח (None אם זה לא דוח הובלה/העמסות)."""
    parsed = _parse_report(path)
    if parsed is None:
        return None
    kind, who, when = parsed
    base = pathlib.Path(acfg["BASE_DIR_DORMECO"])
    year, month = str(when.year), HEBREW_MONTHS[when.month - 1]
    if PICKUP_NAME in who:
        return base / (acfg["TAKE_AWAY"] or PICKUP_NAME) / f"{PICKUP_NAME} {year}" / month
    if OVERSEAS_KEY in who and acfg["OVERSEAS_DIR"]:
        return base / acfg["OVERSEAS_DIR"] / year / month / kind
    if ELIRAN_KEY in who and acfg["ELIRAN_DIR"]:
        return base / acfg["ELIRAN_DIR"] / year / month / kind
    safe_who = re.sub(r'[\\/:*?"<>|]+', " ", who).strip()
    return base / acfg["DORMECO_DRIVERS"] / safe_who / year / month / kind


def save_reports(folders: list[pathlib.Path], acfg: dict[str, Any]) -> tuple[list[pathlib.Path], list[str]]:
    """
    מעתיק את כל דוחות האקסל שבתיקיות (כולל תתי-תיקיות התאריכים) לתיקיות הארכיון.
    מחזיר (הקבצים שנשמרו, שגיאות).
    """
    base = pathlib.Path(acfg["BASE_DIR_DORMECO"])
    if not base.exists():
        raise RuntimeError(f"התיקייה BASE_DIR_DORMECO לא נמצאה:\n{base}\n(האם הכונן מחובר?)")
    saved: list[pathlib.Path] = []
    errors: list[str] = []
    for folder in folders:
        if not folder.exists():
            continue
        for f in sorted(folder.rglob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            dest_dir = target_dir(f, acfg)
            if dest_dir is None:
                continue
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = dest_dir / f.name
                shutil.copy2(f, dest)
                saved.append(dest)
            except OSError as exc:
                errors.append(f"{f.name}: {exc}")
    return saved, errors
