"""
קובץ הגדרות מקומי (דורמקו) - config_files.py
============================================
ההגדרות מתחלקות לשני קבצים בתיקיית input_files:
- config_parameters.txt - ההגדרות המשותפות (נשמר ב-git).
- config_local.txt      - הגדרות של המחשב הזה בלבד: נתיבים (BASE_DIR, SAVE_AS_PATH).
                          הקובץ ב-.gitignore ולא עולה ל-git.
config_local.txt נקרא אחרי config_parameters.txt, וכל פרמטר שמופיע בו גובר על
הערך שבקובץ הראשי. אפשר להעביר אליו גם פרמטרים פרטיים אחרים (מיילים, מפתח API).
"""

from __future__ import annotations

import pathlib

LOCAL_CONFIG_FILE_NAME = "config_local.txt"

LOCAL_CONFIG_TEMPLATE = """# ==============================================================================
# הגדרות מקומיות - של המחשב הזה בלבד (הקובץ לא עולה ל-git)
# כל פרמטר כאן גובר על הערך שבקובץ config_parameters.txt
# ==============================================================================

# תיקיית הפרויקט (בתוכה input_files ו-output_files). ריק = התיקייה שבה נמצא הפרויקט
BASE_DIR =
# תיקייה לשמירת עותק של כל הקבצים (כפתור "שמור בשם" בחלון העריכה)
SAVE_AS_PATH =
"""


def read_kv_file(path: pathlib.Path) -> dict[str, str]:
    """קורא קובץ הגדרות בפורמט KEY = VALUE (שורות שמתחילות ב-# מדולגות)."""
    raw_kv: dict[str, str] = {}
    if not path.exists():
        return raw_kv
    try:
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                raw_kv[k.strip()] = v.strip()
    except Exception as exc:  # noqa: BLE001
        print(f"[!] אזהרה בקריאת {path.name} ({exc}).")
    return raw_kv


def local_config_path(cfg_path: pathlib.Path) -> pathlib.Path:
    """הקובץ המקומי נמצא באותה תיקייה של config_parameters.txt."""
    return cfg_path.parent / LOCAL_CONFIG_FILE_NAME


def read_local_overrides(cfg_path: pathlib.Path, create: bool = False) -> dict[str, str]:
    """הפרמטרים מ-config_local.txt. create=True יוצר קובץ ריק לדוגמה אם הוא חסר."""
    local_path = local_config_path(cfg_path)
    if create and not local_path.exists() and local_path.parent.exists():
        try:
            local_path.write_text(LOCAL_CONFIG_TEMPLATE, encoding="utf-8")
            print(f"[*] נוצר קובץ הגדרות מקומי: {local_path}")
        except Exception:  # noqa: BLE001
            pass
    return read_kv_file(local_path)
