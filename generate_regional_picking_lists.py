#!/usr/bin/env python3
"""
מערכת אוטומטית להפקת רשימות ליקוט (Picking Lists) נפרדות לכל אזור חלוקה (Python 3.11/3.12/3.13+)
====================================================================================================
הסקריפט מותאם לשימוש קל ואינטואיטיבי (גם למשתמשים שאינם טכנולוגיים):
1. תמיכה מלאה בגרירת קובץ אקסל (Drag & Drop) על קיצור דרך בשולחן העבודה / קובץ .bat.
2. תמיכה בבחירת קובץ באמצעות חלון פתיחת קבצים גרפי (File Picker) בלחיצה כפולה.
3. פתיחה אוטומטית של תיקיית התוצאות (Goods_picking_<Name>) בסיום העבודה.

נתיבי ברירת מחדל:
    תיקיית קלט: C:\\Users\\itaik\\OneDrive\\מסמכים\\dormeco-sorting-script\\input_files\\
    תיקיית פלט: C:\\Users\\itaik\\OneDrive\\מסמכים\\dormeco-sorting-script\\output_files\\Goods_picking_<Name>\\

שמות קבצי הפלט באנגלית:
    - list_North_east.xlsx
    - list_North_west.xlsx
    - list_Jerusalem.xlsx
    - list_Sharon.xlsx
    - list_Gush_Dan.xlsx
    - list_South.xlsx

דרישות התקנה:
    pip install openpyxl requests pandas
"""

from __future__ import annotations

import argparse
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
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests

# ==============================================================================
# 0. הגדרת נתיבי הקלט והפלט
# ==============================================================================
INPUT_FILE_NAME = "list2.xlsx"

BASE_DIR = r"C:\Users\itaik\OneDrive\מסמכים\dormeco-sorting-script"
INPUT_FOLDER = r"C:\Users\itaik\OneDrive\מסמכים\dormeco-sorting-script\input_files"
OUTPUT_FOLDER = r"C:\Users\itaik\OneDrive\מסמכים\dormeco-sorting-script\output_files"

# האם לחתוך את הסיומת .xlsx משם תיקיית המשנה Goods_picking_ (True = Goods_picking_list)
STRIP_EXTENSION_FROM_FOLDER_NAME: bool = True

# האם ליצור קובץ אקסל גם לאזורים שבהם אין הזמנות באותו יום?
CREATE_EMPTY_REGION_FILES: bool = False

# האם לפתוח אוטומטית את תיקיית התוצאות בסייר הקבצים של Windows בסיום?
AUTO_OPEN_OUTPUT_FOLDER: bool = True

# נקודת מוצא של המשאית לחישוב סדר המסלול המהיר ביותר: נמל אשדוד
ASHDOD_PORT_COORDS: tuple[float, float] = (31.8305, 34.6425)

# סף משקל מקסימלי לאזור חלוקה (בק"ג) לסימון התראה באדום
MAX_REGION_WEIGHT_KG: float = 7500.0

# ==============================================================================
# 1. הגדרת סדר אזורי החלוקה ומיפוי שמות הקבצים לאנגלית
# ==============================================================================
REGION_ORDER: list[str] = [
    "צפון מזרח",
    "צפון מערב",
    "איזור ירושלים",
    "אזור השרון",
    "איזור גוש דן",
    "דרום הארץ",
]

REGION_TO_ENGLISH: dict[str, str] = {
    "צפון מזרח": "North_east",
    "צפון מערב": "North_west",
    "איזור ירושלים": "Jerusalem",
    "אזור השרון": "Sharon",
    "איזור גוש דן": "Gush_Dan",
    "דרום הארץ": "South",
}

REGION_RANK: dict[str, int] = {name: idx for idx, name in enumerate(REGION_ORDER)}
CACHE_FILE = pathlib.Path("geocoding_cache.json")

# ==============================================================================
# 2. מילון יישובים ואזורי תעשייה כולל קואורדינטות GPS
# ==============================================================================
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
# 3. פונקציות עזר לקריאה בטוחה של אקסל וחישוב כמות אריזות לליקוט
# ==============================================================================
def load_workbook_safe(excel_path: str | pathlib.Path) -> openpyxl.Workbook:
    excel_path = pathlib.Path(excel_path)
    try:
        return openpyxl.load_workbook(excel_path)
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
        return openpyxl.load_workbook(bio)


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


# ==============================================================================
# 4. סיווג גיאוגרפי וחישוב המסלול המהיר מנמל אשדוד
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
        headers = {"User-Agent": "DormacoPickingListGenerator/3.0 (Python/3.12)"}
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


def haversine_km(c1: tuple[float, float], c2: tuple[float, float]) -> float:
    lat1, lon1 = math.radians(c1[0]), math.radians(c1[1])
    lat2, lon2 = math.radians(c2[0]), math.radians(c2[1])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    )
    return 6371.0 * (2 * math.asin(math.sqrt(a))) * 1.25


def fetch_osrm_duration_matrix(
    points: list[tuple[float, float]], use_external_api: bool = True
) -> list[list[float]] | None:
    if not use_external_api or len(points) <= 2 or len(points) > 50:
        return None
    coords_str = ";".join(f"{lon:.5f},{lat:.5f}" for lat, lon in points)
    url = f"https://router.project-osrm.org/table/v1/driving/{coords_str}?annotations=duration"
    try:
        resp = requests.get(url, timeout=5)
        if resp.status_code == 200:
            data = resp.json()
            durations = data.get("durations")
            if durations and len(durations) == len(points):
                return [[float(x if x is not None else 999999.0) for x in row] for row in durations]
    except Exception:
        pass
    return None


def compute_fastest_route_order(
    stop_coords: list[tuple[float, float]],
    origin: tuple[float, float] = ASHDOD_PORT_COORDS,
    use_external_api: bool = True,
) -> list[int]:
    k = len(stop_coords)
    if k <= 1:
        return list(range(k))

    all_points = [origin] + stop_coords
    matrix = fetch_osrm_duration_matrix(all_points, use_external_api=use_external_api)

    if matrix is None:
        n = len(all_points)
        matrix = [
            [haversine_km(all_points[i], all_points[j]) for j in range(n)]
            for i in range(n)
        ]

    if k <= 9:
        best_perm: tuple[int, ...] = tuple(range(1, k + 1))
        best_cost = float("inf")
        for perm in itertools.permutations(range(1, k + 1)):
            cost = matrix[0][perm[0]]
            for idx in range(len(perm) - 1):
                cost += matrix[perm[idx]][perm[idx + 1]]
                if cost >= best_cost:
                    break
            if cost < best_cost:
                best_cost = cost
                best_perm = perm
        return [p - 1 for p in best_perm]

    unvisited = set(range(1, k + 1))
    current = 0
    tour: list[int] = []
    while unvisited:
        nxt = min(unvisited, key=lambda node: matrix[current][node])
        tour.append(nxt)
        unvisited.remove(nxt)
        current = nxt

    improved = True
    while improved:
        improved = False
        for i in range(len(tour) - 1):
            for j in range(i + 1, len(tour)):
                prev_node = 0 if i == 0 else tour[i - 1]
                next_node = tour[j + 1] if j + 1 < len(tour) else None
                old_edge = matrix[prev_node][tour[i]] + (
                    matrix[tour[j]][next_node] if next_node is not None else 0.0
                )
                new_edge = matrix[prev_node][tour[j]] + (
                    matrix[tour[i]][next_node] if next_node is not None else 0.0
                )
                if new_edge + 1e-6 < old_edge:
                    tour[i : j + 1] = reversed(tour[i : j + 1])
                    improved = True

    return [node - 1 for node in tour]


# ==============================================================================
# 5. יצירת קובץ אקסל של רשימת ליקוט (Picking List) עבור אזור בודד
# ==============================================================================
def write_region_picking_list_excel(
    region_name: str,
    reg_records: list[dict[str, Any]],
    output_file_path: str | pathlib.Path,
    max_region_weight: float = MAX_REGION_WEIGHT_KG,
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    eng_name = REGION_TO_ENGLISH.get(region_name, region_name)
    ws.title = f"Picking_{eng_name}"[:31]
    ws.sheet_view.rightToLeft = True
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    NAVY_HEADER = "1F4E78"
    ZEBRA_FILL = "F9FAFC"
    TOTAL_FILL = "EAEEF7"
    BORDER_GRAY = "D9D9D9"
    OVERWEIGHT_FILL_HEX = "FFC7CE"
    OVERWEIGHT_FONT_HEX = "9C0006"

    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
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

    for col_idx, h_text in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.border = thin_border
        c.alignment = Alignment(
            horizontal="right" if col_idx in numeric_cols else "left",
            vertical="center",
        )
    ws.row_dimensions[1].height = 25
    ws.freeze_panes = "A2"

    total_region_weight = sum(float(r["qty"]) for r in reg_records)
    is_overweight = total_region_weight > max_region_weight

    start_row = 2
    for idx, rec in enumerate(reg_records):
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
            start_color=ZEBRA_FILL if r_idx % 2 == 0 else "FFFFFF",
            end_color=ZEBRA_FILL if r_idx % 2 == 0 else "FFFFFF",
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

    if reg_records:
        end_row = start_row + len(reg_records) - 1
        total_row = end_row + 1
        ws.cell(
            row=total_row,
            column=1,
            value=f'סה"כ ({region_name})' + (' - חריגת משקל מעל 7,500 ק"ג!' if is_overweight else ""),
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
                formula=[f"$E${total_row}>{int(max_region_weight)}"],
                stopIfTrue=True,
                fill=overweight_fill,
                font=overweight_font,
            ),
        )
        ws.conditional_formatting.add(
            f"E{total_row}:E{total_row}",
            CellIsRule(
                operator="greaterThan",
                formula=[str(int(max_region_weight))],
                stopIfTrue=True,
                fill=overweight_fill,
                font=overweight_font,
            ),
        )

    max_check_row = start_row + len(reg_records) - 1
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
# 6. פונקציה ראשית: קריאת קובץ הקלט ופיצול לקבצי ליקוט אזוריים באנגלית
# ==============================================================================
def generate_picking_lists_by_region(
    input_file_path: str | pathlib.Path,
    output_dir: str | pathlib.Path = OUTPUT_FOLDER,
    input_file_name_override: str | None = None,
    use_external_api: bool = True,
    prefer_site_name_city: bool = False,
    max_region_weight: float = MAX_REGION_WEIGHT_KG,
    create_empty_regions: bool = CREATE_EMPTY_REGION_FILES,
    use_subfolder: bool = True,
) -> tuple[dict[str, pathlib.Path], pathlib.Path]:
    """
    קוראת את קובץ ההזמנות המקורי, ממיינת כל אזור לפי המסלול המהיר מנמל אשדוד,
    ומייצרת תיקייה ייעודית בתוך output_files בשם:
        Goods_picking_<INPUT_FILE_NAME>
    ובתוכה שומרת קובץ אקסל לכל אזור בשם:
        <INPUT_FILE_NAME[:-5]>_<EnglishRegion>.xlsx (למשל list_Sharon.xlsx)
    """
    input_path = pathlib.Path(input_file_path)
    current_input_name = input_file_name_override or input_path.name

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

    idx_cust_name = find_col(["שם לקוח"], 1)
    idx_address = find_col(["כתובת"], 2)
    idx_dest = find_col(["יעד", "עיר"], 3)
    idx_site = find_col(["שם האתר", "אתר"], 4)
    idx_phone = find_col(["מספר טלפון", "טלפון"], 5)
    idx_order = find_col(["הזמנה"], 6)
    idx_line = find_col(["שורה"], 8)
    idx_sku = find_col(["מק'ט", 'מק"ט', "מקט"], 9)
    idx_desc = find_col(["תאור מוצר", "תיאור מוצר"], 10)
    idx_qty = find_col(["כמות"], 11)
    idx_note_cust = find_col(["הערה למוצר ללקוח", "הערה ללקוח"], 12)
    idx_note_int = find_col(["הערה פנימית להזמנה", "הערה פנימית"], 14)

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

        order_no = str(row_vals[idx_order] or "").strip()
        cust_name = str(row_vals[idx_cust_name] or "").strip()
        sku = str(row_vals[idx_sku] or "").strip()
        product_desc = str(row_vals[idx_desc] or "").strip()

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
            sku=sku,
            description=product_desc,
            qty=qty_val,
            customer_note=str(cust_note or ""),
        )

        stop_key = f"{effective_dest}||{cust_name}||{lat:.5f},{lon:.5f}"

        records.append(
            {
                "region": region,
                "region_rank": REGION_RANK.get(region, 99),
                "effective_dest": effective_dest,
                "order_no": order_no,
                "cust_name": cust_name,
                "sku": sku,
                "product_desc": product_desc,
                "qty": qty_clean,
                "packages": packages_count,
                "notes": notes_combined,
                "line_no": int(row_vals[idx_line] or 0)
                if row_vals[idx_line] is not None
                else 0,
                "lat": lat,
                "lon": lon,
                "stop_key": stop_key,
                "route_stop_num": 1,
            }
        )

    file_prefix = (
        current_input_name[:-5]
        if current_input_name.lower().endswith(".xlsx")
        else pathlib.Path(current_input_name).stem
    )

    base_out_dir = pathlib.Path(str(output_dir).rstrip("\\/"))
    if use_subfolder:
        folder_suffix = file_prefix if STRIP_EXTENSION_FROM_FOLDER_NAME else current_input_name
        target_out_dir = base_out_dir / f"Goods_picking_{folder_suffix}"
    else:
        target_out_dir = base_out_dir

    target_out_dir.mkdir(parents=True, exist_ok=True)
    created_files: dict[str, pathlib.Path] = {}

    for reg in REGION_ORDER:
        reg_records = [rec for rec in records if rec["region"] == reg]
        if not reg_records and not create_empty_regions:
            continue

        if reg_records:
            unique_stops: list[tuple[str, tuple[float, float]]] = []
            seen_keys: set[str] = set()
            for rec in reg_records:
                s_key = rec["stop_key"]
                if s_key not in seen_keys:
                    seen_keys.add(s_key)
                    unique_stops.append((s_key, (rec["lat"], rec["lon"])))

            stop_coords = [coords for _, coords in unique_stops]
            optimal_order = compute_fastest_route_order(
                stop_coords,
                origin=ASHDOD_PORT_COORDS,
                use_external_api=use_external_api,
            )

            stop_key_to_route_num: dict[str, int] = {}
            for route_idx, stop_idx in enumerate(optimal_order, start=1):
                s_key = unique_stops[stop_idx][0]
                stop_key_to_route_num[s_key] = route_idx

            for rec in reg_records:
                rec["route_stop_num"] = stop_key_to_route_num.get(rec["stop_key"], 999)

            reg_records.sort(
                key=lambda item: (
                    item["route_stop_num"],
                    item["order_no"],
                    item["line_no"],
                )
            )

        eng_region = REGION_TO_ENGLISH.get(reg, reg)
        out_filename = f"{file_prefix}_{eng_region}.xlsx"
        out_file_path = target_out_dir / out_filename

        write_region_picking_list_excel(
            region_name=reg,
            reg_records=reg_records,
            output_file_path=out_file_path,
            max_region_weight=max_region_weight,
        )
        created_files[eng_region] = out_file_path

    return created_files, target_out_dir


def open_folder_in_windows_explorer(folder_path: pathlib.Path | str) -> None:
    """פותחת את תיקיית התוצאות ישירות בסייר הקבצים של Windows למשתמש."""
    folder_str = str(folder_path)
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


def prompt_gui_file_picker() -> str | None:
    """פותחת חלון גרפי של Windows לבחירת קובץ אקסל אם לא הועבר קובץ."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="בחר קובץ הזמנות (Excel)",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")],
        )
        root.destroy()
        return selected if selected else None
    except Exception:
        return None


def resolve_input_file(cli_input: str | None) -> tuple[pathlib.Path | None, str | None]:
    """
    מאתרת את קובץ הקלט:
    1. אם הועבר נתיב בשורת הפקודה או נגרר על הקובץ (Drag & Drop), משתמשת בו.
    2. בודקת אם קיים קובץ ב-INPUT_FOLDER + INPUT_FILE_NAME.
    3. בודקת אם יש קובץ .xlsx כלשהו בתיקיית input_files.
    4. אם אין קובץ מוכן, פותחת חלון לבחירת קובץ (File Picker).
    """
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

    in_dir = pathlib.Path(INPUT_FOLDER)
    if in_dir.exists() and in_dir.is_dir():
        xlsx_files = sorted(
            f for f in in_dir.glob("*.xlsx") if not f.name.startswith("~$")
        )
        if xlsx_files:
            return xlsx_files[0], xlsx_files[0].name

    # אם לא נמצא קובץ בתיקיית input_files, פתיחת חלון בחירת קובץ
    gui_selected = prompt_gui_file_picker()
    if gui_selected:
        gp = pathlib.Path(gui_selected)
        return gp, gp.name

    return None, None


def main() -> None:
    parser = argparse.ArgumentParser(
        description="הפקת רשימות ליקוט (Picking Lists) באקסל לכל אזור חלוקה בנפרד"
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default=None,
        help="שם קובץ הקלט או נתיב מלא (תומך גם בגרירת קובץ Drag & Drop)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=OUTPUT_FOLDER,
        help="תיקיית הפלט הראשית (ברירת מחדל: output_files)",
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
        "--include-empty",
        action="store_true",
        help="יצירת קבצי ליקוט גם לאזורים ללא הזמנות",
    )
    parser.add_argument(
        "--no-subfolder",
        action="store_true",
        help="שמירה ישירות בתיקיית הפלט ללא תיקיית משנה Goods_picking_",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="ביטול פתיחה אוטומטית של התיקייה בסיום",
    )
    args = parser.parse_args()

    input_path, input_file_name = resolve_input_file(args.input_file)
    if not input_path or not input_path.exists():
        print("לא נבחר קובץ אקסל או שהקובץ אינו קיים.")
        print(f"אנא שים קובץ בתיקייה:\n  {INPUT_FOLDER}")
        print("או גרור קובץ אקסל ישירות על גבי קיצור הדרך.")
        return

    created, target_dir = generate_picking_lists_by_region(
        input_file_path=input_path,
        output_dir=args.output_dir,
        input_file_name_override=input_file_name,
        use_external_api=not args.offline,
        prefer_site_name_city=args.prefer_site_name,
        create_empty_regions=args.include_empty or CREATE_EMPTY_REGION_FILES,
        use_subfolder=not args.no_subfolder,
    )

    print(f"\n=======================================================")
    print(f"המיון והפקת רשימות הליקוט מתוך '{input_file_name}' הושלמו בהצלחה!")
    print(f"הקבצים נשמרו בתיקייה:\n  {target_dir}\n")
    for eng_reg, fpath in created.items():
        print(f"  - [{eng_reg}]: {fpath.name}")
    print(f"=======================================================\n")

    if AUTO_OPEN_OUTPUT_FOLDER and not args.no_open:
        open_folder_in_windows_explorer(target_dir)


if __name__ == "__main__":
    main()