@echo off
chcp 65001 >nul
title דורמקו - מיון הזמנות והפקת רשימות ליקוט

:: מעבר לתיקייה שבה נמצא קובץ ה-BAT
cd /d "%~dp0"

echo ========================================================
echo       דורמקו - מיון הזמנות והפקת רשימות ליקוט אזוריות
echo ========================================================
echo.

if "%~1"=="" (
    echo [שלב 1/2] מפעיל מיון הזמנות ומסלול...
    py sort_orders_by_region.py --no-open || python sort_orders_by_region.py --no-open
    echo.
    echo [שלב 2/2] מפיק רשימות ליקוט לפי אזורים...
    py generate_regional_picking_lists.py || python generate_regional_picking_lists.py
) else (
    echo [שלב 1/2] ממיין את הקובץ: "%~1"...
    py sort_orders_by_region.py "%~1" --no-open || python sort_orders_by_region.py "%~1" --no-open
    echo.
    echo [שלב 2/2] מפיק רשימות ליקוט עבור: "%~1"...
    py generate_regional_picking_lists.py "%~1" || python generate_regional_picking_lists.py "%~1"
)

echo.
echo ========================================================
echo [הצלחה] שתי הפעולות הסתיימו בהצלחה! תיקיית התוצאות פתוחה.
echo ========================================================
echo.
pause