@echo off
chcp 65001 >nul
title מיון הזמנות לפי אזורים ומסלול - דורמקו

:: מעבר לתיקייה שבה נמצא קובץ ה-BAT
cd /d "%~dp0"

echo ========================================================
echo       מערכת מיון הזמנות לפי אזורי חלוקה ומסלול - דורמקו
echo ========================================================
echo.

:: בדיקה אם נגרר קובץ אקסל ישירות על גבי האייקון (Drag and Drop)
if "%~1"=="" (
    echo [מידע] מפעיל את תוכנית המיון...
    echo אם לא קיים קובץ בתיקיית input_files, ייפתח חלון לבחירת קובץ.
    echo.
    py sort_orders_by_region.py || python sort_orders_by_region.py
) else (
    echo [מידע] מעבד את הקובץ שנגרר:
    echo "%~1"
    echo.
    py sort_orders_by_region.py "%~1" || python sort_orders_by_region.py "%~1"
)

echo.
echo [סיום] המיון הסתיים בהצלחה! תיקיית התוצאות נפתחה.
echo.
pause