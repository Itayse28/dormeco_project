@echo off
chcp 65001 >nul
title הפקת רשימות ליקוט אזוריות - דורמקו

:: מעבר לתיקייה שבה נמצא קובץ ה-BAT
cd /d "%~dp0"

echo ========================================================
echo         מערכת הפקת רשימות ליקוט לפי אזורים - דורמקו
echo ========================================================
echo.

:: בדיקה אם נגרר קובץ אקסל ישירות על גבי האייקון (Drag and Drop)
if "%~1"=="" (
    echo [מידע] מפעיל את התוכנית...
    echo אם לא קיים קובץ בתיקיית input_files, ייפתח חלון לבחירת קובץ.
    echo.
    py generate_regional_picking_lists.py || python generate_regional_picking_lists.py
) else (
    echo [מידע] מעבד את הקובץ שנגרר:
    echo "%~1"
    echo.
    py generate_regional_picking_lists.py "%~1" || python generate_regional_picking_lists.py "%~1"
)

echo.
echo [סיום] הפעולה הסתיימה בהצלחה! תיקיית התוצאות נפתחה.
echo.
pause