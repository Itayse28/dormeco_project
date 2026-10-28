@echo off
chcp 65001 >nul
title Dormeco - Sorting Orders & Generating Picking Lists
cd /d "%~dp0"

echo ===================================================================
echo   שלב 1 מתוך 2: מיון הזמנות, חישוב מסלולי נהגים ושעות הגעה...
echo ===================================================================

if "%~1"=="" (
    python sort_orders_by_region.py --no-open
) else (
    python sort_orders_by_region.py "%~1" --no-open
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [שגיאה] שלב מיון ההזמנות נכשל או בוטל.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===================================================================
echo   שלב 2 מתוך 2: הפקת רשימות ליקוט לפי נהגים ולפי מחסנים...
echo ===================================================================

if "%~1"=="" (
    python generate_regional_picking_lists.py
) else (
    python generate_regional_picking_lists.py "%~1"
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [שגיאה] שלב הפקת רשימות הליקוט נכשל.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===================================================================
echo   כל הפעולות הסתיימו בהצלחה! תיקיית output_files נפתחת כעת.
echo ===================================================================
timeout /t 3 >nul