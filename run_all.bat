@echo off
chcp 65001 > nul
title Dormaco - מיון הזמנות, רשימות ליקוט ודוחות הובלה לנהגים

rem התקנה/בדיקה שקטה של ספריית המפה (אם אין אינטרנט - תוצג מפה סכמטית)
python -m pip install --quiet --disable-pip-version-check tkintermapview > nul 2>&1

echo ===========================================================================
echo [שלב 1/3] מריץ את סקריפט מיון ההזמנות ותכנון מסלולי הנהגים...
echo           ייפתח חלון לעריכת המסלולים: "אישור והמשך" יוצר את כל הקבצים, "סיום" סוגר
echo ===========================================================================
python "%~dp0sort_orders_by_region.py" %* --no-open

if %ERRORLEVEL% EQU 3 (
    echo.
    echo [V] כל הקבצים נוצרו מתוך חלון העריכה ^(קובץ ממוין, רשימות ליקוט ודוחות הובלה^).
    timeout /t 3 > nul
    exit /b 0
)

if %ERRORLEVEL% EQU 2 (
    echo.
    echo [X] עריכת המסלולים בוטלה. לא נוצרו קבצים.
    pause
    exit /b 2
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] אירעה שגיאה בשלב מיון ההזמנות. עוצר הרצה.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===========================================================================
echo [שלב 2/3] מריץ את סקריפט רשימות הליקוט לפי נהגים ומחסנים...
echo ===========================================================================
python "%~dp0generate_regional_picking_lists.py" %* --no-open

if %ERRORLEVEL% EQU 4 (
    echo.
    echo [!] חלק מרשימות הליקוט לא נשמרו ^(קבצים פתוחים באקסל^). ממשיך לדוחות ההובלה.
    goto step3
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] אירעה שגיאה בשלב יצירת רשימות הליקוט. עוצר הרצה.
    pause
    exit /b %ERRORLEVEL%
)

echo.
:step3
echo ===========================================================================
echo [שלב 3/3] מריץ את סקריפט דוחות ההובלה לנהגים...
echo ===========================================================================
python "%~dp0generate_driver_delivery_reports.py" %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] אירעה שגיאה בשלב יצירת דוחות ההובלה לנהגים.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [V] כל 3 השלבים הסתיימו בהצלחה! תיקיית output_files נפתחה.
timeout /t 3 > nul