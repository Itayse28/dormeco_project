@echo off
rem ===========================================================================
rem  Dormaco - מיון הזמנות, רשימות ליקוט ודוחות הובלה
rem  מפעיל את run_all.pyw עם pythonw (בלי חלון שחור) ונסגר מיד.
rem  הפלט נשמר ב-logs\run_log.txt. להרצה עם חלון שחור (לבדיקת תקלות): run_all_console.bat
rem ===========================================================================
chcp 65001 > nul

where pythonw > nul 2> nul
if errorlevel 1 goto try_pyw
start "" pythonw "%~dp0src\run_all.pyw" %*
exit /b 0

:try_pyw
where pyw > nul 2> nul
if errorlevel 1 goto fallback
start "" pyw "%~dp0src\run_all.pyw" %*
exit /b 0

:fallback
rem אין pythonw - מריצים עם חלון רגיל
call "%~dp0run_all_console.bat" %*