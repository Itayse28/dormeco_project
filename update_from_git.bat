@echo off
rem ===========================================================================
rem  Dormaco - עדכון הפרויקט מ-GitHub
rem  מוריד את הגרסה העדכנית מ-GitHub. במקרה של התנגשות - הגרסה שב-GitHub
rem  גוברת, ושינויים מקומיים בקבצים שנמצאים ב-git נדרסים.
rem  קבצים שלא נמצאים ב-git (output_files, logs, config_local.txt) לא נמחקים.
rem ===========================================================================
chcp 65001 > nul
cd /d "%~dp0"

where git > nul 2> nul
if errorlevel 1 (
    echo git is not installed / not in PATH.
    pause
    exit /b 1
)

for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b
echo Updating branch %BRANCH% from GitHub...

git fetch origin
if errorlevel 1 (
    echo.
    echo Fetch failed - check the internet connection / GitHub login.
    pause
    exit /b 1
)

rem cancel a merge/rebase that was left half done
git merge --abort > nul 2> nul
git rebase --abort > nul 2> nul

rem GitHub version wins: local changes to tracked files are overwritten
git reset --hard origin/%BRANCH%
if errorlevel 1 (
    echo.
    echo Update failed.
    pause
    exit /b 1
)

echo.
echo Done - the project matches GitHub.
git log -1 --format="Last change: %%h %%s (%%cr)"
pause
