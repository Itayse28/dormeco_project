@echo off
rem ===========================================================================
rem  Dormaco - משיכת העדכונים מ-GitHub (git pull)
rem  מוריד את כל השינויים מ-GitHub ושומר על שינויים מקומיים שעוד לא הועלו:
rem  שינויים שלא נשמרו ב-git נשמרים בצד ומוחזרים אחרי העדכון (autostash),
rem  ו-commits מקומיים מונחים מעל הגרסה העדכנית (rebase).
rem  אם יש התנגשות - שום דבר לא נדרס: העדכון מבוטל ומוצגת הודעה.
rem  (update_from_git.bat לעומת זאת דורס שינויים מקומיים - הגרסה ב-GitHub גוברת.)
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
echo Pulling branch %BRANCH% from GitHub...
echo.

git pull --rebase --autostash origin %BRANCH%
if errorlevel 1 (
    git rebase --abort > nul 2> nul
    echo.
    echo The update could not be applied automatically - a conflict with local changes,
    echo or no connection to GitHub. Nothing was overwritten.
    echo Ask Claude to resolve it, or use update_from_git.bat to take the GitHub version.
    pause
    exit /b 1
)

echo.
echo Done - the project is up to date.
git log -1 --format="Last change: %%h %%s (%%an, %%cr)"
pause
