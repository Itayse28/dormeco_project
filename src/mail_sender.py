#!/usr/bin/env python3
"""
שליחת קבצים במייל דרך Outlook (דורמקו) - mail_sender.py
======================================================
מודול משותף לדוחות ההעמסות ולדוחות ההובלה.

- create_launcher: יוצר בתיקייה קובץ bat. לחיצה כפולה עליו שולחת כל קובץ
  בתיקייה במייל נפרד דרך Outlook (הגרסה הקלאסית), לנמענים לפי סיומת שם הקובץ.
- run_launcher: מריץ את אותה שליחה מתוך Python (לכפתור "שלח מיילים" בחלון העריכה).
- run_inline: שליחה של קבוצת קבצים מתיקייה בלי קובץ bat (לכפתור "שלח חלקית").
- pdf_single_mail=True: קובצי האקסל מומרים ל-PDF ברגע השליחה (דרך Excel) ונשלחים
  כולם במייל אחד - כך שאפשר לתקן ידנית את האקסל לפני הלחיצה על ה-bat.
"""

from __future__ import annotations

import pathlib
import re
import subprocess
import sys
import tempfile
from typing import Any

MAIL_ATTACH_EXTENSIONS = (".xlsx", ".xls", ".pdf")
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

_BAT_HEADER = r"""@echo off
chcp 65001 > nul
title __TITLE__
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f='%~f0'; $s=[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8); $i=$s.LastIndexOf('#PS-START'); Invoke-Expression $s.Substring($i)"
echo.
pause
exit /b
"""

_PS_BODY = r"""#PS-START
[Console]::OutputEncoding = [Text.Encoding]::UTF8
function Send-FolderFiles {
    $exts = @(__EXTS__)
    # נמענים לפי סיומת שם הקובץ; קובץ בלי סיומת מתאימה מקבל את ברירת המחדל
    $rules = @(__RULES__)
    $defaultTo = __DEFAULT_TO__
    $defaultName = __DEFAULT_NAME__
    $nameFilter = __FILTER__
    $stripPrefix = __STRIP_PREFIX__
    $subjectPrefix = __SUBJECT_PREFIX__
    $folder = Split-Path -Parent $f
    $files = @(Get-ChildItem -LiteralPath $folder -File |
        Where-Object { $exts -contains $_.Extension.ToLower() } |
        Where-Object { ($nameFilter -eq '') -or ($_.BaseName -match $nameFilter) } |
        Sort-Object Name)

    if ($files.Count -eq 0) {
        Write-Host '[X] No files to send were found in this folder.'
        return
    }

    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Host '[X] Could not open Outlook.'
        Write-Host '    Classic (desktop) Outlook must be installed and signed in.'
        Write-Host '    The "new Outlook" app does not support automatic sending.'
        return
    }

    $sent = 0
    foreach ($file in $files) {
        $base = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        $to = $defaultTo
        $whName = $defaultName
        $core = $base
        foreach ($rule in $rules) {
            if ($base.EndsWith($rule.Suffix)) {
                $to = $rule.To; $whName = $rule.Name
                $core = $base.Substring(0, $base.Length - $rule.Suffix.Length)
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($to)) {
            Write-Host ('[!] Skipped (no recipients defined for ' + $whName + '): ' + $file.Name)
            continue
        }
        # נושא: "<קידומת> נהג <שם> <תאריך>"
        if ($core.StartsWith($stripPrefix)) { $core = $core.Substring($stripPrefix.Length) }
        if (-not $core.StartsWith('נהג ')) { $core = 'נהג ' + $core }
        $subject = $subjectPrefix + $core.Trim()
        try {
            $mail = $outlook.CreateItem(0)
            $mail.To = $to
            $mail.Subject = $subject
            $mail.HTMLBody = '<div dir="rtl" style="font-family:Arial;font-size:11pt">' +
                'שלום,<br><br>מצורף:<br><b>' +
                [System.Net.WebUtility]::HtmlEncode($base) + '</b><br><br>' +
                'נשלח אוטומטית ממערכת מיון ההזמנות.</div>'
            [void]$mail.Attachments.Add($file.FullName)
            $mail.Send()
            $sent++
            Write-Host ('[V] Sent (' + $whName + ') to ' + $to + ' : ' + $file.Name)
        } catch {
            Write-Host ('[X] Sending failed for ' + $file.Name + ': ' + $_.Exception.Message)
        }
    }
    try { $outlook.Session.SendAndReceive($false) } catch {}
    Start-Sleep -Seconds 3
    Write-Host ''
    Write-Host ('[V] Done: ' + $sent + ' of ' + $files.Count + ' emails sent.')
}
function Send-FolderAsPdf {
    # ממיר את קובצי האקסל שבתיקייה ל-PDF (ברגע השליחה - כך שתיקונים ידניים באקסל נכללים)
    # ושולח את כל קובצי ה-PDF במייל אחד לנמענים של ברירת המחדל.
    $defaultTo = __DEFAULT_TO__
    $nameFilter = __FILTER__
    $subject = __SINGLE_SUBJECT__
    $folder = Split-Path -Parent $f
    if ([string]::IsNullOrWhiteSpace($subject)) { $subject = Split-Path -Leaf $folder }
    $files = @(Get-ChildItem -LiteralPath $folder -File |
        Where-Object { @('.xlsx', '.xls') -contains $_.Extension.ToLower() } |
        Where-Object { -not $_.Name.StartsWith('~$') } |
        Where-Object { ($nameFilter -eq '') -or ($_.BaseName -match $nameFilter) } |
        Sort-Object Name)

    if ($files.Count -eq 0) {
        Write-Host '[X] No files to send were found in this folder.'
        return
    }
    if ([string]::IsNullOrWhiteSpace($defaultTo)) {
        Write-Host '[X] No recipients defined.'
        return
    }

    $pdfDir = Join-Path ([IO.Path]::GetTempPath()) ('dormeco_pdf_' + [Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $pdfDir)
    Write-Host 'Converting Excel files to PDF (unsaved changes in open files are NOT included)...'
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        Write-Host '[X] Could not open Excel - it is needed to convert the reports to PDF.'
        Remove-Item -LiteralPath $pdfDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $pdfs = @()
    foreach ($file in $files) {
        $pdfPath = Join-Path $pdfDir ($file.BaseName + '.pdf')
        $wb = $null
        try {
            # UpdateLinks=0, ReadOnly=true - עובד גם כשהקובץ פתוח באקסל (נלקחת הגרסה השמורה)
            $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
            $wb.ExportAsFixedFormat(0, $pdfPath)
            $pdfs += $pdfPath
            Write-Host ('[V] Converted: ' + $file.Name)
        } catch {
            Write-Host ('[X] PDF conversion failed for ' + $file.Name + ': ' + $_.Exception.Message)
        } finally {
            if ($wb) { try { $wb.Close($false) } catch {} }
        }
    }
    try { $excel.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

    if ($pdfs.Count -eq 0) {
        Write-Host '[X] No PDF files were created - nothing was sent.'
        Remove-Item -LiteralPath $pdfDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Host '[X] Could not open Outlook.'
        Write-Host '    Classic (desktop) Outlook must be installed and signed in.'
        Write-Host '    The "new Outlook" app does not support automatic sending.'
        Remove-Item -LiteralPath $pdfDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    $sent = 0
    try {
        $names = ($pdfs | ForEach-Object {
            [System.Net.WebUtility]::HtmlEncode([IO.Path]::GetFileNameWithoutExtension($_)) }) -join '<br>'
        $mail = $outlook.CreateItem(0)
        $mail.To = $defaultTo
        $mail.Subject = $subject
        $mail.HTMLBody = '<div dir="rtl" style="font-family:Arial;font-size:11pt">' +
            'שלום,<br><br>מצורפים:<br><b>' + $names + '</b><br><br>' +
            'נשלח אוטומטית ממערכת מיון ההזמנות.</div>'
        foreach ($p in $pdfs) { [void]$mail.Attachments.Add($p) }
        $mail.Send()
        $sent = 1
        Write-Host ('[V] Sent ' + $pdfs.Count + ' PDF files in one email to ' + $defaultTo)
    } catch {
        Write-Host ('[X] Sending failed: ' + $_.Exception.Message)
    }
    try { $outlook.Session.SendAndReceive($false) } catch {}
    Start-Sleep -Seconds 3
    Remove-Item -LiteralPath $pdfDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Host ('[V] Done: ' + $sent + ' of 1 emails sent.')
}
__ENTRY__
"""


def parse_email_list(raw: str) -> list[str]:
    """
    רשימת כתובות מייל מתוך פרמטר. מפרידים: פסיק, נקודה-פסיק או רווח.
    כתובת לא תקינה מדולגת עם אזהרה; כפילויות נמחקות.
    """
    cleaned = re.sub(r"[\[\]\(\)\"'<>]", " ", str(raw or ""))
    emails: list[str] = []
    for token in re.split(r"[;,\s]+", cleaned):
        token = token.strip()
        if not token:
            continue
        if not _EMAIL_RE.match(token):
            print(f"[!] כתובת מייל לא תקינה דולגה: {token}")
            continue
        if token.lower() not in (e.lower() for e in emails):
            emails.append(token)
    return emails


def kv_get(raw_kv: dict[str, str], key: str) -> str:
    """קריאת פרמטר מקובץ ההגדרות, סלחנית לרווחים בשם (למשל SEND_PICKING_LIST_TO _COOLING)."""
    want = key.replace(" ", "").upper()
    for k, v in raw_kv.items():
        if str(k).replace(" ", "").upper() == want:
            return str(v or "")
    return ""


def _ps_quote(value: Any) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def create_launcher(
    folder: pathlib.Path,
    launcher_name: str,
    title: str,
    default_to: list[str],
    default_name: str,
    rules: list[tuple[str, list[str], str]] | None = None,
    name_filter: str = "",
    strip_prefix: str = "",
    subject_prefix: str = "",
    pdf_single_mail: bool = False,
    single_subject: str = "",
) -> pathlib.Path | None:
    """
    יוצר קובץ bat ששולח כל קובץ בתיקייה במייל נפרד.
    rules: [(סיומת שם קובץ, נמענים, שם להצגה)] - קובץ שמסתיים בסיומת נשלח לנמענים שלה.
    name_filter: ביטוי רגולרי - נשלחים רק קבצים ששמם מתאים (ריק = כולם).
    הנושא: subject_prefix + "נהג <שם> <תאריך>" (משם הקובץ, בלי strip_prefix ובלי הסיומת).
    pdf_single_mail: ממיר את קובצי האקסל ל-PDF בזמן השליחה ושולח את כולם במייל אחד
    לנמענים של ברירת המחדל (נושא: single_subject, או שם התיקייה אם ריק).
    """
    rules = rules or []
    if not default_to and not any(r[1] for r in rules):
        return None
    ps_body = build_ps_body(
        default_to, default_name, rules, name_filter, strip_prefix, subject_prefix,
        pdf_single_mail, single_subject,
    )
    content = (_BAT_HEADER.replace("__TITLE__", title) + ps_body)
    content = content.replace("\r\n", "\n").replace("\n", "\r\n")
    out = folder / launcher_name
    out.write_bytes(content.encode("utf-8"))  # UTF-8 בלי BOM - חובה לקובצי bat
    return out


def build_ps_body(
    default_to: list[str],
    default_name: str,
    rules: list[tuple[str, list[str], str]] | None = None,
    name_filter: str = "",
    strip_prefix: str = "",
    subject_prefix: str = "",
    pdf_single_mail: bool = False,
    single_subject: str = "",
) -> str:
    """סקריפט ה-PowerShell של השליחה (מתחיל בשורה #PS-START)."""
    rules = rules or []
    rules_ps = ",".join(
        f"@{{Suffix={_ps_quote(suf)}; To={_ps_quote('; '.join(to))}; Name={_ps_quote(nm)}}}"
        for suf, to, nm in rules
    )
    ps_body = (
        _PS_BODY.replace("__EXTS__", ",".join(_ps_quote(e) for e in MAIL_ATTACH_EXTENSIONS))
        .replace("__RULES__", rules_ps)
        .replace("__DEFAULT_TO__", _ps_quote("; ".join(default_to)))
        .replace("__DEFAULT_NAME__", _ps_quote(default_name))
        .replace("__FILTER__", _ps_quote(name_filter))
        .replace("__STRIP_PREFIX__", _ps_quote(strip_prefix))
        .replace("__SUBJECT_PREFIX__", _ps_quote(subject_prefix))
        .replace("__SINGLE_SUBJECT__", _ps_quote(single_subject))
        .replace("__ENTRY__", "Send-FolderAsPdf" if pdf_single_mail else "Send-FolderFiles")
    )
    return ps_body


def run_launcher(launcher: pathlib.Path, powershell: str = "powershell") -> tuple[int, int, str]:
    """
    מריץ את השליחה של קובץ bat שנוצר ב-create_launcher, בלי חלון ובלי "pause".
    מחזיר (כמה נשלחו, כמה קבצים, הפלט המלא).
    """
    path = str(launcher.resolve()).replace("'", "''")
    cmd = (
        f"$f='{path}'; $s=[IO.File]::ReadAllText($f,[Text.Encoding]::UTF8); "
        "$i=$s.LastIndexOf('#PS-START'); Invoke-Expression $s.Substring($i)"
    )
    kwargs: dict[str, Any] = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    proc = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
        capture_output=True,
        **kwargs,
    )
    out = (proc.stdout or b"").decode("utf-8", errors="replace")
    out += (proc.stderr or b"").decode("utf-8", errors="replace")
    print(out)
    m = re.search(r"Done: (\d+) of (\d+)", out)
    if m:
        return int(m.group(1)), int(m.group(2)), out
    return 0, 0, out


def _run_powershell(args: list[str], powershell: str) -> str:
    kwargs: dict[str, Any] = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    proc = subprocess.run([powershell, *args], capture_output=True, **kwargs)
    out = (proc.stdout or b"").decode("utf-8", errors="replace")
    out += (proc.stderr or b"").decode("utf-8", errors="replace")
    print(out)
    return out


def run_inline(
    folder: pathlib.Path,
    default_to: list[str],
    default_name: str,
    rules: list[tuple[str, list[str], str]] | None = None,
    name_filter: str = "",
    strip_prefix: str = "",
    subject_prefix: str = "",
    powershell: str = "powershell",
    pdf_single_mail: bool = False,
    single_subject: str = "",
) -> tuple[int, int, str]:
    """
    שולח מיד את הקבצים המתאימים בתיקייה (כל קובץ במייל נפרד), בלי ליצור bat.
    מחזיר (כמה נשלחו, כמה קבצים, הפלט המלא).
    """
    if not default_to and not any(r[1] for r in (rules or [])):
        return 0, 0, "no recipients"
    body = build_ps_body(
        default_to, default_name, rules, name_filter, strip_prefix, subject_prefix,
        pdf_single_mail, single_subject,
    )
    fake_self = str((folder / "_send.bat").resolve()).replace("'", "''")
    script = f"$f = '{fake_self}'\r\n" + body.replace("\r\n", "\n").replace("\n", "\r\n")
    tmp = tempfile.NamedTemporaryFile("wb", suffix=".ps1", delete=False)
    try:
        # UTF-8 עם BOM - כדי ש-PowerShell 5 יקרא נכון את העברית
        tmp.write(script.encode("utf-8-sig"))
        tmp.close()
        out = _run_powershell(
            ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp.name], powershell
        )
    finally:
        try:
            pathlib.Path(tmp.name).unlink()
        except OSError:
            pass
    m = re.search(r"Done: (\d+) of (\d+)", out)
    if m:
        return int(m.group(1)), int(m.group(2)), out
    return 0, 0, out