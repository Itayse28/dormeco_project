#!/usr/bin/env python3
"""
הרצת כל התהליך בלי חלון CMD (דורמקו) - run_all.pyw
=================================================
מופעל מתוך run_all.bat (שנסגר מיד) עם pythonw, ולכן אין חלון שחור.
שלבים:
  1. sort_orders_by_region.py  - מיון + חלון עריכת המסלולים
     (אם הקבצים נוצרו מתוך החלון, שלבים 2 ו-3 כבר בוצעו ולא מורצים שוב)
  2. generate_regional_picking_lists.py - רשימות ליקוט
  3. generate_driver_delivery_reports.py - דוחות הובלה
כל הפלט נשמר בקובץ logs\run_log.txt בתיקיית הפרויקט. שגיאה מוצגת בחלון הודעה.
"""

from __future__ import annotations

import datetime
import os
import pathlib
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent  # תיקיית הפרויקט (הקוד נמצא ב-src)
LOG = ROOT / "logs" / "run_log.txt"
CREATE_NO_WINDOW = 0x08000000


def show_error(text: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror("Dormaco", text, parent=root)
        root.destroy()
    except Exception:  # noqa: BLE001
        pass


def run_step(script: str, extra: list[str], log) -> int:  # noqa: ANN001
    exe = sys.executable  # pythonw.exe - בלי חלון
    cmd = [exe, str(HERE / script), *sys.argv[1:], *extra]
    log.write(f"\n===== {script} =====\n")
    log.flush()
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    kwargs = {"creationflags": CREATE_NO_WINDOW} if sys.platform == "win32" else {}
    proc = subprocess.run(
        cmd, stdout=log, stderr=subprocess.STDOUT, cwd=str(ROOT), env=env, **kwargs
    )
    log.flush()
    return proc.returncode


def main() -> int:
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG, "w", encoding="utf-8") as log:
        log.write(f"===== {datetime.datetime.now():%d-%m-%Y %H:%M:%S} =====\n")
        log.write(f"args: {sys.argv[1:]}\n")

        rc = run_step("sort_orders_by_region.py", ["--no-open"], log)
        if rc == 3:  # כל הקבצים נוצרו מתוך חלון העריכה
            return 0
        if rc == 2:  # בוטל בחלון
            return 2
        if rc != 0:
            show_error(
                "אירעה שגיאה בשלב מיון ההזמנות.\n"
                f"הפרטים בקובץ:\n{LOG}"
            )
            return 1

        rc = run_step("generate_regional_picking_lists.py", ["--no-open"], log)
        if rc not in (0, 4):
            show_error(f"אירעה שגיאה ביצירת רשימות הליקוט.\nהפרטים בקובץ:\n{LOG}")
            return 1
        picking_partial = rc == 4

        rc = run_step("generate_driver_delivery_reports.py", [], log)
        if rc == 4 or picking_partial:
            show_error(
                "חלק מהקבצים לא נשמרו כי הם פתוחים באקסל.\n"
                f"סגרו אותם והריצו שוב. הפרטים בקובץ:\n{LOG}"
            )
            return 1
        if rc != 0:
            show_error(f"אירעה שגיאה ביצירת דוחות ההובלה.\nהפרטים בקובץ:\n{LOG}")
            return 1
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        show_error(f"שגיאה בהרצה: {exc}\nהפרטים בקובץ:\n{LOG}")
        sys.exit(1)
