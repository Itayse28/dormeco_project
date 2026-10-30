#!/usr/bin/env python3
"""
בדיקת תקלות במפה (דורמקו) - check_map.py
=======================================
הרצה:  python check_map.py   (או לחיצה כפולה על check_map.bat)
הסקריפט בודק שלב-שלב למה המפה לא מוצגת ומדפיס מה לעשות.
"""

from __future__ import annotations

import pathlib
import platform
import sys

OK = "[V]"
BAD = "[X]"


def main() -> None:
    here = pathlib.Path(__file__).resolve().parent
    print("=" * 70)
    print("בדיקת מפה - דורמקו")
    print("=" * 70)
    print(f"Python:  {sys.executable}")
    print(f"גרסה:    {platform.python_version()}  |  {platform.platform()}")
    print(f"תיקייה:  {here}")
    print()

    # 1. קבצי הפרויקט
    for name in ("map_panel.py", "route_editor_gui.py", "sort_orders_by_region.py"):
        exists = (here / name).exists()
        print(f"{OK if exists else BAD} הקובץ {name} {'נמצא' if exists else 'חסר בתיקייה!'}")
    print()

    # 2. tkinter
    try:
        import tkinter

        print(f"{OK} tkinter זמין (Tk {tkinter.TkVersion})")
    except Exception as exc:  # noqa: BLE001
        print(f"{BAD} tkinter לא זמין: {exc}")
        print("    התקינו מחדש את Python מ-python.org וסמנו את האפשרות tcl/tk.")
        return

    # 3. tkintermapview
    try:
        import tkintermapview  # noqa: F401
        from importlib.metadata import version

        print(f"{OK} tkintermapview מותקן (גרסה {version('tkintermapview')})")
    except Exception as exc:  # noqa: BLE001
        print(f"{BAD} tkintermapview לא נטען: {type(exc).__name__}: {exc}")
        print("    הריצו בדיוק את הפקודה הזו (היא מתקינה לאותו Python שמריץ את הפרויקט):")
        print(f'    "{sys.executable}" -m pip install tkintermapview')
        print()
        print("    הערה: אם התקנתם עם 'pip install' רגיל, ייתכן שזה הותקן ל-Python אחר.")
        return
    print()

    # 4. map_panel + חיבור לשרתי מפות
    sys.path.insert(0, str(here))
    try:
        import map_panel
    except Exception as exc:  # noqa: BLE001
        print(f"{BAD} טעינת map_panel.py נכשלה: {type(exc).__name__}: {exc}")
        return

    print("בודק חיבור לשרתי המפות (עד כחצי דקה)...")
    server, errors = map_panel.probe_tile_servers(timeout=8)
    for e in errors:
        print(f"{BAD} {e}")
    if server:
        print(f"{OK} שרת מפות זמין: {server[0]}")
        print()
        print("הכל תקין - המפה האמיתית אמורה להופיע בחלון העריכה.")
        print("פותח חלון בדיקה עם מפה... (סגרו אותו כדי לסיים)")
        root = tkinter.Tk()
        root.title("בדיקת מפה")
        root.geometry("800x600")
        m = map_panel.tkintermapview.TkinterMapView(root, corner_radius=0)
        m.set_tile_server(server[1], max_zoom=server[2])
        m.pack(fill="both", expand=True)
        m.set_position(31.83, 34.65)
        m.set_zoom(10)
        root.mainloop()
        return

    print()
    print("אין גישה לאף שרת מפות. סיבות נפוצות:")
    print(" - אין חיבור לאינטרנט במחשב.")
    print(" - רשת ארגונית / חומת אש שחוסמת את השרתים.")
    joined = " ".join(errors)
    if "SSL" in joined or "CERTIFICATE" in joined.upper():
        print(" - שגיאת אישור אבטחה (SSL) - נפוץ ברשתות ארגוניות. נסו:")
        print(f'     "{sys.executable}" -m pip install pip-system-certs')
        print("   ואז הריצו שוב את הבדיקה.")
    if "ProxyError" in joined:
        print(" - יש פרוקסי ארגוני. בקשו ממנהל הרשת לאפשר גישה ל-tile.openstreetmap.org")
    print()
    print("בכל מקרה, חלון העריכה ימשיך לעבוד עם המפה הסכמטית.")


if __name__ == "__main__":
    try:
        main()
    finally:
        try:
            input("\nלחצו Enter לסגירה...")
        except EOFError:
            pass