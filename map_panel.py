#!/usr/bin/env python3
"""
פאנל מפה לחלון עריכת המסלולים (דורמקו) - map_panel.py
====================================================
- החלון נפתח מיד עם מפה סכמטית (נקודות וקווים בלבד, בלי אינטרנט).
- ברקע נבדק החיבור לשרתי המפות. אם יש חיבור והספרייה tkintermapview מותקנת,
  המפה מתחלפת אוטומטית למפה אמיתית עם זום בגלגלת וגרירה להזזה.
- בתחתית המפה מוצגת שורת מצב שמסבירה איזו מפה מוצגת ולמה.
התקנה חד-פעמית של המפה האמיתית:  python -m pip install tkintermapview
בדיקת תקלות:                     python check_map.py
"""

from __future__ import annotations

import math
import threading
import tkinter as tk
import types
from typing import Any, Callable

MAP_IMPORT_ERROR: str | None = None
try:
    import requests
    import tkintermapview  # type: ignore
    from tkintermapview import map_widget as _tmv_widget  # type: ignore
except Exception as _exc:  # pragma: no cover - כל תקלה בספרייה => מפה סכמטית
    tkintermapview = None
    _tmv_widget = None
    MAP_IMPORT_ERROR = f"{type(_exc).__name__}: {_exc}"

# שרתי מפות לניסיון לפי הסדר. הראשון שעונה - נבחר.
TILE_SERVERS: list[tuple[str, str, int]] = [
    ("OpenStreetMap", "https://tile.openstreetmap.org/{z}/{x}/{y}.png", 19),
    ("CARTO", "https://a.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}.png", 20),
    ("OpenStreetMap DE", "https://tile.openstreetmap.de/{z}/{x}/{y}.png", 18),
]
# אריח בדיקה (זום 8) שמכסה את מרכז הארץ
TEST_TILE = {"z": 8, "x": 152, "y": 104}
USER_AGENT = "DormacoRoutePlanner/1.0 (driver route editor; python-requests)"


def _tile_get(url: str, *args: Any, **kwargs: Any):  # noqa: ANN202
    """בקשת אריח עם User-Agent מזהה ו-timeout. שרתי OSM חוסמים את ברירת המחדל של הספרייה."""
    headers = dict(kwargs.pop("headers", None) or {})
    headers["User-Agent"] = USER_AGENT
    kwargs.setdefault("timeout", 10)
    resp = requests.get(url, *args, headers=headers, **kwargs)
    if resp.status_code != 200:
        raise requests.exceptions.ConnectionError(f"HTTP {resp.status_code}")
    return resp


if _tmv_widget is not None:
    # הספרייה משתמשת ב-requests.get בלי timeout ועם User-Agent חסום - מחליפים
    _tmv_widget.requests = types.SimpleNamespace(
        get=_tile_get, exceptions=requests.exceptions
    )


def probe_tile_servers(timeout: float = 5.0) -> tuple[tuple[str, str, int] | None, list[str]]:
    """בודק אילו שרתי מפות זמינים. מחזיר (השרת הראשון שעובד, רשימת שגיאות)."""
    errors: list[str] = []
    if tkintermapview is None:
        return None, [f"tkintermapview: {MAP_IMPORT_ERROR}"]
    for name, url, max_zoom in TILE_SERVERS:
        test_url = (
            url.replace("{z}", str(TEST_TILE["z"]))
            .replace("{x}", str(TEST_TILE["x"]))
            .replace("{y}", str(TEST_TILE["y"]))
        )
        try:
            resp = requests.get(
                test_url, headers={"User-Agent": USER_AGENT}, timeout=timeout
            )
            ctype = resp.headers.get("Content-Type", "")
            if resp.status_code == 200 and ctype.startswith("image"):
                return (name, url, max_zoom), errors
            errors.append(f"{name}: HTTP {resp.status_code} ({ctype})")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{name}: {type(exc).__name__}: {str(exc)[:160]}")
    return None, errors

FONT = "Arial"

# מבנה הנתונים שהפאנל מקבל:
# routes = [ {"id": int, "color": "#hex", "stops": [(lat, lon, num, key), ...]}, ... ]
# pool   = [ (lat, lon, num, key), ... ]          (עצירות שממתינות לשיבוץ)
# depot  = (lat, lon)   מחסנים באשדוד - נקודת היציאה של כל המסלולים
# home   = (lat, lon)


class MapPanel(tk.Frame):
    def __init__(
        self,
        master: tk.Misc,
        on_stop_click: Callable[[str], None],
        on_show_all: Callable[[], None],
    ) -> None:
        super().__init__(master, bg="#F3F5F8")
        self.on_stop_click = on_stop_click
        self._last: tuple | None = None

        bar = tk.Frame(self, bg="#F3F5F8")
        bar.pack(side="top", fill="x", pady=(0, 4))
        tk.Button(
            bar, text="הצג את כל המסלולים", font=(FONT, 10),
            command=on_show_all,
        ).pack(side="left", padx=4)
        tk.Label(
            bar,
            text="לחיצה על עצירה ברשימה או במפה מסמנת את המסלול של הנהג",
            font=(FONT, 9), fg="#555555", bg="#F3F5F8", anchor="e",
        ).pack(side="right", padx=4)

        self.status = tk.Label(
            self, text="", font=(FONT, 9), fg="#8A5A00", bg="#F3F5F8",
            anchor="e", justify="right", wraplength=500,
        )
        self.status.pack(side="bottom", fill="x")

        self.backend: _BaseBackend = _SchematicBackend(self, self._clicked)
        self.backend.widget.pack(side="top", fill="both", expand=True)
        self._draw_args: tuple | None = None

        if tkintermapview is None:
            self._set_status(
                "מפה סכמטית - ספריית המפות לא נטענה. "
                "התקינו: python -m pip install tkintermapview  "
                f"(פרטים: {MAP_IMPORT_ERROR})",
                warn=True,
            )
        else:
            self._set_status("בודק חיבור לשרת המפות...", warn=False)
            self._probe_result: tuple | None = None
            threading.Thread(target=self._probe_worker, daemon=True).start()
            self.after(200, self._poll_probe)

    def _set_status(self, text: str, warn: bool) -> None:
        self.status.configure(text=text, fg="#B3261E" if warn else "#1E6B34")
        print(f"[מפה] {text}")

    def _probe_worker(self) -> None:
        self._probe_result = probe_tile_servers()

    def _poll_probe(self) -> None:
        if self._probe_result is None:
            self.after(200, self._poll_probe)
            return
        server, errors = self._probe_result
        if server is None:
            self._set_status(
                "מפה סכמטית - אין גישה לשרתי המפות. "
                + (errors[0] if errors else "")
                + "   להרצת בדיקה מלאה: python check_map.py",
                warn=True,
            )
            return
        try:
            new_backend = _OsmBackend(self, self._clicked, server)
        except Exception as exc:  # noqa: BLE001
            self._set_status(f"מפה סכמטית - שגיאה ביצירת המפה: {exc}", warn=True)
            return
        self.backend.widget.destroy()
        self.backend = new_backend
        self.backend.widget.pack(side="top", fill="both", expand=True)
        self._set_status(f"מפה: {server[0]}", warn=False)
        if self._draw_args is not None:
            routes, pool, depot, home, hl = self._draw_args
            self.backend.draw(routes, pool, depot, home, hl)
            self.fit(routes, pool, depot, hl)

    def _clicked(self, key: str) -> None:
        self.on_stop_click(key)

    def draw(
        self,
        routes: list[dict[str, Any]],
        pool: list[tuple[float, float, int, str]],
        depot: tuple[float, float],
        home: tuple[float, float],
        highlight_id: int | None,
        fit: bool = False,
    ) -> None:
        self._draw_args = (routes, pool, depot, home, highlight_id)
        self.backend.draw(routes, pool, depot, home, highlight_id)
        if fit:
            self.fit(routes, pool, depot, highlight_id)

    def fit(
        self,
        routes: list[dict[str, Any]],
        pool: list[tuple[float, float, int, str]],
        depot: tuple[float, float],
        highlight_id: int | None = None,
    ) -> None:
        pts: list[tuple[float, float]] = []
        for r in routes:
            if highlight_id is None or r["id"] == highlight_id:
                pts += [(s[0], s[1]) for s in r["stops"]]
        if highlight_id is None:
            pts += [(s[0], s[1]) for s in pool]
        pts.append(depot)
        self.backend.fit(pts)

    def focus(self, lat: float, lon: float) -> None:
        self.backend.focus(lat, lon)


class _BaseBackend:
    widget: tk.Widget

    def draw(self, routes, pool, depot, home, highlight_id) -> None:  # noqa: ANN001
        raise NotImplementedError

    def fit(self, pts: list[tuple[float, float]]) -> None:
        raise NotImplementedError

    def focus(self, lat: float, lon: float) -> None:
        raise NotImplementedError


def _lighten(hex_color: str, amount: float = 0.55) -> str:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    r = int(r + (255 - r) * amount)
    g = int(g + (255 - g) * amount)
    b = int(b + (255 - b) * amount)
    return f"#{r:02x}{g:02x}{b:02x}"


# ------------------------------------------------------------------------------
# מפת OpenStreetMap אמיתית
# ------------------------------------------------------------------------------
class _OsmBackend(_BaseBackend):
    def __init__(
        self,
        master: tk.Misc,
        on_click: Callable[[str], None],
        server: tuple[str, str, int],
    ) -> None:
        self.on_click = on_click
        self.widget = tkintermapview.TkinterMapView(master, corner_radius=0)
        self.map = self.widget
        self.map.set_tile_server(server[1], max_zoom=server[2])
        self.map.set_position(31.9, 34.9)
        self.map.set_zoom(8)

    def draw(self, routes, pool, depot, home, highlight_id) -> None:  # noqa: ANN001
        m = self.map
        m.delete_all_path()
        m.delete_all_marker()

        def path_for(r: dict[str, Any], width: int, color: str) -> None:
            pts = [depot] + [(s[0], s[1]) for s in r["stops"]]
            if len(pts) >= 2:
                m.set_path(pts, color=color, width=width)

        # קודם המסלולים הלא-מסומנים, ואז המסומן מעליהם
        for r in routes:
            if highlight_id is not None and r["id"] != highlight_id:
                path_for(r, 2, _lighten(r["color"]))
        for r in routes:
            if highlight_id is None or r["id"] == highlight_id:
                path_for(r, 6 if highlight_id is not None else 3, r["color"])

        for r in routes:
            dim = highlight_id is not None and r["id"] != highlight_id
            col = _lighten(r["color"]) if dim else r["color"]
            for lat, lon, num, key in r["stops"]:
                m.set_marker(
                    lat, lon, text=str(num),
                    text_color="#999999" if dim else "#111111",
                    font=(FONT, 10, "bold"),
                    marker_color_circle="white",
                    marker_color_outside=col,
                    command=lambda mk: self.on_click(mk.data),
                    data=key,
                )
        for lat, lon, _num, key in pool:
            m.set_marker(
                lat, lon, text="?", text_color="#B3261E", font=(FONT, 11, "bold"),
                marker_color_circle="white", marker_color_outside="#F6A623",
                command=lambda mk: self.on_click(mk.data), data=key,
            )
        m.set_marker(
            depot[0], depot[1], text="מחסנים", text_color="#000000",
            font=(FONT, 10, "bold"),
            marker_color_circle="#000000", marker_color_outside="#444444",
        )
        m.set_marker(
            home[0], home[1], text="בית", text_color="#000000",
            font=(FONT, 10, "bold"),
            marker_color_circle="#000000", marker_color_outside="#888888",
        )

    def fit(self, pts: list[tuple[float, float]]) -> None:
        if not pts:
            return
        lat_max = max(p[0] for p in pts) + 0.04
        lat_min = min(p[0] for p in pts) - 0.04
        lon_min = min(p[1] for p in pts) - 0.04
        lon_max = max(p[1] for p in pts) + 0.04
        self.map.after(
            150,
            lambda: self.map.fit_bounding_box((lat_max, lon_min), (lat_min, lon_max)),
        )

    def focus(self, lat: float, lon: float) -> None:
        self.map.set_position(lat, lon)


# ------------------------------------------------------------------------------
# מפה סכמטית (בלי אינטרנט ובלי ספריות נוספות)
# ------------------------------------------------------------------------------
class _SchematicBackend(_BaseBackend):
    PAD = 30

    def __init__(self, master: tk.Misc, on_click: Callable[[str], None]) -> None:
        self.on_click = on_click
        self.widget = tk.Canvas(master, bg="#EAF2F8", highlightthickness=0)
        self.c = self.widget
        self._args: tuple | None = None
        self._bounds: tuple[float, float, float, float] | None = None
        self.c.bind("<Configure>", lambda _e: self._redraw())

    def draw(self, routes, pool, depot, home, highlight_id) -> None:  # noqa: ANN001
        self._args = (routes, pool, depot, home, highlight_id)
        if self._bounds is None:
            pts = [(s[0], s[1]) for r in routes for s in r["stops"]]
            pts += [(s[0], s[1]) for s in pool] + [depot]
            self._set_bounds(pts)
        self._redraw()

    def _set_bounds(self, pts: list[tuple[float, float]]) -> None:
        if not pts:
            return
        lat_min = min(p[0] for p in pts) - 0.03
        lat_max = max(p[0] for p in pts) + 0.03
        lon_min = min(p[1] for p in pts) - 0.03
        lon_max = max(p[1] for p in pts) + 0.03
        self._bounds = (lat_min, lat_max, lon_min, lon_max)

    def _proj(self, lat: float, lon: float) -> tuple[float, float]:
        lat_min, lat_max, lon_min, lon_max = self._bounds  # type: ignore[misc]
        w = max(self.c.winfo_width() - 2 * self.PAD, 10)
        h = max(self.c.winfo_height() - 2 * self.PAD, 10)
        k = math.cos(math.radians((lat_min + lat_max) / 2))
        span_x = (lon_max - lon_min) * k
        span_y = lat_max - lat_min
        scale = min(w / max(span_x, 1e-6), h / max(span_y, 1e-6))
        off_x = self.PAD + (w - span_x * scale) / 2
        off_y = self.PAD + (h - span_y * scale) / 2
        x = off_x + (lon - lon_min) * k * scale
        y = off_y + (lat_max - lat) * scale
        return x, y

    def _redraw(self) -> None:
        c = self.c
        c.delete("all")
        if self._args is None or self._bounds is None:
            return
        routes, pool, depot, home, hl = self._args

        def line(r: dict[str, Any], width: int, color: str) -> None:
            pts = [depot] + [(s[0], s[1]) for s in r["stops"]]
            if len(pts) < 2:
                return
            flat: list[float] = []
            for lat, lon in pts:
                flat += list(self._proj(lat, lon))
            c.create_line(*flat, fill=color, width=width, capstyle="round", joinstyle="round")

        for r in routes:
            if hl is not None and r["id"] != hl:
                line(r, 2, _lighten(r["color"]))
        for r in routes:
            if hl is None or r["id"] == hl:
                line(r, 5 if hl is not None else 3, r["color"])

        def dot(lat: float, lon: float, text: str, fill: str, fg: str, key: str | None, rad: int = 10) -> None:
            x, y = self._proj(lat, lon)
            tag = f"k{id(key)}_{x:.0f}_{y:.0f}"
            c.create_oval(x - rad, y - rad, x + rad, y + rad, fill=fill, outline="white", width=2, tags=(tag,))
            c.create_text(x, y, text=text, fill=fg, font=(FONT, 9, "bold"), tags=(tag,))
            if key is not None:
                c.tag_bind(tag, "<Button-1>", lambda _e, k=key: self.on_click(k))
                c.tag_bind(tag, "<Enter>", lambda _e: c.configure(cursor="hand2"))
                c.tag_bind(tag, "<Leave>", lambda _e: c.configure(cursor=""))

        order = sorted(routes, key=lambda r: 1 if (hl is not None and r["id"] == hl) else 0)
        for r in order:
            dim = hl is not None and r["id"] != hl
            fill = _lighten(r["color"]) if dim else r["color"]
            for lat, lon, num, key in r["stops"]:
                dot(lat, lon, str(num), fill, "white", key)
        for lat, lon, _n, key in pool:
            dot(lat, lon, "?", "#F6A623", "black", key)
        dot(depot[0], depot[1], "מ", "#222222", "white", None, 12)
        x, y = self._proj(*depot)
        c.create_text(x, y + 20, text="מחסנים אשדוד", font=(FONT, 9, "bold"))

    def fit(self, pts: list[tuple[float, float]]) -> None:
        self._set_bounds(pts)
        self._redraw()

    def focus(self, lat: float, lon: float) -> None:
        pass  # במפה הסכמטית כל הנקודות תמיד מוצגות