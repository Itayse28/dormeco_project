#!/usr/bin/env python3
"""
ממשק עריכת מסלולי נהגים (דורמקו) - route_editor_gui.py
=====================================================
נפתח אוטומטית מתוך sort_orders_by_region.py אחרי תכנון המסלולים האוטומטי.

מה אפשר לעשות בחלון:
- לגרור עצירה למעלה/למטה בתוך המסלול של אותו נהג (שינוי סדר).
- לגרור עצירה לעמודה של נהג אחר (העברה בין נהגים).
- "הוסף נהג" - פותח עמודה ריקה שאפשר לגרור אליה עצירות.
- "בטל נהג" - מוחק את הנהג; העצירות שלו עוברות לעמודה "ממתינות לשיבוץ".
- "סדר אוטומטי" - מחשב מחדש את הסדר הקצר ביותר לעצירות של אותו נהג.
- שעות ההגעה, המשקל וזמן המשמרת מתעדכנים מיד אחרי כל שינוי.
- נהג שחורג ממשקל המשאית או משעות המשמרת מסומן באדום (מותר לאשר בכל זאת).
- "אישור והמשך" - הסקריפט ממשיך ומייצר את קובץ האקסל לפי המסלולים שבחלון.
- "ביטול" / סגירת החלון - ההרצה נעצרת ולא נוצרים קבצים.
- מפה בצד שמאל של החלון: כל נהג בצבע משלו (אותו צבע כמו הפס בראש העמודה שלו).
  לחיצה על עצירה ברשימה או במפה מסמנת את המסלול של אותו נהג. (map_panel.py)

המודול לא מכיר את פרטי החישוב: הוא מקבל פונקציית simulate מהסקריפט הראשי.
"""

from __future__ import annotations

import copy
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox
from typing import Any, Callable

MAP_PANEL_ERROR: str | None = None
try:
    from map_panel import MapPanel
except Exception as _exc:  # pragma: no cover - אם קובץ המפה חסר, מוצגת הודעה במקומה
    MapPanel = None  # type: ignore[assignment,misc]
    MAP_PANEL_ERROR = f"{type(_exc).__name__}: {_exc}"

POOL = -1  # מזהה העמודה "ממתינות לשיבוץ"

FONT = "Arial"
F_TITLE = (FONT, 15, "bold")
F_HEAD = (FONT, 11, "bold")
F_TEXT = (FONT, 10)
F_SMALL = (FONT, 9)
F_LIST = (FONT, 9)
F_BTN = (FONT, 10)
F_BTN_BIG = (FONT, 12, "bold")

BG_APP = "#F3F5F8"
BG_OK = "#FFFFFF"
BG_BAD = "#FDE2E2"
BG_POOL = "#FFF1D6"
BG_EMPTY = "#F7F7F7"
BG_DROP = "#FFF9C4"
FG_BAD = "#B3261E"
FG_OK = "#1E6B34"
FG_MUTED = "#555555"

REGION_COLORS = [
    "#D6E4F5",  # צפון מזרח
    "#D9EAD3",  # צפון מערב
    "#EADCF2",  # ירושלים
    "#FCE5CD",  # שרון
    "#DDE7EE",  # גוש דן
    "#FFF2CC",  # דרום
]

DRIVER_COLORS = [
    "#1E88E5", "#E53935", "#43A047", "#FB8C00", "#8E24AA", "#00897B",
    "#F4511E", "#3949AB", "#7CB342", "#6D4C41", "#D81B60", "#00ACC1",
]

MIN_CARD_WIDTH = 255      # רוחב מינימלי רצוי לכרטיס נהג (פיקסלים ב-100% תצוגה)
MIN_CARDS_PER_ROW = 4     # לפחות 4 נהגים בשורה (=> 8 נהגים במסך עם 2 שורות)
CARD_ROWS_ON_SCREEN = 2   # כמה שורות של כרטיסי נהגים נכנסות במסך אחד
MAP_WIDTH_RATIO = 0.22    # רוחב המפה כברירת מחדל (חלק מרוחב החלון)


class RouteEditor:
    def __init__(
        self,
        root: tk.Tk,
        initial_routes: list[list[str]],
        stops_by_key: dict[str, dict[str, Any]],
        simulate: Callable[[list[str], bool], dict[str, Any]],
        cfg: dict[str, Any],
        region_order: list[str],
    ) -> None:
        self.root = root
        self.stops = stops_by_key
        self.simulate = simulate
        self.cfg = cfg
        self.region_order = list(region_order)
        self.max_weight = float(cfg.get("MAX_TRUCK_WEIGHT_KG", 7500))
        self.max_shift_min = float(cfg.get("MAX_DRIVER_SHIFT_HOURS", 12.0)) * 60.0

        self.initial = [list(r) for r in initial_routes]
        self.routes: list[list[str]] = copy.deepcopy(self.initial)
        # מזהה קבוע לכל נהג - כדי שהצבע שלו לא ישתנה כשמוסיפים/מבטלים נהגים
        self.route_ids: list[int] = list(range(len(self.routes)))
        self._next_id = len(self.routes)
        self.selected_id: int | None = None
        self.pool: list[str] = []
        self.depot = tuple(cfg.get("MAIN_WAREHOUSE_COORDS", (31.8315, 34.6465)))
        self.home = tuple(cfg.get("DRIVER_HOME_COORDS", (32.3215, 34.8532)))
        self.map_panel = None
        self._col_frames: dict[int, tk.Frame] = {}
        self.result: list[list[str]] | None = None

        self._sim_cache: dict[tuple[str, ...], dict[str, Any]] = {}
        self._drop_map: dict[str, int] = {}  # שם widget -> מזהה עמודה
        self._listboxes: dict[int, tk.Listbox] = {}
        self._drag: dict[str, Any] | None = None
        self._reselect: tuple[int, int] | None = None
        self._first_render = True

        self._build_window()
        self.render()

    # ------------------------------------------------------------------
    # חישובים
    # ------------------------------------------------------------------
    def _sim(self, keys: list[str]) -> dict[str, Any]:
        t = tuple(keys)
        if t not in self._sim_cache:
            self._sim_cache[t] = self.simulate(list(keys), True)
        return self._sim_cache[t]

    def _problems(self, sim: dict[str, Any]) -> list[str]:
        out = []
        if sim["total_weight"] > self.max_weight + 1e-6:
            out.append(f'חריגת משקל (מקס\' {self.max_weight:,.0f})')
        if sim["total_shift_min"] > self.max_shift_min + 1e-6:
            out.append(f"חריגת שעות (מקס' {self.max_shift_min / 60:g})")
        return out

    def _region_title(self, keys: list[str]) -> str:
        regs = {self.stops[k]["region"] for k in keys}
        ordered = [r for r in self.region_order if r in regs]
        ordered += sorted(r for r in regs if r not in self.region_order)
        return " + ".join(ordered)

    def _color(self, col: int) -> str:
        if col == POOL:
            return "#F6A623"
        return DRIVER_COLORS[self.route_ids[col] % len(DRIVER_COLORS)]

    def _col_of_id(self, rid: int | None) -> int | None:
        if rid is None or rid not in self.route_ids:
            return None
        return self.route_ids.index(rid)

    def _region_color(self, keys: list[str]) -> str:
        if not keys:
            return "#E6E6E6"
        reg = self.stops[keys[0]]["region"]
        if reg in self.region_order:
            idx = self.region_order.index(reg)
            if idx < len(REGION_COLORS):
                return REGION_COLORS[idx]
        return "#E6E6E6"

    def _stop_text(self, key: str, pos: int, sim: dict[str, Any] | None) -> str:
        st = self.stops[key]
        parts = [f"{pos}."]
        if sim is not None:
            sched = sim["schedule_by_stop"].get(key)
            if sched:
                parts.append(sched["arrival_str"])
        parts.append(st.get("effective_dest") or "-")
        cust = st.get("cust_name") or ""
        if cust:
            parts.append(cust)
        parts.append(f'{st.get("weight", 0):,.0f} ק"ג')
        tags = []
        if st.get("needs_cooling"):
            tags.append("קירור")
        if st.get("needs_hazmat"):
            tags.append('חומ"ס')
        if st.get("needs_goldbond"):
            tags.append("גולד בונד")
        if tags:
            parts.append("/".join(tags))
        return " | ".join(parts)

    # ------------------------------------------------------------------
    # בניית החלון
    # ------------------------------------------------------------------
    def _build_window(self) -> None:
        r = self.root
        r.title("דורמקו - עריכת מסלולי נהגים")
        r.configure(bg=BG_APP)
        r.geometry("1400x820")
        try:
            r.state("zoomed")
        except tk.TclError:
            pass
        r.protocol("WM_DELETE_WINDOW", self.on_cancel)

        top = tk.Frame(r, bg=BG_APP, padx=12, pady=8)
        top.pack(side="top", fill="x")
        tk.Label(
            top, text="עריכת מסלולי נהגים", font=F_TITLE, bg=BG_APP
        ).pack(side="right")
        self.info_label = tk.Label(top, text="", font=F_HEAD, bg=BG_APP)
        self.info_label.pack(side="left")

        tk.Label(
            r,
            text=(
                "גררו עצירה עם העכבר כדי לשנות את הסדר שלה או להעביר אותה לנהג אחר. "
                "השעות והמשקלים מתעדכנים אוטומטית. בסיום לחצו על \"אישור והמשך\"."
            ),
            font=F_TEXT,
            bg=BG_APP,
            fg=FG_MUTED,
            anchor="e",
            justify="right",
            padx=12,
        ).pack(side="top", fill="x")

        bottom = tk.Frame(r, bg=BG_APP, padx=12, pady=10)
        bottom.pack(side="bottom", fill="x")
        tk.Button(
            bottom,
            text="אישור והמשך",
            font=F_BTN_BIG,
            bg="#2E7D32",
            fg="white",
            activebackground="#256628",
            activeforeground="white",
            padx=24,
            pady=6,
            command=self.on_approve,
        ).pack(side="right", padx=6)
        tk.Button(
            bottom, text="+ הוסף נהג", font=F_BTN, padx=12, pady=6,
            command=self.add_driver,
        ).pack(side="right", padx=6)
        tk.Button(
            bottom, text="חזרה לתכנון המקורי", font=F_BTN, padx=12, pady=6,
            command=self.reset,
        ).pack(side="right", padx=6)
        tk.Button(
            bottom, text="ביטול", font=F_BTN, padx=12, pady=6,
            command=self.on_cancel,
        ).pack(side="left", padx=6)

        paned = tk.PanedWindow(
            r, orient="horizontal", sashwidth=8, sashrelief="raised", bg="#C9CED6"
        )
        paned.pack(side="top", fill="both", expand=True, padx=8)
        if MapPanel is not None:
            self.map_panel = MapPanel(
                paned, on_stop_click=self.on_map_stop_click, on_show_all=self.show_all
            )
            paned.add(self.map_panel, minsize=250, stretch="always")
        else:
            msg = (
                "המפה לא נטענה.\n\nודאו שהקובץ map_panel.py נמצא באותה תיקייה "
                f"כמו route_editor_gui.py.\n\nפרטי השגיאה:\n{MAP_PANEL_ERROR}"
            )
            print(f"[מפה] {msg}")
            paned.add(
                tk.Label(
                    paned, text=msg, font=F_TEXT, fg=FG_BAD, bg=BG_APP,
                    justify="right", wraplength=380,
                ),
                minsize=250, stretch="always",
            )
        mid = tk.Frame(r, bg=BG_APP)
        paned.add(mid, minsize=400, stretch="always")
        self._paned = paned
        r.after(80, self._place_sash)

        self.canvas = tk.Canvas(
            mid, bg=BG_APP, highlightthickness=0, yscrollincrement=30
        )
        vbar = tk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="left", fill="y")
        self.canvas.pack(side="right", fill="both", expand=True)

        self.board = tk.Frame(self.canvas, bg=BG_APP)
        self.board_win = self.canvas.create_window(0, 0, window=self.board, anchor="nw")
        self.canvas.bind("<Configure>", lambda _e: self._layout())
        r.bind_all("<MouseWheel>", self._on_wheel)

        self.ghost: tk.Toplevel | None = None
        self._cards: list[tuple[int, tk.Frame]] = []
        self._scale = max(r.winfo_fpixels("1i") / 96.0, 1.0)
        self._used_grid_cols = 0
        self._list_lines = 8

    def _place_sash(self) -> None:
        w = self.root.winfo_width()
        if w < 300:
            self.root.after(100, self._place_sash)
            return
        self._paned.sash_place(0, int(w * MAP_WIDTH_RATIO), 0)

    def _layout(self) -> None:
        """מסדר את כרטיסי הנהגים ברשת (כמה שנכנס ברוחב, ואז שורה חדשה)."""
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 50 or ch < 50:
            return
        self.canvas.itemconfigure(self.board_win, width=cw)

        per_row = max(MIN_CARDS_PER_ROW, int(cw // (MIN_CARD_WIDTH * self._scale)))
        for i in range(max(per_row, self._used_grid_cols)):
            if i < per_row:
                self.board.grid_columnconfigure(i, weight=1, uniform="card")
            else:
                self.board.grid_columnconfigure(i, weight=0, uniform="")
        self._used_grid_cols = per_row

        for pos, (_col, frame) in enumerate(self._cards):
            row, c = divmod(pos, per_row)
            # מימין לשמאל: הנהג הראשון בפינה הימנית העליונה
            frame.grid(row=row, column=per_row - 1 - c, sticky="nsew", padx=4, pady=4)

        # גובה הרשימות כך ש-CARD_ROWS_ON_SCREEN שורות של כרטיסים ייכנסו במסך
        if self._cards:
            self.board.update_idletasks()
            fixed = 0
            for col, frame in self._cards:
                lb = self._listboxes.get(col)
                if lb is not None:
                    fixed = max(fixed, frame.winfo_reqheight() - lb.winfo_reqheight())
            line_px = tkfont.Font(root=self.root, font=F_LIST).metrics("linespace") + 1
            card_h = ch / CARD_ROWS_ON_SCREEN - 10 * self._scale
            lines = max(4, int((card_h - fixed) // line_px))
            if lines != self._list_lines:
                self._list_lines = lines
                for lb in self._listboxes.values():
                    lb.configure(height=lines)

        self.board.update_idletasks()
        bh = self.board.winfo_reqheight()
        self.canvas.configure(scrollregion=(0, 0, cw, max(bh, ch)))

    def _on_wheel(self, event: tk.Event) -> None:
        w = event.widget
        if not str(w).startswith(str(self.canvas)):
            return  # המפה מטפלת בגלגלת בעצמה (זום)
        if isinstance(w, tk.Listbox) and w.yview() != (0.0, 1.0):
            return  # גלילה בתוך רשימה ארוכה
        if self.canvas.yview() == (0.0, 1.0):
            return
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    # ------------------------------------------------------------------
    # ציור הלוח
    # ------------------------------------------------------------------
    def render(self) -> None:
        yview = self.canvas.yview()
        for child in self.board.winfo_children():
            child.destroy()
        self._drop_map.clear()
        self._listboxes.clear()
        self._col_frames.clear()
        self._cards = []
        if self._col_of_id(self.selected_id) is None:
            self.selected_id = None

        list_height = self._list_lines

        if self.pool:
            self._build_column(POOL, self.pool, list_height, None)

        drv_num = 0
        for col, keys in enumerate(self.routes):
            if keys:
                drv_num += 1
                self._build_column(col, keys, list_height, drv_num)
            else:
                self._build_column(col, keys, list_height, 0)

        self._update_info()
        self._layout()
        self._draw_map(fit=self._first_render)
        if self._first_render:
            self.canvas.yview_moveto(0.0)
            self._first_render = False
        else:
            self.canvas.yview_moveto(yview[0])

        if self._reselect is not None:
            col, idx = self._reselect
            lb = self._listboxes.get(col)
            if lb is not None and 0 <= idx < lb.size():
                lb.selection_set(idx)
                lb.see(idx)
            self._reselect = None

    def _register_drop(self, widget: tk.Widget, col: int) -> None:
        self._drop_map[str(widget)] = col

    def _build_column(
        self, col: int, keys: list[str], list_height: int, drv_num: int | None
    ) -> None:
        is_pool = col == POOL
        sim = self._sim(keys) if (keys and not is_pool) else None
        problems = self._problems(sim) if sim else []

        if is_pool:
            bg = BG_POOL
        elif not keys:
            bg = BG_EMPTY
        elif problems:
            bg = BG_BAD
        else:
            bg = BG_OK

        frame = tk.Frame(
            self.board, bg=bg, bd=1, relief="solid", padx=5, pady=4,
            highlightthickness=3, highlightbackground=bg, highlightcolor=bg,
        )
        self._cards.append((col, frame))
        self._register_drop(frame, col)
        self._col_frames[col] = frame
        self._style_frame(col)
        stripe = tk.Frame(frame, bg=self._color(col), height=6)
        stripe.pack(side="top", fill="x", pady=(0, 2))
        self._register_drop(stripe, col)

        # כותרת
        if is_pool:
            title = "ממתינות לשיבוץ"
            head_bg = "#F6C26B"
        elif not keys:
            title = "נהג נוסף (ריק)"
            head_bg = "#E6E6E6"
        else:
            title = f"נהג {drv_num}  ({self._region_title(keys)})"
            head_bg = self._region_color(keys)
        head = tk.Label(
            frame, text=title, font=F_HEAD, bg=head_bg, anchor="e", padx=4, pady=2,
            cursor="hand2",
        )
        head.pack(side="top", fill="x")
        self._register_drop(head, col)
        if not is_pool:
            head.bind("<Button-1>", lambda _e, c=col: self.select_route(c))

        # שורות מידע (תמיד 3 שורות, כדי שכל הכרטיסים יהיו באותו גובה)
        info_lines: list[tuple[str, str]] = []
        if is_pool:
            info_lines = [
                (f"{len(keys)} עצירות", FG_BAD),
                ("גררו כל אחת לנהג כלשהו", FG_BAD),
                ("", FG_MUTED),
            ]
        elif not keys:
            info_lines = [("גררו לכאן עצירות", FG_MUTED), ("", FG_MUTED), ("", FG_MUTED)]
        else:
            wh_names = [
                part.split(" (")[0]
                for part in str(sim.get("warehouses_summary", "")).split(" -> ")
                if part
            ]
            info_lines.append(
                (
                    f'{len(keys)} עצירות | {sim["total_weight"]:,.0f} ק"ג | '
                    f'משמרת {sim["total_shift_hhmm"]}',
                    "black",
                )
            )
            info_lines.append(
                (f"חזרה {sim['return_home_str']} | " + ", ".join(wh_names), FG_MUTED)
            )
            if problems:
                info_lines.append((" | ".join(problems), FG_BAD))
            else:
                info_lines.append(("תקין", FG_OK))

        for text, color in info_lines:
            lbl = tk.Label(
                frame, text=text, font=F_SMALL, bg=bg, fg=color,
                anchor="e", justify="right",
            )
            lbl.pack(side="top", fill="x")
            self._register_drop(lbl, col)

        # רשימת העצירות
        list_frame = tk.Frame(frame, bg=bg)
        list_frame.pack(side="top", fill="both", expand=True, pady=(3, 3))
        self._register_drop(list_frame, col)
        lb = tk.Listbox(
            list_frame,
            font=F_LIST,
            width=10,
            height=list_height,
            justify="right",
            activestyle="none",
            exportselection=False,
            selectbackground="#1976D2",
            selectforeground="white",
            bg="white",
        )
        sb = tk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        lb.pack(side="right", fill="both", expand=True)
        for i, key in enumerate(keys, start=1):
            lb.insert("end", self._stop_text(key, i, sim))
        self._listboxes[col] = lb
        self._register_drop(lb, col)

        lb.bind("<ButtonPress-1>", lambda e, c=col: self._on_press(e, c))
        lb.bind("<B1-Motion>", self._on_motion)
        lb.bind("<ButtonRelease-1>", self._on_release)

        # כפתורי עמודה
        btns = tk.Frame(frame, bg=bg)
        btns.pack(side="top", fill="x")
        self._register_drop(btns, col)
        if not is_pool:
            tk.Button(
                btns, text="בטל נהג", font=F_SMALL, fg=FG_BAD, padx=3, pady=0,
                command=lambda c=col: self.remove_driver(c),
            ).pack(side="left", padx=1)
            if len(keys) > 1:
                tk.Button(
                    btns, text="סדר אוטומטי", font=F_SMALL, padx=3, pady=0,
                    command=lambda c=col: self.auto_order(c),
                ).pack(side="left", padx=1)
        tk.Button(
            btns, text="▼", font=F_SMALL, width=2, padx=1, pady=0,
            command=lambda c=col: self.nudge(c, +1),
        ).pack(side="right", padx=1)
        tk.Button(
            btns, text="▲", font=F_SMALL, width=2, padx=1, pady=0,
            command=lambda c=col: self.nudge(c, -1),
        ).pack(side="right", padx=1)

    def _update_info(self) -> None:
        active = [r for r in self.routes if r]
        bad = sum(1 for r in active if self._problems(self._sim(r)))
        total_stops = sum(len(r) for r in active)
        parts = [f"{len(active)} נהגים", f"{total_stops} עצירות"]
        color = FG_OK
        if bad:
            parts.append(f"{bad} בחריגה")
            color = FG_BAD
        if self.pool:
            parts.append(f"{len(self.pool)} ממתינות לשיבוץ")
            color = FG_BAD
        self.info_label.configure(text="   |   ".join(parts), fg=color)

    # ------------------------------------------------------------------
    # גרירה ושחרור
    # ------------------------------------------------------------------
    def _col_list(self, col: int) -> list[str]:
        return self.pool if col == POOL else self.routes[col]

    def _resolve_col(self, widget: tk.Misc | None) -> int | None:
        w = widget
        while w is not None:
            col = self._drop_map.get(str(w))
            if col is not None:
                return col
            w = getattr(w, "master", None)
        return None

    def _on_press(self, event: tk.Event, col: int) -> None:
        lb: tk.Listbox = event.widget
        self._drag = None
        if lb.size() == 0:
            return
        idx = lb.nearest(event.y)
        bbox = lb.bbox(idx)
        if not bbox or event.y > bbox[1] + bbox[3] + 2:
            return
        self._drag = {
            "col": col,
            "idx": idx,
            "x": event.x_root,
            "y": event.y_root,
            "active": False,
            "hover": None,
        }

    def _on_motion(self, event: tk.Event) -> str | None:
        d = self._drag
        if d is None:
            return None
        if not d["active"]:
            if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 6:
                return "break"
            d["active"] = True
            key = self._col_list(d["col"])[d["idx"]]
            st = self.stops[key]
            self.ghost = tk.Toplevel(self.root)
            self.ghost.overrideredirect(True)
            try:
                self.ghost.attributes("-topmost", True)
                self.ghost.attributes("-alpha", 0.9)
            except tk.TclError:
                pass
            tk.Label(
                self.ghost,
                text=f"{st.get('effective_dest', '')}  -  {st.get('cust_name', '')}",
                font=F_TEXT, bg="#1976D2", fg="white", padx=10, pady=4,
            ).pack()

        if self.ghost is not None:
            self.ghost.geometry(f"+{event.x_root + 14}+{event.y_root + 8}")

        target = self.root.winfo_containing(event.x_root, event.y_root)
        col = self._resolve_col(target)
        if col != d["hover"]:
            self._set_hover(d["hover"], False)
            self._set_hover(col, True)
            d["hover"] = col

        self._autoscroll(event)
        return "break"

    def _set_hover(self, col: int | None, on: bool) -> None:
        if col is None:
            return
        lb = self._listboxes.get(col)
        if lb is not None:
            try:
                lb.configure(bg=BG_DROP if on else "white")
            except tk.TclError:
                pass

    def _autoscroll(self, event: tk.Event) -> None:
        cy = self.canvas.winfo_rooty()
        ch = self.canvas.winfo_height()
        if event.y_root < cy + 30:
            self.canvas.yview_scroll(-1, "units")
        elif event.y_root > cy + ch - 30:
            self.canvas.yview_scroll(1, "units")
        d = self._drag
        lb = self._listboxes.get(d["hover"]) if d and d["hover"] is not None else None
        if lb is not None:
            ly = event.y_root - lb.winfo_rooty()
            if ly < 15:
                lb.yview_scroll(-1, "units")
            elif ly > lb.winfo_height() - 15:
                lb.yview_scroll(1, "units")

    def _on_release(self, event: tk.Event) -> str | None:
        d = self._drag
        self._drag = None
        if d is None:
            return None
        if not d["active"]:
            self.select_route(d["col"], d["idx"])
            return None
        if self.ghost is not None:
            self.ghost.destroy()
            self.ghost = None
        self._set_hover(d["hover"], False)

        target = self.root.winfo_containing(event.x_root, event.y_root)
        col = self._resolve_col(target)
        if col is None:
            return "break"

        dst_list = self._col_list(col)
        lb = self._listboxes.get(col)
        if target is lb and lb is not None and lb.size() > 0:
            y = event.y_root - lb.winfo_rooty()
            near = lb.nearest(y)
            bbox = lb.bbox(near)
            idx = near + 1 if (bbox and y > bbox[1] + bbox[3] / 2) else near
        else:
            idx = len(dst_list)

        self.move(d["col"], d["idx"], col, idx)
        return "break"

    # ------------------------------------------------------------------
    # פעולות
    # ------------------------------------------------------------------
    def move(self, src_col: int, src_idx: int, dst_col: int, dst_idx: int) -> None:
        src = self._col_list(src_col)
        dst = self._col_list(dst_col)
        if not (0 <= src_idx < len(src)):
            return
        key = src.pop(src_idx)
        if src is dst and src_idx < dst_idx:
            dst_idx -= 1
        dst_idx = max(0, min(dst_idx, len(dst)))
        dst.insert(dst_idx, key)
        self._reselect = (dst_col, dst_idx)
        if dst_col != POOL:
            self.selected_id = self.route_ids[dst_col]
        self.render()

    def nudge(self, col: int, delta: int) -> None:
        lb = self._listboxes.get(col)
        if lb is None or not lb.curselection():
            return
        idx = lb.curselection()[0]
        lst = self._col_list(col)
        new_idx = idx + delta
        if not (0 <= new_idx < len(lst)):
            return
        lst[idx], lst[new_idx] = lst[new_idx], lst[idx]
        self._reselect = (col, new_idx)
        self.render()

    def auto_order(self, col: int) -> None:
        keys = self.routes[col]
        if len(keys) < 2:
            return
        sim = self.simulate(list(keys), False)
        self.routes[col] = [s["stop_key"] for s in sim["ordered_stops"]]
        self.render()

    def add_driver(self) -> None:
        self.routes.append([])
        self.route_ids.append(self._next_id)
        self._next_id += 1
        self.render()
        self.canvas.yview_moveto(1.0)

    def remove_driver(self, col: int) -> None:
        keys = self.routes[col]
        if keys:
            ok = messagebox.askyesno(
                "בטל נהג",
                f"לנהג יש {len(keys)} עצירות.\n"
                "העצירות יעברו לעמודה \"ממתינות לשיבוץ\" ותצטרכו לגרור אותן לנהגים אחרים.\n\n"
                "לבטל את הנהג?",
                parent=self.root,
            )
            if not ok:
                return
            self.pool.extend(keys)
        del self.routes[col]
        del self.route_ids[col]
        self.render()

    def reset(self) -> None:
        ok = messagebox.askyesno(
            "חזרה לתכנון המקורי",
            "כל השינויים שביצעתם יימחקו. להמשיך?",
            parent=self.root,
        )
        if not ok:
            return
        self.routes = copy.deepcopy(self.initial)
        self.route_ids = list(range(len(self.routes)))
        self._next_id = len(self.routes)
        self.selected_id = None
        self.pool = []
        self._first_render = True
        self.render()

    # ------------------------------------------------------------------
    # מפה ובחירת נהג
    # ------------------------------------------------------------------
    def _style_frame(self, col: int) -> None:
        frame = self._col_frames.get(col)
        if frame is None:
            return
        keys = self._col_list(col)
        is_sel = col != POOL and self.selected_id == self.route_ids[col]
        bad = bool(keys) and col != POOL and bool(self._problems(self._sim(keys)))
        if is_sel:
            color = self._color(col)
        elif bad:
            color = FG_BAD
        else:
            color = frame.cget("bg")
        frame.configure(highlightbackground=color, highlightcolor=color)

    def _draw_map(self, fit: bool = False) -> None:
        if self.map_panel is None:
            return
        routes_data = []
        for col, keys in enumerate(self.routes):
            if not keys:
                continue
            routes_data.append(
                {
                    "id": self.route_ids[col],
                    "color": self._color(col),
                    "stops": [
                        (*self.stops[k]["coords"], i, k)
                        for i, k in enumerate(keys, start=1)
                    ],
                }
            )
        pool_data = [(*self.stops[k]["coords"], i, k) for i, k in enumerate(self.pool, 1)]
        self.map_panel.draw(
            routes_data, pool_data, self.depot, self.home, self.selected_id, fit=fit
        )
        self._map_routes = routes_data
        self._map_pool = pool_data

    def select_route(self, col: int, idx: int | None = None) -> None:
        """מסמן נהג (ואם צוין - עצירה מסוימת שלו) ברשימות ובמפה."""
        if col == POOL:
            self.selected_id = None
        else:
            self.selected_id = self.route_ids[col]
        for c in list(self._col_frames):
            self._style_frame(c)
        self._draw_map()
        if idx is not None and self.map_panel is not None:
            keys = self._col_list(col)
            if 0 <= idx < len(keys):
                lat, lon = self.stops[keys[idx]]["coords"]
                self.map_panel.focus(lat, lon)

    def on_map_stop_click(self, key: str) -> None:
        """לחיצה על עצירה במפה: מסמנת אותה ברשימה של הנהג ומגלגלת אליה."""
        col: int | None = None
        if key in self.pool:
            col, idx = POOL, self.pool.index(key)
        else:
            for c, keys in enumerate(self.routes):
                if key in keys:
                    col, idx = c, keys.index(key)
                    break
        if col is None:
            return
        for lb in self._listboxes.values():
            lb.selection_clear(0, "end")
        lb = self._listboxes.get(col)
        if lb is not None:
            lb.selection_set(idx)
            lb.see(idx)
        self._scroll_to_column(col)
        self.select_route(col)

    def _scroll_to_column(self, col: int) -> None:
        frame = self._col_frames.get(col)
        if frame is None:
            return
        self.canvas.update_idletasks()
        region = self.canvas.cget("scrollregion").split()
        total = float(region[3]) if len(region) == 4 else 0.0
        view_h = self.canvas.winfo_height()
        if total <= view_h:
            return
        fy = frame.winfo_y()
        top, bottom = self.canvas.yview()
        if top * total <= fy and fy + frame.winfo_height() <= bottom * total:
            return  # הכרטיס כבר גלוי
        self.canvas.yview_moveto(max(0.0, min((fy - 8) / total, 1.0)))

    def show_all(self) -> None:
        self.selected_id = None
        for c in list(self._col_frames):
            self._style_frame(c)
        self._draw_map(fit=True)

    def on_approve(self) -> None:
        if self.pool:
            messagebox.showwarning(
                "יש עצירות שלא שובצו",
                f"{len(self.pool)} עצירות עדיין ממתינות לשיבוץ.\n"
                "גררו כל אחת מהן לנהג לפני האישור.",
                parent=self.root,
            )
            return
        active = [r for r in self.routes if r]
        if not active:
            messagebox.showerror("אין מסלולים", "לא נשארו נהגים עם עצירות.", parent=self.root)
            return
        bad = [i for i, r in enumerate(active, start=1) if self._problems(self._sim(r))]
        if bad:
            names = ", ".join(f"נהג {i}" for i in bad)
            ok = messagebox.askyesno(
                "יש חריגות",
                f"הנהגים הבאים חורגים ממשקל המשאית או משעות המשמרת:\n{names}\n\n"
                "להמשיך בכל זאת?",
                icon="warning",
                parent=self.root,
            )
            if not ok:
                return
        self.result = [list(r) for r in active]
        self.root.destroy()

    def on_cancel(self) -> None:
        ok = messagebox.askyesno(
            "ביטול",
            "לבטל את ההרצה?\nקבצי האקסל לא ייווצרו.",
            icon="warning",
            parent=self.root,
        )
        if ok:
            self.result = None
            self.root.destroy()


def _enable_windows_dpi_awareness() -> None:
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
    except Exception:
        pass


def run_route_editor(
    initial_routes: list[list[str]],
    stops_by_key: dict[str, dict[str, Any]],
    simulate: Callable[[list[str], bool], dict[str, Any]],
    cfg: dict[str, Any],
    region_order: list[str],
) -> list[list[str]] | None:
    """
    פותח את חלון העריכה וממתין עד שהמשתמש יאשר או יבטל.
    מחזיר רשימת מסלולים (כל מסלול = רשימת stop_key לפי הסדר), או None אם בוטל.
    """
    _enable_windows_dpi_awareness()
    root = tk.Tk()
    editor = RouteEditor(root, initial_routes, stops_by_key, simulate, cfg, region_order)
    root.lift()
    try:
        root.attributes("-topmost", True)
        root.after(600, lambda: root.attributes("-topmost", False))
    except tk.TclError:
        pass
    root.focus_force()
    root.mainloop()
    return editor.result