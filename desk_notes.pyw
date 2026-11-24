"""桌面右侧备忘。

黄点贴在屏幕右缘，点一下把清单拉出来。左侧圆圈切换：待办 → 正在做 → 完成；右边垃圾桶只归档，
连写下时间和归档时间一起记在本机 CSV 里，不真删。

CSV：本目录 data\\notes.csv（记事本 / Excel / WPS 都能直接打开）
"""

from __future__ import annotations

import csv
import ctypes
import os
import sqlite3
import sys
import traceback
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
CSV_PATH = DATA_DIR / "notes.csv"
LEGACY_XLSX_PATH = DATA_DIR / "notes.xlsx"
LEGACY_DB_PATH = DATA_DIR / "notes.db"
LOG_PATH = DATA_DIR / "error.log"

HEADERS = ("ID", "内容", "状态", "写下时间", "更新时间", "归档时间")
STATUS_TODO = "todo"
STATUS_DOING = "doing"
STATUS_DONE = "done"
STATUS_CYCLE = (STATUS_TODO, STATUS_DOING, STATUS_DONE)
STATUS_LABEL = {
    STATUS_TODO: "待办",
    STATUS_DOING: "正在做",
    STATUS_DONE: "完成",
}

DOT = 34
PANEL_W = 400
PANEL_H = 600
ANIM_MS = 200

YELLOW = "#FFD000"
YELLOW_HOVER = "#FFE14A"
YELLOW_EDGE = "#E0AC00"
CHEVRON = "#3A2E00"
PAPER = "#FFFDF8"
INK = "#2C261C"
MUTED = "#8E8578"
LINE = "#F0E6D4"
ROW_HOVER = "#FFF6DE"
DOING_ROW = "#FFF8E3"
DONE_FG = "#A89B8A"
DOING_FG = "#8A6A12"
INPUT_BG = "#FFF3C9"
DANGER = "#E15A48"
TOAST_BG = "#2C261C"
RESTORE = "#9A7420"

PLACEHOLDER = "写一条，回车添加"
WEEKDAYS = "一二三四五六日"

SPI_GETWORKAREA = 0x0030
GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_CAPTION = 0x00C00000
WS_THICKFRAME = 0x00040000
WS_BORDER = 0x00800000
WS_DLGFRAME = 0x00400000
WS_SYSMENU = 0x00080000
WS_MINIMIZEBOX = 0x00020000
WS_MAXIMIZEBOX = 0x00010000
WS_EX_DLGMODALFRAME = 0x00000001
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_WINDOWEDGE = 0x00000100
WS_EX_CLIENTEDGE = 0x00000200
WS_EX_APPWINDOW = 0x00040000
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020
HWND_TOPMOST = -1
RGN_OR = 2
ERROR_ALREADY_EXISTS = 183

_MUTEX = None
_USER32 = ctypes.WinDLL("user32", use_last_error=True)
_GDI32 = ctypes.WinDLL("gdi32", use_last_error=True)

_USER32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]
_USER32.GetWindowLongW.restype = ctypes.c_long
_USER32.SetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]
_USER32.SetWindowLongW.restype = ctypes.c_long
_USER32.SetWindowPos.argtypes = [
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_uint,
]
_USER32.SetWindowPos.restype = ctypes.c_int
_USER32.SetWindowRgn.argtypes = [ctypes.c_void_p, ctypes.c_void_p, wintypes.BOOL]
_USER32.SetWindowRgn.restype = ctypes.c_int
_USER32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
_USER32.ShowWindow.restype = wintypes.BOOL
_USER32.SystemParametersInfoW.argtypes = [
    wintypes.UINT,
    wintypes.UINT,
    ctypes.c_void_p,
    wintypes.UINT,
]
_USER32.SystemParametersInfoW.restype = wintypes.BOOL
_USER32.LockWindowUpdate.argtypes = [ctypes.c_void_p]
_USER32.LockWindowUpdate.restype = wintypes.BOOL

_GDI32.CreateEllipticRgn.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
_GDI32.CreateEllipticRgn.restype = ctypes.c_void_p
_GDI32.CreateRectRgn.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int]
_GDI32.CreateRectRgn.restype = ctypes.c_void_p
_GDI32.CombineRgn.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_int]
_GDI32.CombineRgn.restype = ctypes.c_int
_GDI32.DeleteObject.argtypes = [ctypes.c_void_p]
_GDI32.DeleteObject.restype = ctypes.c_int


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def enable_dpi() -> int:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    user32 = ctypes.WinDLL("user32")
    gdi32 = ctypes.WinDLL("gdi32")
    user32.GetDC.argtypes = [ctypes.c_void_p]
    user32.GetDC.restype = ctypes.c_void_p
    user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    gdi32.GetDeviceCaps.argtypes = [ctypes.c_void_p, ctypes.c_int]
    gdi32.GetDeviceCaps.restype = ctypes.c_int
    hdc = user32.GetDC(None)
    dpi = gdi32.GetDeviceCaps(hdc, 88) or 96
    user32.ReleaseDC(None, hdc)
    return int(dpi)


def work_area() -> tuple[int, int, int, int]:
    rect = RECT()
    ok = _USER32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
    if not ok:
        return (0, 0, 1920, 1040)
    return (int(rect.left), int(rect.top), int(rect.right), int(rect.bottom))


def ensure_single_instance() -> None:
    global _MUTEX
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, ctypes.c_wchar_p]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    _MUTEX = kernel32.CreateMutexW(None, False, "Local\\YYDeskNotesMutex")
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        sys.exit(0)


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_label() -> str:
    now = datetime.now()
    return f"{now.month}月{now.day}日 周{WEEKDAYS[now.weekday()]}"


def fmt_when(value: str | None) -> str:
    if not value:
        return ""
    try:
        dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return value
    return f"{dt.month}月{dt.day}日 {dt.hour:02d}:{dt.minute:02d}"


def log_exc(exc: BaseException | None = None) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as fh:
        fh.write("\n" + now_text() + "\n")
        if exc is None:
            traceback.print_exc(file=fh)
        else:
            traceback.print_exception(type(exc), exc, exc.__traceback__, file=fh)


def pick_font(root: tk.Misc) -> str:
    try:
        families = set(tkfont.families(root))
    except tk.TclError:
        families = set()
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "Segoe UI"):
        if name in families:
            return name
    return "TkDefaultFont"


class Store:
    def __init__(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.notes: list[dict] = []
        if path.exists():
            self.notes = self._load_csv(path)
        elif LEGACY_XLSX_PATH.exists():
            self.notes = self._load_xlsx(LEGACY_XLSX_PATH)
            self._save()
            self._backup(LEGACY_XLSX_PATH)
        elif LEGACY_DB_PATH.exists():
            self.notes = self._load_sqlite(LEGACY_DB_PATH)
            self._save()
            self._backup(LEGACY_DB_PATH)
        else:
            self._save()

    def close(self) -> None:
        self._save()

    def active_notes(self) -> list[dict]:
        return [dict(note) for note in self.notes if not note["deleted_at"]]

    def archived_notes(self) -> list[dict]:
        archived = [dict(note) for note in self.notes if note["deleted_at"]]
        archived.sort(key=lambda note: (note["deleted_at"] or "", note["id"]), reverse=True)
        return archived

    def counts(self) -> tuple[int, int, int, int]:
        active = pending = doing = archived = 0
        for note in self.notes:
            if note["deleted_at"]:
                archived += 1
                continue
            active += 1
            status = note.get("status", STATUS_TODO)
            if status == STATUS_DOING:
                doing += 1
            elif status != STATUS_DONE:
                pending += 1
        return active, pending, doing, archived

    def add(self, content: str) -> dict:
        stamp = now_text()
        note = {
            "id": self._next_id(),
            "content": content[:4000],
            "status": STATUS_TODO,
            "created_at": stamp,
            "updated_at": stamp,
            "deleted_at": None,
        }
        self.notes.append(note)
        self._save()
        return dict(note)

    def update_content(self, note_id: int, content: str) -> None:
        note = self._find(note_id)
        if note is None or note["deleted_at"]:
            return
        note["content"] = content[:4000]
        note["updated_at"] = now_text()
        self._save()

    def set_status(self, note_id: int, status: str) -> None:
        note = self._find(note_id)
        if note is None or status not in STATUS_LABEL:
            return
        note["status"] = status
        note["updated_at"] = now_text()
        self._save()

    def archive(self, note_id: int) -> None:
        note = self._find(note_id)
        if note is None or note["deleted_at"]:
            return
        stamp = now_text()
        note["deleted_at"] = stamp
        note["updated_at"] = stamp
        self._save()

    def restore(self, note_id: int) -> None:
        note = self._find(note_id)
        if note is None:
            return
        note["deleted_at"] = None
        note["updated_at"] = now_text()
        self._save()

    def purge(self, note_id: int) -> bool:
        note = self._find(note_id)
        if note is None:
            return False
        self.notes = [item for item in self.notes if int(item["id"]) != int(note_id)]
        self._save()
        return True

    def _next_id(self) -> int:
        return max((int(note["id"]) for note in self.notes), default=0) + 1

    def _find(self, note_id: int) -> dict | None:
        for note in self.notes:
            if int(note["id"]) == int(note_id):
                return note
        return None

    def _save(self) -> None:
        tmp = self.path.with_suffix(".csv.tmp")
        with tmp.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh)
            writer.writerow(HEADERS)
            for note in sorted(self.notes, key=lambda item: int(item["id"])):
                writer.writerow(
                    [
                        int(note["id"]),
                        note["content"],
                        STATUS_LABEL.get(note.get("status", STATUS_TODO), "待办"),
                        note["created_at"] or "",
                        note["updated_at"] or "",
                        note["deleted_at"] or "",
                    ]
                )
        os.replace(tmp, self.path)

    def _load_csv(self, path: Path) -> list[dict]:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.DictReader(fh))
        notes: list[dict] = []
        for item in rows:
            note_id = item.get("ID", item.get("id"))
            if note_id is None or str(note_id).strip() == "":
                continue
            content = str(item.get("内容", item.get("content", "")) or "").strip()
            notes.append(
                {
                    "id": int(note_id),
                    "content": content,
                    "status": self._parse_status(item),
                    "created_at": self._cell_text(item.get("写下时间", item.get("created_at"))),
                    "updated_at": self._cell_text(item.get("更新时间", item.get("updated_at"))),
                    "deleted_at": self._cell_text(item.get("归档时间", item.get("deleted_at"))) or None,
                }
            )
        return notes

    @staticmethod
    def _parse_status(item: dict) -> str:
        status = item.get("状态", item.get("status", None))
        if status is not None and str(status).strip() != "":
            text_v = str(status).strip().lower()
            mapping = {
                "todo": STATUS_TODO,
                "待办": STATUS_TODO,
                "否": STATUS_TODO,
                "0": STATUS_TODO,
                "doing": STATUS_DOING,
                "正在做": STATUS_DOING,
                "进行中": STATUS_DOING,
                "done": STATUS_DONE,
                "完成": STATUS_DONE,
                "是": STATUS_DONE,
                "1": STATUS_DONE,
                "yes": STATUS_DONE,
                "true": STATUS_DONE,
            }
            return mapping.get(text_v, mapping.get(str(status).strip(), STATUS_TODO))
        done_raw = item.get("完成", item.get("done", ""))
        if str(done_raw).strip() in {"1", "是", "true", "True", "YES", "yes", "完成", "done"}:
            return STATUS_DONE
        return STATUS_TODO

    def _load_xlsx(self, path: Path) -> list[dict]:
        from openpyxl import load_workbook

        book = load_workbook(path, data_only=True)
        sheet = book.active
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return []
        header = [str(cell or "").strip() for cell in rows[0]]
        notes: list[dict] = []
        for raw in rows[1:]:
            if raw is None or all(cell is None or str(cell).strip() == "" for cell in raw):
                continue
            item = dict(zip(header, raw))
            note_id = item.get("ID", item.get("id"))
            if note_id is None or str(note_id).strip() == "":
                continue
            content = str(item.get("内容", item.get("content", "")) or "").strip()
            notes.append(
                {
                    "id": int(note_id),
                    "content": content,
                    "status": self._parse_status(item),
                    "created_at": self._cell_text(item.get("写下时间", item.get("created_at"))),
                    "updated_at": self._cell_text(item.get("更新时间", item.get("updated_at"))),
                    "deleted_at": self._cell_text(item.get("归档时间", item.get("deleted_at"))) or None,
                }
            )
        return notes

    def _load_sqlite(self, path: Path) -> list[dict]:
        conn = sqlite3.connect(path)
        conn.row_factory = sqlite3.Row
        rows = [
            {
                "id": int(row["id"]),
                "content": row["content"],
                "status": STATUS_DONE if int(row["done"] or 0) else STATUS_TODO,
                "created_at": row["created_at"],
                "updated_at": row["updated_at"],
                "deleted_at": row["deleted_at"],
            }
            for row in conn.execute("SELECT * FROM notes ORDER BY id ASC")
        ]
        conn.close()
        return rows

    @staticmethod
    def _backup(path: Path) -> None:
        bak = path.with_name(path.name + ".bak")
        try:
            if bak.exists():
                bak.unlink()
            path.rename(bak)
        except OSError:
            pass

    @staticmethod
    def _cell_text(value) -> str:
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return str(value).strip()


class NoteRow(tk.Frame):
    def __init__(self, master: tk.Misc, app: "DeskNotes", note: dict) -> None:
        super().__init__(master, bg=PAPER, height=app.px(52))
        self.pack_propagate(False)
        self.app = app
        self.note_id = int(note["id"])
        self.status = note.get("status", STATUS_TODO)
        if self.status not in STATUS_LABEL:
            self.status = STATUS_TODO
        self._row_bg = PAPER
        self._trash_hover = False

        self.check = tk.Canvas(
            self,
            width=app.px(28),
            height=app.px(28),
            bg=PAPER,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.entry = tk.Entry(
            self,
            relief="flat",
            bd=0,
            highlightthickness=0,
            bg=PAPER,
            fg=INK,
            insertbackground=INK,
            selectbackground="#FFE08A",
            selectforeground=INK,
            font=app.font_body,
        )
        self.trash = tk.Canvas(
            self,
            width=app.px(32),
            height=app.px(32),
            bg=PAPER,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.sep = tk.Frame(self, bg=LINE, height=1)

        self.check.place(x=app.px(12), rely=0.5, anchor="w", width=app.px(28), height=app.px(28))
        self.trash.place(relx=1, x=-app.px(6), rely=0.5, anchor="e", width=app.px(32), height=app.px(32))
        self.entry.place(
            x=app.px(46),
            rely=0.5,
            anchor="w",
            relwidth=1,
            width=-app.px(46 + 44),
        )
        self.sep.place(x=app.px(16), rely=1.0, y=-1, relwidth=1, width=-app.px(32), height=1)

        self.entry.insert(0, note["content"])
        self.apply_status(self.status)
        self.entry.bind("<KeyRelease>", lambda _e: app.schedule_save(self.note_id))
        self.entry.bind("<<Paste>>", self._on_paste)
        self.entry.bind("<FocusOut>", lambda _e: app.save_content(self.note_id))
        self.check.bind("<Button-1>", self._on_check)
        self.trash.bind("<Button-1>", self._on_trash)
        self.trash.bind("<Enter>", self._trash_enter)
        self.trash.bind("<Leave>", self._trash_leave)
        self._bind_hover()

    def base_bg(self) -> str:
        if self.status == STATUS_DOING:
            return DOING_ROW
        return PAPER

    def apply_status(self, status: str) -> None:
        self.status = status if status in STATUS_LABEL else STATUS_TODO
        if self.status == STATUS_DONE:
            fg, font = DONE_FG, self.app.font_done
        elif self.status == STATUS_DOING:
            fg, font = DOING_FG, self.app.font_body
        else:
            fg, font = INK, self.app.font_body
        self.entry.configure(fg=fg, font=font)
        self._set_bg(self.base_bg())
        self.draw_check()
        self.draw_trash()

    def draw_check(self) -> None:
        canvas = self.check
        canvas.delete("all")
        size = self.app.px(18)
        pad = self.app.px(5)
        if self.status == STATUS_DONE:
            canvas.create_oval(
                pad, pad, pad + size, pad + size,
                fill=YELLOW, outline=YELLOW_EDGE, width=1,
            )
            canvas.create_line(
                pad + size * 0.28, pad + size * 0.52,
                pad + size * 0.46, pad + size * 0.70,
                pad + size * 0.76, pad + size * 0.32,
                fill=CHEVRON, width=2, capstyle="round", joinstyle="round",
            )
        elif self.status == STATUS_DOING:
            canvas.create_oval(
                pad, pad, pad + size, pad + size,
                fill=self._row_bg, outline=YELLOW_EDGE, width=2,
            )
            inner = size * 0.42
            inset = (size - inner) / 2
            canvas.create_oval(
                pad + inset, pad + inset,
                pad + inset + inner, pad + inset + inner,
                fill=YELLOW, outline="",
            )
        else:
            canvas.create_oval(
                pad, pad, pad + size, pad + size,
                fill=self._row_bg, outline="#CDBFA8", width=1,
            )

    def draw_trash(self) -> None:
        canvas = self.trash
        canvas.delete("all")
        width = max(canvas.winfo_width(), self.app.px(32))
        height = max(canvas.winfo_height(), self.app.px(32))
        if self._trash_hover:
            canvas.create_oval(2, 2, width - 2, height - 2, fill="#FDE8E4", outline="")
            color = DANGER
        elif self.status == STATUS_DONE:
            color = "#8A7D6C"
        else:
            color = "#C8BBA6"
        unit = width / 32.0
        cx, cy = width / 2, height / 2 + unit
        canvas.create_line(
            cx - 7 * unit, cy - 4 * unit, cx + 7 * unit, cy - 4 * unit,
            fill=color, width=1, capstyle="round",
        )
        canvas.create_line(
            cx - 2.2 * unit, cy - 4 * unit,
            cx - 2.2 * unit, cy - 6.4 * unit,
            cx + 2.2 * unit, cy - 6.4 * unit,
            cx + 2.2 * unit, cy - 4 * unit,
            fill=color, width=1, capstyle="round", joinstyle="round",
        )
        canvas.create_rectangle(
            cx - 5.5 * unit, cy - 3 * unit,
            cx + 5.5 * unit, cy + 6.2 * unit,
            outline=color, width=1,
        )
        canvas.create_line(cx - 2 * unit, cy - 1 * unit, cx - 2 * unit, cy + 4 * unit, fill=color)
        canvas.create_line(cx + 2 * unit, cy - 1 * unit, cx + 2 * unit, cy + 4 * unit, fill=color)

    def _set_bg(self, color: str) -> None:
        self._row_bg = color
        self.configure(bg=color)
        self.entry.configure(bg=color)
        self.check.configure(bg=color)
        if not self._trash_hover:
            self.trash.configure(bg=color)
        self.draw_check()

    def _bind_hover(self) -> None:
        def enter(_event=None) -> None:
            self._set_bg(ROW_HOVER)

        def leave(_event=None) -> None:
            x, y = self.winfo_pointerxy()
            target = self.winfo_containing(x, y)
            if target is not None and self._is_inside(target):
                return
            self._set_bg(self.base_bg())

        for widget in (self, self.check, self.entry, self.trash):
            widget.bind("<Enter>", enter, add="+")
            widget.bind("<Leave>", leave, add="+")

    def _is_inside(self, widget: tk.Misc) -> bool:
        name = str(self)
        current = str(widget)
        return current == name or current.startswith(name + ".")

    def _on_check(self, _event) -> str:
        self.app.cycle_status(self.note_id)
        return "break"

    def _on_trash(self, _event) -> str:
        self.app.archive_note(self.note_id)
        return "break"

    def _trash_enter(self, _event) -> None:
        self._trash_hover = True
        self.draw_trash()

    def _trash_leave(self, _event) -> None:
        self._trash_hover = False
        self.trash.configure(bg=self._row_bg)
        self.draw_trash()

    def _on_paste(self, _event) -> str | None:
        try:
            text = self.clipboard_get()
        except tk.TclError:
            return None
        if "\n" not in text and "\r" not in text:
            return None
        flat = " ".join(part.strip() for part in text.replace("\r", "\n").split("\n") if part.strip())
        if flat:
            self.entry.insert("insert", flat[:4000])
            self.app.schedule_save(self.note_id)
        return "break"


class DeskNotes(tk.Tk):
    def __init__(self, dpi: int) -> None:
        super().__init__()
        self.dpi = dpi
        self.scale = dpi / 96.0
        expect = dpi / 72.0
        try:
            current = float(self.tk.call("tk", "scaling"))
        except tk.TclError:
            current = expect
        if abs(current - expect) > 0.08:
            self.tk.call("tk", "scaling", expect)

        self.store = Store(CSV_PATH)
        self.family = pick_font(self)
        self.font_title = tkfont.Font(family=self.family, size=16, weight="bold")
        self.font_body = tkfont.Font(family=self.family, size=12)
        self.font_done = tkfont.Font(family=self.family, size=12, overstrike=1)
        self.font_small = tkfont.Font(family=self.family, size=9)
        self.font_small_bold = tkfont.Font(family=self.family, size=9, weight="bold")
        self.font_meta = tkfont.Font(family=self.family, size=8)

        self.dot = self.px(DOT)
        self.panel_w = self.px(PANEL_W)
        self.work = work_area()
        work_h = self.work[3] - self.work[1]
        self.panel_h = max(self.px(360), min(self.px(PANEL_H), work_h - self.px(36)))

        self.opened = False
        self.mode = "notes"
        self._progress = 0.0
        self._anim_gen = 0
        self._hwnd: int | None = None
        self._borderless_done = False
        self.rows: dict[int, NoteRow] = {}
        self._save_after: dict[int, str] = {}
        self._undo_id: int | None = None
        self._toast_after: str | None = None
        self._menu: tk.Menu | None = None

        self.withdraw()
        self.title("桌面备忘")
        self.configure(bg=PAPER)
        self.resizable(False, False)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.withdraw()

        self._build()
        self.reload_notes()
        self.refresh_counts()
        self._prepare_placeholder()
        self.withdraw()
        self.update_idletasks()
        self._hwnd = int(self.wm_frame(), 16)
        self._apply_borderless()
        self._apply_layout()
        self.deiconify()
        self.update()
        self._apply_layout()
        self.bind("<Escape>", self._on_escape)
        self.bind_all("<MouseWheel>", self._on_wheel)
        self.after(1500, self._keep_docked)
        self.after(30000, self._tick_clock)
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

    def px(self, value: int) -> int:
        return max(1, int(round(value * self.scale)))

    def report_callback_exception(self, exc_type, exc, tb) -> None:  # type: ignore[override]
        log_exc(exc)

    def _build(self) -> None:
        self.panel = tk.Frame(
            self,
            bg=PAPER,
            width=self.panel_w,
            height=self.panel_h,
            highlightthickness=1,
            highlightbackground="#E7DCC4",
            highlightcolor="#E7DCC4",
            bd=0,
        )
        self.panel.grid_columnconfigure(0, weight=1)
        self.panel.grid_rowconfigure(2, weight=1)

        accent = tk.Frame(self.panel, bg=YELLOW, height=self.px(4))
        accent.grid(row=0, column=0, sticky="ew")
        accent.grid_propagate(False)

        header = tk.Frame(self.panel, bg=PAPER, height=self.px(72))
        header.grid(row=1, column=0, sticky="ew")
        header.grid_propagate(False)
        self.title_label = tk.Label(
            header, text="备忘", bg=PAPER, fg=INK, font=self.font_title,
        )
        self.title_label.place(x=self.px(18), y=self.px(12))
        self.subtitle = tk.Label(
            header, text="点圆圈：待办 → 正在做 → 完成", bg=PAPER, fg=MUTED, font=self.font_small,
        )
        self.subtitle.place(x=self.px(18), y=self.px(44))
        self.date_label = tk.Label(
            header, text=today_label(), bg=PAPER, fg=MUTED, font=self.font_small,
        )
        self.date_label.place(relx=1, x=-self.px(16), y=self.px(14), anchor="ne")
        self.more = tk.Label(
            header, text="···", bg=PAPER, fg=MUTED, font=self.font_body, cursor="hand2",
        )
        self.more.place(relx=1, x=-self.px(16), y=self.px(42), anchor="ne")
        self.more.bind("<Button-1>", self.popup_menu)
        self.more.bind("<Enter>", lambda _e: self.more.configure(fg=INK))
        self.more.bind("<Leave>", lambda _e: self.more.configure(fg=MUTED))

        self.body = tk.Frame(self.panel, bg=PAPER)
        self.body.grid(row=2, column=0, sticky="nsew")
        self.body.grid_rowconfigure(0, weight=1)
        self.body.grid_columnconfigure(0, weight=1)

        self.notes_view, self.list_canvas, self.list_inner = self._make_scroller(self.body)
        self.notes_view.grid(row=0, column=0, sticky="nsew")
        self.empty_label = tk.Label(
            self.notes_view,
            text="还没有备忘\n在下面写一条",
            bg=PAPER,
            fg=MUTED,
            font=self.font_body,
            justify="center",
        )

        self.archive_view, self.archive_canvas, self.archive_inner = self._make_scroller(self.body)

        self.toast = tk.Frame(self.panel, bg=TOAST_BG, height=self.px(34))
        self.toast.grid(row=3, column=0, sticky="ew")
        self.toast.grid_propagate(False)
        tk.Label(
            self.toast,
            text="已归档，日期已记下",
            bg=TOAST_BG,
            fg="#FFFDF8",
            font=self.font_small,
        ).pack(side="left", padx=self.px(14))
        undo = tk.Label(
            self.toast,
            text="撤销",
            bg=TOAST_BG,
            fg=YELLOW,
            font=self.font_small_bold,
            cursor="hand2",
        )
        undo.pack(side="right", padx=self.px(14))
        undo.bind("<Button-1>", lambda _e: self.undo_archive())
        self.toast.grid_remove()

        self.composer_wrap = tk.Frame(self.panel, bg=PAPER)
        self.composer_wrap.grid(row=4, column=0, sticky="ew")
        well = tk.Frame(self.composer_wrap, bg=INPUT_BG)
        well.pack(fill="x", padx=self.px(14), pady=(self.px(8), self.px(4)))
        plus = tk.Label(
            well, text="+", bg=INPUT_BG, fg="#C4A24A", font=self.font_title, cursor="hand2",
        )
        plus.pack(side="left", padx=(self.px(10), self.px(2)))
        plus.bind("<Button-1>", lambda _e: self.add_note())
        self.composer = tk.Entry(
            well,
            relief="flat",
            bd=0,
            highlightthickness=0,
            bg=INPUT_BG,
            fg=MUTED,
            insertbackground=INK,
            selectbackground="#FFE08A",
            selectforeground=INK,
            font=self.font_body,
        )
        self.composer.pack(side="left", fill="x", expand=True, padx=(0, self.px(10)), pady=self.px(8))
        self.composer.bind("<Return>", self.add_note)
        self.composer.bind("<KP_Enter>", self.add_note)
        self.composer.bind("<FocusIn>", self._composer_focus_in)
        self.composer.bind("<FocusOut>", self._composer_focus_out)
        self.composer.bind("<<Paste>>", self._composer_paste)

        footer = tk.Frame(self.panel, bg=PAPER, height=self.px(34))
        footer.grid(row=5, column=0, sticky="ew")
        footer.grid_propagate(False)
        self.footer_link = tk.Label(
            footer, text="归档", bg=PAPER, fg=MUTED, font=self.font_small, cursor="hand2",
        )
        self.footer_link.place(x=self.px(16), rely=0.5, anchor="w")
        self.footer_link.bind("<Button-1>", lambda _e: self.on_footer())
        self.footer_link.bind("<Enter>", lambda _e: self.footer_link.configure(fg=INK))
        self.footer_link.bind("<Leave>", lambda _e: self.footer_link.configure(fg=MUTED))
        tk.Label(
            footer,
            text="右键黄点可以退出",
            bg=PAPER,
            fg="#C0B6A6",
            font=self.font_meta,
        ).place(relx=1, x=-self.px(14), rely=0.5, anchor="e")

        self.dot_canvas = tk.Canvas(
            self,
            width=self.dot,
            height=self.dot,
            bg=YELLOW,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.dot_canvas.bind("<Button-1>", lambda _e: self.toggle())
        self.dot_canvas.bind("<Button-3>", self.popup_menu)
        self.dot_canvas.bind("<Enter>", lambda _e: self._paint_dot(True))
        self.dot_canvas.bind("<Leave>", lambda _e: self._paint_dot(False))
        self._paint_dot(False)

    def _make_scroller(self, parent: tk.Misc):
        wrap = tk.Frame(parent, bg=PAPER)
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        canvas = tk.Canvas(wrap, bg=PAPER, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(
            wrap,
            command=canvas.yview,
            width=self.px(8),
            relief="flat",
            bd=0,
            bg="#E4D7BE",
            troughcolor=PAPER,
            activebackground="#D5C6A6",
            highlightthickness=0,
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        inner = tk.Frame(canvas, bg=PAPER)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_inner(_event, canvas=canvas, scrollbar=scrollbar) -> None:
            canvas.configure(scrollregion=canvas.bbox("all") or (0, 0, 0, 0))
            self._sync_scrollbar(canvas, scrollbar)

        def on_canvas(event, canvas=canvas, window_id=window_id) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        inner.bind("<Configure>", on_inner)
        canvas.bind("<Configure>", on_canvas)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        scrollbar.grid_remove()
        return wrap, canvas, inner

    def _sync_scrollbar(self, canvas: tk.Canvas, scrollbar: tk.Scrollbar) -> None:
        bbox = canvas.bbox("all")
        if not bbox or bbox[3] <= canvas.winfo_height() + 1:
            scrollbar.grid_remove()
        else:
            scrollbar.grid()

    def _prepare_placeholder(self) -> None:
        self.composer.delete(0, "end")
        self.composer.insert(0, PLACEHOLDER)
        self.composer.configure(fg=MUTED)

    def _composer_focus_in(self, _event=None) -> None:
        if self.composer.get() == PLACEHOLDER:
            self.composer.delete(0, "end")
            self.composer.configure(fg=INK)

    def _composer_focus_out(self, _event=None) -> None:
        if not self.composer.get().strip():
            self._prepare_placeholder()

    def _composer_paste(self, _event) -> str | None:
        try:
            text = self.clipboard_get()
        except tk.TclError:
            return None
        if "\n" not in text and "\r" not in text:
            return None
        self._composer_focus_in()
        self.add_lines(text)
        return "break"

    def _paint_dot(self, hover: bool) -> None:
        canvas = self.dot_canvas
        canvas.configure(bg=YELLOW_HOVER if hover else YELLOW)
        canvas.delete("mark")
        size = self.dot
        canvas.create_oval(
            1, 1, size - 2, size - 2,
            outline=YELLOW_EDGE, width=1, tags="mark",
        )
        cx = size / 2
        cy = size / 2
        color = CHEVRON
        if self.opened:
            canvas.create_line(
                cx - 4, cy - 5, cx + 3, cy, cx - 4, cy + 5,
                fill=color, width=2, capstyle="round", joinstyle="round", tags="mark",
            )
        else:
            canvas.create_line(
                cx + 4, cy - 5, cx - 3, cy, cx + 4, cy + 5,
                fill=color, width=2, capstyle="round", joinstyle="round", tags="mark",
            )

    def _apply_borderless(self) -> None:
        if not self._hwnd:
            return
        hwnd = ctypes.c_void_p(self._hwnd)
        style = _USER32.GetWindowLongW(hwnd, GWL_STYLE)
        style &= ~(
            WS_CAPTION | WS_THICKFRAME | WS_BORDER | WS_DLGFRAME
            | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX
        )
        _USER32.SetWindowLongW(hwnd, GWL_STYLE, style)
        exstyle = _USER32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        exstyle &= ~(WS_EX_APPWINDOW | WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE | WS_EX_DLGMODALFRAME)
        exstyle |= WS_EX_TOOLWINDOW
        _USER32.SetWindowLongW(hwnd, GWL_EXSTYLE, exstyle)
        _USER32.SetWindowPos(
            hwnd,
            ctypes.c_void_p(HWND_TOPMOST),
            0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOACTIVATE,
        )
        self._borderless_done = True

    def _apply_region(self, width: int, height: int) -> None:
        if not self._hwnd:
            return
        hwnd = ctypes.c_void_p(self._hwnd)
        if width <= self.dot + 1:
            region = _GDI32.CreateEllipticRgn(0, 0, width, height)
            _USER32.SetWindowRgn(hwnd, region, True)
            return
        dot_y = max(0, (height - self.dot) // 2)
        rect = _GDI32.CreateRectRgn(self.dot, 0, width, height)
        ellipse = _GDI32.CreateEllipticRgn(0, dot_y, self.dot, dot_y + self.dot)
        combined = _GDI32.CreateRectRgn(0, 0, 1, 1)
        merged = _GDI32.CombineRgn(combined, rect, ellipse, RGN_OR)
        _GDI32.DeleteObject(rect)
        _GDI32.DeleteObject(ellipse)
        if merged == 0:
            _GDI32.DeleteObject(combined)
            return
        _USER32.SetWindowRgn(hwnd, combined, True)

    def _shell_geometry(self, open_shell: bool) -> tuple[int, int, int, int]:
        left, top, right, bottom = self.work
        if open_shell:
            width = self.dot + self.panel_w
            height = self.panel_h
        else:
            width = height = self.dot
        x = right - width
        y = top + max(0, ((bottom - top) - height) // 2)
        return width, height, x, y

    def _apply_shell(self, open_shell: bool) -> None:
        """只在展开/收起起止改窗口尺寸和异形区域，避免动画逐帧 SetWindowRgn 闪烁。"""
        hwnd = ctypes.c_void_p(self._hwnd) if self._hwnd else None
        if hwnd:
            _USER32.LockWindowUpdate(hwnd)
        try:
            width, height, x, y = self._shell_geometry(open_shell)
            if not open_shell:
                self.panel.place_forget()
            self.geometry(f"{width}x{height}+{x}+{y}")
            dot_y = max(0, (height - self.dot) // 2)
            self.dot_canvas.place(x=0, y=dot_y, width=self.dot, height=self.dot)
            tk.Misc.lift(self.dot_canvas)
            self.update_idletasks()
            self._apply_region(width, height)
        finally:
            if hwnd:
                _USER32.LockWindowUpdate(None)

    def _place_panel(self, progress: float) -> None:
        """progress: 0=面板在右侧外侧，1=完全展开。只移动控件，不改窗口外形。"""
        height = self.panel_h
        panel_x = int(round(self.dot + self.panel_w * (1.0 - progress)))
        self.panel.place(x=panel_x, y=0, width=self.panel_w, height=height)
        dot_y = max(0, (height - self.dot) // 2)
        self.dot_canvas.place(x=0, y=dot_y, width=self.dot, height=self.dot)
        tk.Misc.lift(self.dot_canvas)

    def _apply_layout(self) -> None:
        if self._progress <= 0.001:
            self._apply_shell(False)
            return
        self._apply_shell(True)
        self._place_panel(self._progress)

    def _animate_to(self, target: float) -> None:
        self._anim_gen += 1
        generation = self._anim_gen
        origin = self._progress
        started = datetime.now().timestamp()

        # 展开：先一次扩成完整窗口，再只滑面板（不逐帧改 region）
        if target >= 1.0 and origin < 0.999:
            self._progress = 0.0
            self._apply_shell(True)
            self._place_panel(0.0)
            origin = 0.0
        # 收起：外壳保持完整，只滑面板，结束后再缩成黄点
        elif target <= 0.0 and origin > 0.001:
            self._apply_shell(True)
            self._place_panel(origin)

        def tick() -> None:
            if generation != self._anim_gen:
                return
            elapsed = (datetime.now().timestamp() - started) * 1000
            t = min(1.0, elapsed / ANIM_MS)
            eased = 1 - (1 - t) ** 3
            self._progress = origin + (target - origin) * eased
            if self._progress <= 0.001:
                self.panel.place_forget()
            else:
                self._place_panel(max(self._progress, 0.001))
            if t < 1:
                self.after(12, tick)
                return
            if target <= 0.0:
                self._progress = 0.0
                self._apply_shell(False)
            elif self.opened and self.mode == "notes":
                self._focus_composer()

        tick()

    def _focus_composer(self) -> None:
        if not (self.opened and self.mode == "notes"):
            return
        try:
            self.composer.focus_set()
            self._composer_focus_in()
        except tk.TclError:
            pass

    def toggle(self) -> None:
        self.set_open(not self.opened)

    def set_open(self, opened: bool) -> None:
        self.opened = opened
        self._paint_dot(False)
        self.date_label.configure(text=today_label())
        self._animate_to(1.0 if opened else 0.0)

    def _keep_docked(self) -> None:
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        area = work_area()
        if area != self.work and self._anim_gen >= 0:
            self.work = area
            work_h = area[3] - area[1]
            self.panel_h = max(self.px(360), min(self.px(PANEL_H), work_h - self.px(36)))
            self._apply_layout()
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            return
        self.after(2000, self._keep_docked)

    def _tick_clock(self) -> None:
        self.date_label.configure(text=today_label())
        self.after(30000, self._tick_clock)

    def _on_escape(self, _event=None):
        if self.mode == "archive":
            self.show_notes()
            return "break"
        if self.opened:
            self.set_open(False)
            return "break"
        return None

    def _on_wheel(self, event) -> None:
        if not self.opened:
            return
        canvas = self.archive_canvas if self.mode == "archive" else self.list_canvas
        canvas.yview_scroll(int(-event.delta / 120), "units")

    def popup_menu(self, event=None) -> str:
        if self._menu is not None:
            try:
                self._menu.destroy()
            except tk.TclError:
                pass
        menu = tk.Menu(self, tearoff=0, font=self.font_small)
        self._menu = menu
        if self.mode == "archive":
            menu.add_command(label="返回备忘", command=self.show_notes)
        else:
            menu.add_command(label="查看归档", command=self.show_archive)
        menu.add_command(label="打开 CSV", command=self.open_csv)
        menu.add_command(label="打开数据文件夹", command=self.open_data_folder)
        menu.add_separator()
        menu.add_command(label="退出", command=self.quit_app)

        # 本窗口是 topmost + 黄点始终置顶，系统菜单偶发会落在它下面被挡住。
        # 弹出时先取消置顶，并强制菜单开在黄点左侧。
        menu_w = self.px(168)
        handle_left = self.winfo_rootx()
        if event is not None and getattr(event, "x_root", 0):
            y = int(event.y_root)
        else:
            y = self.more.winfo_rooty() + self.more.winfo_height()
        x = max(0, handle_left - menu_w - self.px(6))

        try:
            self.attributes("-topmost", False)
            self.update_idletasks()
            menu.tk_popup(x, y)
        finally:
            try:
                menu.grab_release()
            except tk.TclError:
                pass
            self.after(80, self._restore_topmost)
        return "break"

    def _restore_topmost(self) -> None:
        try:
            if not self.winfo_exists():
                return
            self.attributes("-topmost", True)
            tk.Misc.lift(self.dot_canvas)
        except tk.TclError:
            pass

    def open_csv(self) -> None:
        self.flush_saves()
        self.store.close()
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(CSV_PATH)  # noqa: S606

    def open_data_folder(self) -> None:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(DATA_DIR)  # noqa: S606

    def reload_notes(self) -> None:
        self.flush_saves()
        for row in self.rows.values():
            row.destroy()
        self.rows.clear()
        for note in self.store.active_notes():
            self._append_row(note, scroll=False)
        self._update_empty()
        self.list_inner.update_idletasks()
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all") or (0, 0, 0, 0))

    def _append_row(self, note: dict, scroll: bool = True) -> None:
        row = NoteRow(self.list_inner, self, note)
        row.pack(fill="x")
        self.rows[int(note["id"])] = row
        self._update_empty()
        if scroll:
            self.list_inner.update_idletasks()
            self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all") or (0, 0, 0, 0))
            self.list_canvas.yview_moveto(1.0)

    def _update_empty(self) -> None:
        if self.rows:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.42, anchor="center")
            tk.Misc.lift(self.empty_label)

    def refresh_counts(self) -> None:
        active, pending, doing, archived = self.store.counts()
        if self.mode == "notes":
            if active == 0:
                self.subtitle.configure(text="点圆圈：待办 → 正在做 → 完成")
            elif doing:
                parts = [f"{doing} 件正在做"]
                if pending:
                    parts.append(f"{pending} 件待办")
                self.subtitle.configure(text=" · ".join(parts))
            elif pending == 0:
                self.subtitle.configure(text="都做完了")
            else:
                self.subtitle.configure(text=f"{pending} 件待办")
            self.footer_link.configure(text=f"归档  {archived}" if archived else "归档")

    def add_note(self, _event=None):
        text = self.composer.get()
        if text == PLACEHOLDER:
            return "break"
        self.add_lines(text)
        return "break"

    def add_lines(self, text: str) -> None:
        lines = [
            line.strip()
            for line in text.replace("\r", "\n").split("\n")
            if line.strip()
        ]
        if not lines:
            return
        for line in lines:
            note = self.store.add(line)
            self._append_row(note)
        self.composer.delete(0, "end")
        self.composer.configure(fg=INK)
        self.refresh_counts()

    def schedule_save(self, note_id: int) -> None:
        job = self._save_after.pop(note_id, None)
        if job:
            try:
                self.after_cancel(job)
            except tk.TclError:
                pass
        self._save_after[note_id] = self.after(280, lambda: self.save_content(note_id))

    def save_content(self, note_id: int) -> None:
        job = self._save_after.pop(note_id, None)
        if job:
            try:
                self.after_cancel(job)
            except tk.TclError:
                pass
        row = self.rows.get(note_id)
        if row is None:
            return
        self.store.update_content(note_id, row.entry.get())

    def flush_saves(self) -> None:
        for note_id in list(self._save_after):
            self.save_content(note_id)

    def cycle_status(self, note_id: int) -> None:
        row = self.rows.get(note_id)
        if row is None:
            return
        current = row.status if row.status in STATUS_CYCLE else STATUS_TODO
        nxt = STATUS_CYCLE[(STATUS_CYCLE.index(current) + 1) % len(STATUS_CYCLE)]
        self.store.set_status(note_id, nxt)
        row.apply_status(nxt)
        self.refresh_counts()

    def archive_note(self, note_id: int) -> None:
        row = self.rows.get(note_id)
        if row is None:
            return
        self.save_content(note_id)
        self.store.archive(note_id)
        self.rows.pop(note_id, None)
        row.destroy()
        self._update_empty()
        self.refresh_counts()
        self._show_toast(note_id)

    def _show_toast(self, note_id: int) -> None:
        self._undo_id = note_id
        if self.mode == "notes":
            self.toast.grid()
        if self._toast_after:
            try:
                self.after_cancel(self._toast_after)
            except tk.TclError:
                pass
        self._toast_after = self.after(6000, self._hide_toast)

    def _hide_toast(self) -> None:
        self._undo_id = None
        self._toast_after = None
        self.toast.grid_remove()

    def undo_archive(self) -> None:
        note_id = self._undo_id
        if not note_id:
            return
        self._hide_toast()
        self.store.restore(note_id)
        self.reload_notes()
        self.refresh_counts()

    def restore_note(self, note_id: int) -> None:
        self.store.restore(note_id)
        self.reload_notes()
        self.reload_archive()
        self.refresh_counts()

    def show_archive(self) -> None:
        if not self.opened:
            self.set_open(True)
        self.mode = "archive"
        self.notes_view.grid_remove()
        self.archive_view.grid(row=0, column=0, sticky="nsew")
        self.composer_wrap.grid_remove()
        self.toast.grid_remove()
        self.title_label.configure(text="归档")
        self.subtitle.configure(text="删掉的都留着，连日期一起")
        self.footer_link.configure(text="返回备忘")
        self.reload_archive()

    def show_notes(self) -> None:
        self.mode = "notes"
        self.archive_view.grid_remove()
        self.notes_view.grid(row=0, column=0, sticky="nsew")
        self.composer_wrap.grid()
        self.title_label.configure(text="备忘")
        if self._undo_id:
            self.toast.grid()
        self.refresh_counts()

    def on_footer(self) -> None:
        if self.mode == "archive":
            self.show_notes()
        else:
            self.show_archive()

    def reload_archive(self) -> None:
        for child in self.archive_inner.winfo_children():
            child.destroy()
        notes = self.store.archived_notes()
        if not notes:
            tk.Label(
                self.archive_inner,
                text="还没有归档\n点每行右边的垃圾桶，会连日期留在这里",
                bg=PAPER,
                fg=MUTED,
                font=self.font_body,
                justify="center",
            ).pack(pady=self.px(80))
            return
        for note in notes:
            self._archive_item(note)
        self.archive_inner.update_idletasks()
        self.archive_canvas.configure(
            scrollregion=self.archive_canvas.bbox("all") or (0, 0, 0, 0)
        )
        self.archive_canvas.yview_moveto(0)

    def _archive_item(self, note: dict) -> None:
        block = tk.Frame(self.archive_inner, bg=PAPER)
        block.pack(fill="x")
        meta = (
            f"写下 {fmt_when(note['created_at'])}"
            f"    归档 {fmt_when(note['deleted_at'])}"
        )
        tk.Label(
            block, text=meta, bg=PAPER, fg=MUTED, font=self.font_meta, anchor="w",
        ).pack(fill="x", padx=self.px(18), pady=(self.px(12), 0))
        line = tk.Frame(block, bg=PAPER)
        line.pack(fill="x", padx=self.px(18), pady=(self.px(2), self.px(8)))
        status = note.get("status", STATUS_TODO)
        if status == STATUS_DONE:
            fg, font = DONE_FG, self.font_done
        elif status == STATUS_DOING:
            fg, font = DOING_FG, self.font_body
        else:
            fg, font = INK, self.font_body
        tk.Label(
            line,
            text=STATUS_LABEL.get(status, "待办"),
            bg=PAPER,
            fg=MUTED,
            font=self.font_meta,
        ).pack(side="left", padx=(0, self.px(8)))
        tk.Label(
            line,
            text=note["content"],
            bg=PAPER,
            fg=fg,
            font=font,
            anchor="w",
            justify="left",
            wraplength=max(80, self.panel_w - self.px(170)),
        ).pack(side="left", fill="x", expand=True)
        note_id = int(note["id"])
        trash = tk.Canvas(
            line,
            width=self.px(28),
            height=self.px(28),
            bg=PAPER,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        trash.pack(side="right", padx=(self.px(2), 0))
        self._draw_trash_icon(trash, DANGER)
        trash.bind("<Button-1>", lambda _e, note_id=note_id: self.purge_note(note_id))
        undo = tk.Label(
            line,
            text="↩",
            bg=PAPER,
            fg=RESTORE,
            font=self.font_body,
            cursor="hand2",
            padx=self.px(4),
        )
        undo.pack(side="right", padx=(self.px(2), 0))
        undo.bind("<Button-1>", lambda _e, note_id=note_id: self.restore_note(note_id))
        tk.Frame(block, bg=LINE, height=1).pack(fill="x", padx=self.px(18))

    def _draw_trash_icon(self, canvas: tk.Canvas, color: str) -> None:
        canvas.delete("all")
        width = max(canvas.winfo_reqwidth(), self.px(28))
        height = max(canvas.winfo_reqheight(), self.px(28))
        unit = width / 28.0
        cx, cy = width / 2, height / 2 + unit
        canvas.create_line(
            cx - 6 * unit, cy - 4 * unit, cx + 6 * unit, cy - 4 * unit,
            fill=color, width=1, capstyle="round",
        )
        canvas.create_line(
            cx - 2 * unit, cy - 4 * unit,
            cx - 2 * unit, cy - 6 * unit,
            cx + 2 * unit, cy - 6 * unit,
            cx + 2 * unit, cy - 4 * unit,
            fill=color, width=1, capstyle="round", joinstyle="round",
        )
        canvas.create_rectangle(
            cx - 5 * unit, cy - 3 * unit,
            cx + 5 * unit, cy + 5.5 * unit,
            outline=color, width=1,
        )
        canvas.create_line(cx - 1.8 * unit, cy - 1 * unit, cx - 1.8 * unit, cy + 3.5 * unit, fill=color)
        canvas.create_line(cx + 1.8 * unit, cy - 1 * unit, cx + 1.8 * unit, cy + 3.5 * unit, fill=color)

    def purge_note(self, note_id: int) -> None:
        preview = ""
        for note in self.store.archived_notes():
            if int(note["id"]) == int(note_id):
                preview = str(note.get("content") or "").strip()
                break
        if len(preview) > 40:
            preview = preview[:40] + "…"
        message = "确定彻底删除这条归档吗？删除后不可恢复。"
        if preview:
            message = f"确定彻底删除这条归档吗？\n\n「{preview}」\n\n删除后不可恢复。"
        try:
            self.attributes("-topmost", False)
            self.update_idletasks()
            ok = messagebox.askyesno("彻底删除", message, parent=self, default="no")
        finally:
            self.after(80, self._restore_topmost)
        if not ok:
            return
        if self.store.purge(note_id):
            self.reload_archive()
            self.refresh_counts()

    def quit_app(self) -> None:
        try:
            self.flush_saves()
        except Exception as exc:
            log_exc(exc)
        try:
            self.store.close()
        except Exception:
            pass
        self.destroy()


def main() -> None:
    ensure_single_instance()
    dpi = enable_dpi()
    app = DeskNotes(dpi)
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:
        log_exc(exc)
        raise
