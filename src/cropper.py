# -*- coding: utf-8 -*-
"""Диалог выбора фото с обрезкой под заданное соотношение сторон."""

from __future__ import annotations

import io
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Optional

from PIL import Image, ImageOps, ImageTk

import theme

RATIOS = {
    "3:4": (3, 4), "1:1": (1, 1), "4:3": (4, 3), "2:3": (2, 3), "3:2": (3, 2),
    "4:5": (4, 5), "5:4": (5, 4), "9:16": (9, 16), "16:9": (16, 9),
}
_CW, _CH, _MARGIN = 520, 560, 20
_WORK_SIDE = 2400  # исходник уменьшается до этого размера для быстрой работы
_JPEG_QUALITY = 90


class CropDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, path: str, ratio: str = "3:4", max_side: int = 800):
        super().__init__(parent)
        self.title("Фото — обрезка")
        self.transient(parent)
        self.resizable(False, False)
        self.configure(background=theme.BG)
        self.result: Optional[bytes] = None

        img = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
        img.thumbnail((_WORK_SIDE, _WORK_SIDE))
        self.img = img

        self.ratio = tk.StringVar(value=ratio if ratio in RATIOS else "3:4")
        self.max_side = tk.StringVar(value=str(max_side))
        self.zoom = tk.DoubleVar(value=1.0)

        top = ttk.Frame(self, padding=6)
        top.pack(fill="x")
        ttk.Label(top, text="Формат:").pack(side="left")
        box = ttk.Combobox(top, textvariable=self.ratio, values=list(RATIOS), state="readonly", width=6)
        box.pack(side="left", padx=(4, 12))
        box.bind("<<ComboboxSelected>>", lambda e: self._reset())
        ttk.Label(top, text="Макс. сторона, px:").pack(side="left")
        ttk.Entry(top, textvariable=self.max_side, width=6).pack(side="left", padx=4)

        self.canvas = tk.Canvas(self, width=_CW, height=_CH, bg="#1f2937", highlightthickness=0)
        self.canvas.pack()
        self.canvas.bind("<ButtonPress-1>", self._drag_begin)
        self.canvas.bind("<B1-Motion>", self._drag_move)
        self.canvas.bind("<MouseWheel>", lambda e: self._zoom_by(1.1 if e.delta > 0 else 1 / 1.1))
        self.canvas.bind("<Button-4>", lambda e: self._zoom_by(1.1))
        self.canvas.bind("<Button-5>", lambda e: self._zoom_by(1 / 1.1))

        zoom_row = ttk.Frame(self, padding=(6, 4))
        zoom_row.pack(fill="x")
        ttk.Label(zoom_row, text="Масштаб:").pack(side="left")
        ttk.Scale(zoom_row, from_=1.0, to=5.0, variable=self.zoom, command=self._on_zoom_slider).pack(
            side="left", fill="x", expand=True, padx=6)

        btns = ttk.Frame(self, padding=6)
        btns.pack(fill="x")
        ttk.Button(btns, text="Без обрезки", command=lambda: self._finish(crop=False)).pack(side="left")
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Готово", style="Accent.TButton", command=lambda: self._finish(crop=True)).pack(side="right", padx=6)

        self._photo_ref = None
        self._drag = None
        self._reset()
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())

    # --- геометрия -----------------------------------------------------
    def _frame(self):
        rw, rh = RATIOS[self.ratio.get()]
        k = min((_CW - 2 * _MARGIN) / rw, (_CH - 2 * _MARGIN) / rh)
        fw, fh = rw * k, rh * k
        x0, y0 = (_CW - fw) / 2, (_CH - fh) / 2
        return x0, y0, x0 + fw, y0 + fh

    def _scale(self) -> float:
        x0, y0, x1, y1 = self._frame()
        cover = max((x1 - x0) / self.img.width, (y1 - y0) / self.img.height)
        return cover * self.zoom.get()

    def _reset(self):
        self.zoom.set(1.0)
        self._last_zoom = 1.0
        self.cx, self.cy = _CW / 2, _CH / 2
        self._redraw()

    def _clamp(self):
        x0, y0, x1, y1 = self._frame()
        s = self._scale()
        hw, hh = self.img.width * s / 2, self.img.height * s / 2
        self.cx = min(max(self.cx, x1 - hw), x0 + hw)
        self.cy = min(max(self.cy, y1 - hh), y0 + hh)

    def _redraw(self):
        self._clamp()
        c = self.canvas
        c.delete("all")
        s = self._scale()
        ix0, iy0 = self.cx - self.img.width * s / 2, self.cy - self.img.height * s / 2
        ix1, iy1 = ix0 + self.img.width * s, iy0 + self.img.height * s
        vx0, vy0, vx1, vy1 = max(ix0, 0), max(iy0, 0), min(ix1, _CW), min(iy1, _CH)
        if vx1 - vx0 >= 1 and vy1 - vy0 >= 1:
            part = self.img.crop((
                (vx0 - ix0) / s, (vy0 - iy0) / s, (vx1 - ix0) / s, (vy1 - iy0) / s,
            )).resize((max(1, round(vx1 - vx0)), max(1, round(vy1 - vy0))), Image.BILINEAR)
            self._photo_ref = ImageTk.PhotoImage(part)
            c.create_image(vx0, vy0, image=self._photo_ref, anchor="nw")
        fx0, fy0, fx1, fy1 = self._frame()
        for box in ((0, 0, _CW, fy0), (0, fy1, _CW, _CH), (0, fy0, fx0, fy1), (fx1, fy0, _CW, fy1)):
            c.create_rectangle(*box, fill="black", stipple="gray50", outline="")
        c.create_rectangle(fx0, fy0, fx1, fy1, outline="#ffffff", width=2)

    # --- взаимодействие -----------------------------------------------
    def _drag_begin(self, e):
        self._drag = (e.x, e.y, self.cx, self.cy)

    def _drag_move(self, e):
        if self._drag:
            x, y, cx, cy = self._drag
            self.cx, self.cy = cx + e.x - x, cy + e.y - y
            self._redraw()

    def _apply_zoom(self, new: float):
        old = self.zoom.get()
        new = min(max(new, 1.0), 5.0)
        x0, y0, x1, y1 = self._frame()
        fcx, fcy = (x0 + x1) / 2, (y0 + y1) / 2
        k = new / old
        self.cx, self.cy = fcx + (self.cx - fcx) * k, fcy + (self.cy - fcy) * k
        self.zoom.set(new)
        self._last_zoom = new
        self._redraw()

    def _zoom_by(self, factor: float):
        self._apply_zoom(self.zoom.get() * factor)

    def _on_zoom_slider(self, value):
        # слайдер уже записал новое значение в self.zoom; пересчитываем центр от прежнего
        new = float(value)
        old = self._last_zoom
        self._last_zoom = new
        x0, y0, x1, y1 = self._frame()
        fcx, fcy = (x0 + x1) / 2, (y0 + y1) / 2
        k = new / old if old else 1.0
        self.cx, self.cy = fcx + (self.cx - fcx) * k, fcy + (self.cy - fcy) * k
        self._redraw()

    # --- результат -----------------------------------------------------
    def _finish(self, crop: bool):
        try:
            max_side = int(self.max_side.get())
            if max_side < 50:
                raise ValueError
        except ValueError:
            messagebox.showerror("Фото", "Макс. сторона должна быть числом не меньше 50.", parent=self)
            return
        if crop:
            x0, y0, x1, y1 = self._frame()
            s = self._scale()
            ix0, iy0 = self.cx - self.img.width * s / 2, self.cy - self.img.height * s / 2
            out = self.img.crop((round((x0 - ix0) / s), round((y0 - iy0) / s),
                                 round((x1 - ix0) / s), round((y1 - iy0) / s)))
        else:
            out = self.img.copy()
        out.thumbnail((max_side, max_side), Image.LANCZOS)
        buf = io.BytesIO()
        out.save(buf, format="JPEG", quality=_JPEG_QUALITY)
        self.result = buf.getvalue()
        self.destroy()


def pick_and_crop(parent: tk.Misc, path: str, ratio: str, max_side: int = 800):
    """-> (jpeg-байты | None, выбранный формат)."""
    try:
        dlg = CropDialog(parent, path, ratio, max_side)
    except Exception as exc:
        messagebox.showerror("Фото", f"Не удалось открыть изображение:\n{exc}", parent=parent)
        return None, ratio
    parent.wait_window(dlg)
    return dlg.result, dlg.ratio.get()
