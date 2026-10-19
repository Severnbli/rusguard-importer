# -*- coding: utf-8 -*-
"""Единое оформление приложения (ttk, тема clam) с масштабированием интерфейса."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageDraw, ImageTk

BG = "#f3f4f6"
CARD = "#ffffff"
FG = "#111827"
MUTED = "#6b7280"
BORDER = "#d1d5db"
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
OK, WARN, ERR = "#15803d", "#b45309", "#b91c1c"
ZEBRA = "#f9fafb"

FONT_FAMILY = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
MONO_FAMILY = "Consolas" if os.name == "nt" else "DejaVu Sans Mono"
SCALE = 1.0
_images: list = []  # ссылки на PhotoImage, иначе их удалит сборщик мусора
BASE_SIZE = 11


def px(value: float) -> int:
    """Размер в пикселях с учётом масштаба интерфейса."""
    return round(value * SCALE)


def font(delta: int = 0, bold: bool = False) -> tuple:
    size = max(7, round((BASE_SIZE + delta) * SCALE))
    return (FONT_FAMILY, size, "bold") if bold else (FONT_FAMILY, size)


def _checkbox_image(size: int, checked: bool, disabled: bool) -> tk.PhotoImage:
    k = 4  # рисуем в увеличенном размере и сжимаем — сглаживание краёв
    big = size * k
    im = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    fill = "#e5e7eb" if disabled else (ACCENT if checked else CARD)
    border = "#d1d5db" if disabled else (ACCENT if checked else "#9ca3af")
    d.rounded_rectangle((k, k, big - k - 1, big - k - 1), radius=big // 5, fill=fill, outline=border, width=2 * k)
    if checked:
        w = max(2, big // 9)
        pts = [(big * 0.24, big * 0.52), (big * 0.43, big * 0.70), (big * 0.77, big * 0.30)]
        d.line(pts, fill="#ffffff", width=w, joint="curve")
    photo = ImageTk.PhotoImage(im.resize((size, size), Image.LANCZOS))
    _images.append(photo)
    return photo


def _install_checkbutton(style: ttk.Style) -> None:
    size = px(24)
    off, on = _checkbox_image(size, False, False), _checkbox_image(size, True, False)
    off_dis, on_dis = _checkbox_image(size, False, True), _checkbox_image(size, True, True)
    style.element_create(
        "Big.Check.indicator", "image", str(off),
        ("selected", "disabled", str(on_dis)), ("selected", str(on)), ("disabled", str(off_dis)),
        sticky="w",
    )
    style.layout("TCheckbutton", [
        ("Checkbutton.padding", {"sticky": "nswe", "children": [
            ("Big.Check.indicator", {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                ("Checkbutton.label", {"sticky": "nswe"}),
            ]}),
        ]}),
    ])


def apply(root: tk.Tk, scale_percent: int = 100) -> None:
    global SCALE
    SCALE = min(max(scale_percent, 80), 250) / 100
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    root.configure(background=BG)
    root.option_add("*Font", font())
    root.option_add("*TCombobox*Listbox.font", font())
    root.option_add("*Listbox.background", CARD)
    root.option_add("*Listbox.selectBackground", ACCENT)

    style.configure(".", background=BG, foreground=FG, font=font(), bordercolor=BORDER,
                    lightcolor=BG, darkcolor=BG, focuscolor=BG)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1)
    style.configure("Card.TLabel", background=CARD)
    style.configure("Panel.TFrame", background=CARD)
    style.configure("Card.TCheckbutton", background=CARD)
    style.map("Card.TCheckbutton", background=[("active", CARD)])
    style.configure("CardTitle.TLabel", background=CARD, font=font(1, True))
    style.configure("Photo.TFrame", background="#e5e7eb")
    style.configure("Photo.TLabel", background="#e5e7eb", foreground=MUTED, anchor="center")
    _install_checkbutton(style)
    style.configure("TCheckbutton", background=BG, padding=px(4))
    style.map("TCheckbutton", background=[("active", BG)])
    style.configure("TLabelframe", background=BG, bordercolor=BORDER)
    style.configure("TLabelframe.Label", background=BG, foreground=MUTED, font=font(0, True))

    style.configure("TButton", padding=(px(16), px(10)), background="#e5e7eb", foreground=FG, relief="flat")
    style.map("TButton", background=[("active", "#d1d5db"), ("disabled", "#f3f4f6")],
              foreground=[("disabled", "#9ca3af")])
    style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", font=font(1, True),
                    padding=(px(28), px(12)))
    style.map("Accent.TButton", background=[("active", ACCENT_DARK), ("disabled", "#93c5fd")],
              foreground=[("disabled", "#ffffff")])
    style.configure("Danger.TButton", foreground=ERR)

    style.configure("Seg.Toolbutton", padding=(px(20), px(12)), background="#e5e7eb", foreground=FG,
                    relief="flat", font=font(1))
    style.map("Seg.Toolbutton", background=[("selected", ACCENT), ("active", "#d1d5db")],
              foreground=[("selected", "#ffffff")])

    style.configure("TEntry", padding=px(9), fieldbackground=CARD)
    style.configure("TCombobox", padding=px(9), fieldbackground=CARD, arrowsize=px(18))
    style.map("TCombobox", fieldbackground=[("readonly", CARD), ("disabled", "#f3f4f6")])
    style.map("TEntry", fieldbackground=[("disabled", "#f3f4f6")])
    style.configure("TScrollbar", arrowsize=px(16))

    style.configure("Treeview", rowheight=px(38), background=CARD, fieldbackground=CARD, borderwidth=0)
    style.configure("Treeview.Heading", font=font(0, True), background="#e5e7eb",
                    padding=(px(10), px(8)), relief="flat")
    style.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading", background=[("active", "#d1d5db")])

    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(px(22), px(10)), background="#e5e7eb", font=font(1))
    style.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", ACCENT)])
    style.configure("TScale", background=CARD)
    style.configure("TPanedwindow", background=BG)
    style.configure("Sash", sashthickness=px(8), gripcount=0, background=BG)
