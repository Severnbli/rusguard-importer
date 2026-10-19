# -*- coding: utf-8 -*-
"""Единое оформление приложения (ttk, тема clam)."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk

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
FONT = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")


def apply(root: tk.Tk) -> None:
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    root.configure(background=BG)
    root.option_add("*Font", FONT)
    root.option_add("*TCombobox*Listbox.font", FONT)
    root.option_add("*Listbox.background", CARD)
    root.option_add("*Listbox.selectBackground", ACCENT)

    style.configure(".", background=BG, foreground=FG, font=FONT, bordercolor=BORDER,
                    lightcolor=BG, darkcolor=BG, focuscolor=BG)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("Title.TLabel", font=(FONT_FAMILY, 16, "bold"))
    style.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1)
    style.configure("Card.TLabel", background=CARD)
    style.configure("CardTitle.TLabel", background=CARD, font=FONT_BOLD)
    style.configure("Photo.TFrame", background="#e5e7eb")
    style.configure("Photo.TLabel", background="#e5e7eb", foreground=MUTED, anchor="center")
    style.configure("TCheckbutton", background=BG)
    style.map("TCheckbutton", background=[("active", BG)])
    style.configure("TLabelframe", background=BG, bordercolor=BORDER)
    style.configure("TLabelframe.Label", background=BG, foreground=MUTED, font=FONT_BOLD)

    style.configure("TButton", padding=(12, 6), background="#e5e7eb", foreground=FG, relief="flat")
    style.map("TButton", background=[("active", "#d1d5db"), ("disabled", "#f3f4f6")],
              foreground=[("disabled", "#9ca3af")])
    style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", font=FONT_BOLD,
                    padding=(22, 9))
    style.map("Accent.TButton", background=[("active", ACCENT_DARK), ("disabled", "#93c5fd")],
              foreground=[("disabled", "#ffffff")])
    style.configure("Danger.TButton", foreground=ERR)

    # переключатель режима в виде сегментных кнопок
    style.configure("Seg.Toolbutton", padding=(16, 8), background="#e5e7eb", foreground=FG,
                    relief="flat", font=FONT)
    style.map("Seg.Toolbutton", background=[("selected", ACCENT), ("active", "#d1d5db")],
              foreground=[("selected", "#ffffff")])

    style.configure("TEntry", padding=6, fieldbackground=CARD)
    style.configure("TCombobox", padding=6, fieldbackground=CARD, arrowsize=14)
    style.map("TCombobox", fieldbackground=[("readonly", CARD), ("disabled", "#f3f4f6")])
    style.map("TEntry", fieldbackground=[("disabled", "#f3f4f6")])

    style.configure("Treeview", rowheight=30, background=CARD, fieldbackground=CARD, borderwidth=0)
    style.configure("Treeview.Heading", font=FONT_BOLD, background="#e5e7eb", padding=(8, 6), relief="flat")
    style.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading", background=[("active", "#d1d5db")])

    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(18, 8), background="#e5e7eb", font=FONT)
    style.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", ACCENT)])
    style.configure("TScale", background=CARD)
