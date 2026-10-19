# -*- coding: utf-8 -*-
"""Окно пароля приложения: установка при первом запуске, вход, сброс данных при забытом пароле."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

import settings
import theme


class PasswordDialog(tk.Toplevel):
    def __init__(self, root: tk.Tk, setup: bool):
        super().__init__(root)
        self.setup = setup
        self.action = "exit"  # ok | reset | exit
        self.title("RusGuard Importer")
        self.configure(background=theme.BG)
        self.resizable(False, False)
        pad = theme.px

        body = ttk.Frame(self, padding=pad(24))
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Задайте пароль приложения" if setup else "Введите пароль",
                  font=theme.font(3, True)).pack(anchor="w")
        if setup:
            ttk.Label(body, text="Пароль защищает сохранённые настройки подключения. "
                                 "Восстановить его нельзя: при утере все сохранённые данные будут стёрты.",
                      style="Muted.TLabel", wraplength=pad(380)).pack(anchor="w", pady=(6, 12))
        else:
            ttk.Label(body, text="").pack(pady=pad(4))

        self.pw = tk.StringVar()
        self.pw2 = tk.StringVar()
        ttk.Label(body, text="Пароль").pack(anchor="w")
        first = ttk.Entry(body, textvariable=self.pw, show="*", width=34)
        first.pack(fill="x", pady=(2, pad(10)))
        first.bind("<Return>", lambda e: self._submit())
        if setup:
            ttk.Label(body, text="Повторите пароль").pack(anchor="w")
            second = ttk.Entry(body, textvariable=self.pw2, show="*")
            second.pack(fill="x", pady=(2, pad(10)))
            second.bind("<Return>", lambda e: self._submit())

        self.error = tk.StringVar()
        ttk.Label(body, textvariable=self.error, foreground=theme.ERR, wraplength=pad(380)).pack(anchor="w")

        row = ttk.Frame(body)
        row.pack(fill="x", pady=(pad(12), 0))
        ttk.Button(row, text="Задать пароль" if setup else "Войти", style="Accent.TButton",
                   command=self._submit).pack(side="left")
        ttk.Button(row, text="Выход", command=self.destroy).pack(side="left", padx=8)
        if not setup:
            ttk.Button(row, text="Забыл пароль", command=self._forgot).pack(side="right")

        first.focus_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{x}+{y}")
        self.wait_visibility()
        self.grab_set()

    def _submit(self):
        password = self.pw.get()
        if self.setup:
            if len(password) < settings.MIN_PASSWORD_LENGTH:
                self.error.set(f"Пароль должен быть не короче {settings.MIN_PASSWORD_LENGTH} символов.")
                return
            if password != self.pw2.get():
                self.error.set("Пароли не совпадают.")
                return
            settings.set_app_password(password)
        elif not settings.unlock(password):
            self.error.set("Неверный пароль.")
            self.pw.set("")
            return
        self.action = "ok"
        self.destroy()

    def _forgot(self):
        if messagebox.askyesno(
            "Забыл пароль",
            "Все сохранённые данные (подключение к БД, пароль БД, настройки) будут удалены, "
            "а пароль приложения нужно будет задать заново. Продолжить?",
            icon="warning", parent=self,
        ):
            self.action = "reset"
            self.destroy()


def ensure_unlocked(root: tk.Tk) -> bool:
    """Показывает окно пароля до открытия приложения. False — пользователь вышел."""
    while True:
        dlg = PasswordDialog(root, setup=not settings.has_app_password())
        root.wait_window(dlg)
        if dlg.action == "ok":
            return True
        if dlg.action == "reset":
            settings.reset()
            continue
        return False
