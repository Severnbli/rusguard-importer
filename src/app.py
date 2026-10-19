#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""RusGuard Importer — единое приложение: добавление ключей и создание профилей в RusGuard
(+ карточки в Access). Запуск: python app.py"""

from __future__ import annotations

import io
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from PIL import Image, ImageTk

import cropper
import db
import importer
import settings
import sheets
import theme
from keys import KEY_FORMATS
from sheets import Entry

_MODES = {
    sheets.MODE_ADD: "Добавить ключ существующему сотруднику",
    sheets.MODE_NEW: "Создать новый профиль",
}
_STATUS = {
    importer.OK: ("Готово", "#1a7f37"),
    importer.FAILED_TO_START: ("Не удалось начать — см. лог", "#b91c1c"),
    importer.PROBLEMS: ("Есть проблемы — см. лог", "#b45309"),
}
_IMAGE_TYPES = [("Изображения", "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp"), ("Все файлы", "*.*")]


class SearchDialog(tk.Toplevel):
    """Поиск сотрудников в БД по ФИО; result — список выбранных (id, ФИО, должность)."""

    def __init__(self, parent, cfg: dict):
        super().__init__(parent)
        self.title("Найти сотрудника в БД")
        self.cfg = cfg
        self.result: list = []
        self.found: list = []
        self.transient(parent)
        self.configure(background=theme.BG)

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        self.query = tk.StringVar()
        entry = ttk.Entry(top, textvariable=self.query, width=40)
        entry.pack(side="left", fill="x", expand=True)
        entry.bind("<Return>", lambda e: self._search())
        ttk.Button(top, text="Найти", command=self._search).pack(side="left", padx=(6, 0))

        self.listbox = tk.Listbox(self, selectmode="extended", width=70, height=14, relief="flat",
                                  highlightthickness=1, highlightbackground=theme.BORDER, activestyle="none")
        self.listbox.pack(fill="both", expand=True, padx=10)
        self.listbox.bind("<Double-Button-1>", lambda e: self._accept())

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Добавить выбранных", style="Accent.TButton", command=self._accept).pack(side="right")
        ttk.Button(bottom, text="Закрыть", command=self.destroy).pack(side="right", padx=6)
        entry.focus_set()
        self.grab_set()

    def _search(self):
        text = self.query.get().strip()
        if not text:
            return
        try:
            self.found = db.search_employees(self.cfg, text)
        except Exception as exc:
            messagebox.showerror("Поиск", str(exc), parent=self)
            return
        self.listbox.delete(0, "end")
        for _id, fio, position in self.found:
            self.listbox.insert("end", f"{fio}" + (f" — {position}" if position else ""))
        if not self.found:
            self.listbox.insert("end", "Ничего не найдено")

    def _accept(self):
        self.result = [self.found[i] for i in self.listbox.curselection() if i < len(self.found)]
        if self.result:
            self.destroy()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("RusGuard Importer")
        root.geometry("1120x860")
        root.minsize(980, 720)
        self.cfg = settings.load()
        self.queue: "queue.Queue" = queue.Queue()
        self.busy = False
        self.entries: list = []
        self.positions: list = []
        self.groups: list = []
        self._loading = False
        self._thumb = None
        self._prev_mode = sheets.MODE_NEW

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=12, pady=12)
        main_tab, setup_tab = ttk.Frame(notebook, padding=14), ttk.Frame(notebook, padding=18)
        notebook.add(main_tab, text="Импорт")
        notebook.add(setup_tab, text="Настройки")
        self._build_main(main_tab)
        self._build_setup(setup_tab)
        self._load_lookups_async()
        root.after(80, self._poll)

    # ------------------------------------------------------------------ main tab
    def _build_main(self, tab: ttk.Frame):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=3)
        tab.rowconfigure(5, weight=1)

        # --- шапка: режим и формат ключей
        head = ttk.Frame(tab)
        head.grid(row=0, column=0, sticky="we")
        self.mode = tk.StringVar(value=sheets.MODE_NEW)
        for value, text in _MODES.items():
            ttk.Radiobutton(head, text=text, value=value, variable=self.mode, style="Seg.Toolbutton",
                            command=self._on_mode_change).pack(side="left", padx=(0, 2))
        self.key_format = tk.StringVar(value=KEY_FORMATS.get(self.cfg["key_format"], KEY_FORMATS["dec"]))
        ttk.Combobox(head, textvariable=self.key_format, values=list(KEY_FORMATS.values()),
                     state="readonly", width=30).pack(side="right")
        ttk.Label(head, text="Формат ключей:", style="Muted.TLabel").pack(side="right", padx=(0, 8))

        # --- панель инструментов
        bar = ttk.Frame(tab)
        bar.grid(row=1, column=0, sticky="we", pady=(12, 6))
        self.btn_add = ttk.Button(bar, text="+ Добавить строку", command=self._on_add_row)
        self.btn_find = ttk.Button(bar, text="Найти в БД…", command=self._on_find)
        self.btn_add.pack(side="left")
        self.btn_find.pack(side="left")
        ttk.Button(bar, text="Удалить", style="Danger.TButton", command=self._on_delete).pack(
            side="left", padx=(6, 0))
        ttk.Button(bar, text="Обновить списки", command=lambda: self._load_lookups_async(True)).pack(
            side="right")
        ttk.Button(bar, text="Шаблон Excel…", command=self._on_save_template).pack(side="right", padx=6)
        ttk.Button(bar, text="Из Excel…", command=self._on_load_xlsx).pack(side="right")

        # --- таблица + карточка редактирования
        body = ttk.Frame(tab)
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        tree_card = ttk.Frame(body, style="Card.TFrame")
        tree_card.grid(row=0, column=0, sticky="nsew")
        cols = ("fio", "keys", "position", "group", "photo")
        self.tree = ttk.Treeview(tree_card, columns=cols, show="headings", selectmode="browse")
        for col, title, width, stretch in (("fio", "ФИО", 230, True), ("keys", "Ключи", 140, False),
                                           ("position", "Должность", 130, False),
                                           ("group", "Группа", 100, False), ("photo", "Фото", 60, False)):
            self.tree.heading(col, text=title, anchor="w")
            self.tree.column(col, width=width, anchor="w", stretch=stretch)
        self.tree.tag_configure("odd", background=theme.ZEBRA)
        sb = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._load_selected())

        panel = ttk.Frame(body, style="Card.TFrame", padding=14)
        panel.grid(row=0, column=1, sticky="ns", padx=(10, 0))
        ttk.Label(panel, text="Выбранная строка", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        photo_frame = ttk.Frame(panel, style="Photo.TFrame", width=150, height=200)
        photo_frame.pack(pady=(0, 6))
        photo_frame.pack_propagate(False)
        self.thumb_label = ttk.Label(photo_frame, text="нет фото", style="Photo.TLabel")
        self.thumb_label.pack(fill="both", expand=True)
        photo_btns = ttk.Frame(panel, style="Card.TFrame")
        photo_btns.pack(fill="x", pady=(0, 10))
        ttk.Button(photo_btns, text="Выбрать фото…", command=self._on_pick_photo).pack(fill="x")
        ttk.Button(photo_btns, text="Убрать фото", command=self._on_clear_photo).pack(fill="x", pady=(4, 0))

        self.v_name, self.v_keys = tk.StringVar(), tk.StringVar()
        self.v_pos, self.v_grp = tk.StringVar(), tk.StringVar()
        ttk.Label(panel, text="ФИО", style="Card.TLabel").pack(anchor="w")
        self.e_name = ttk.Entry(panel, textvariable=self.v_name, width=30)
        self.e_name.pack(fill="x", pady=(2, 8))
        ttk.Label(panel, text="Ключ(и) — через запятую", style="Card.TLabel").pack(anchor="w")
        ttk.Entry(panel, textvariable=self.v_keys).pack(fill="x", pady=(2, 8))
        ttk.Label(panel, text="Должность", style="Card.TLabel").pack(anchor="w")
        self.c_pos = ttk.Combobox(panel, textvariable=self.v_pos)
        self.c_pos.pack(fill="x", pady=(2, 8))
        ttk.Label(panel, text="Группа", style="Card.TLabel").pack(anchor="w")
        self.c_grp = ttk.Combobox(panel, textvariable=self.v_grp)
        self.c_grp.pack(fill="x", pady=(2, 0))
        for var in (self.v_name, self.v_keys, self.v_pos, self.v_grp):
            var.trace_add("write", lambda *a: self._on_edit())

        # --- параметры запуска
        opts = ttk.Frame(tab)
        opts.grid(row=3, column=0, sticky="we", pady=(12, 0))
        opts.columnconfigure(1, weight=1)
        self.dry_run = tk.BooleanVar(value=True)
        self.allow_partial = tk.BooleanVar()
        flags = ttk.Frame(opts)
        flags.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(flags, text="Dry-run (проверить, ничего не сохранять)",
                        variable=self.dry_run).pack(side="left")
        ttk.Checkbutton(flags, text="Сохранять валидные строки при ошибках в других",
                        variable=self.allow_partial).pack(side="left", padx=(18, 0))
        self.make_access = tk.BooleanVar()
        self.access_check = ttk.Checkbutton(opts, text="Создать Access-базу из шаблона", variable=self.make_access)
        self.access_check.grid(row=1, column=0, sticky="w", pady=(6, 0))
        if os.name != "nt":
            self.access_check.configure(text="Создать Access-базу из шаблона (только Windows)")
        self.access_out = tk.StringVar()
        self.access_label = ttk.Label(opts, text="Файл:", style="Muted.TLabel")
        self.access_entry = ttk.Entry(opts, textvariable=self.access_out)
        self.access_btn = ttk.Button(opts, text="Обзор…", command=self._on_pick_access)
        self.access_label.grid(row=1, column=1, sticky="e", padx=(12, 6), pady=(6, 0))
        self.access_entry.grid(row=1, column=2, sticky="we", pady=(6, 0))
        self.access_btn.grid(row=1, column=3, padx=(6, 0), pady=(6, 0))
        opts.columnconfigure(2, weight=1)

        # --- запуск и лог
        run_row = ttk.Frame(tab)
        run_row.grid(row=4, column=0, sticky="we", pady=12)
        self.run_button = ttk.Button(run_row, text="Импортировать", style="Accent.TButton",
                                     command=self._on_import)
        self.run_button.pack(side="left")
        self.status = tk.StringVar(value="Готово к запуску")
        self.status_label = ttk.Label(run_row, textvariable=self.status, style="Muted.TLabel")
        self.status_label.pack(side="left", padx=14)

        self.log_text = scrolledtext.ScrolledText(
            tab, height=7, wrap="word", state="disabled", relief="flat", borderwidth=0,
            background=theme.CARD, highlightthickness=1, highlightbackground=theme.BORDER,
            font=(theme.FONT_FAMILY if os.name == "nt" else "DejaVu Sans Mono", 9), padx=8, pady=6,
        )
        self.log_text.grid(row=5, column=0, sticky="nsew")
        self._refresh_mode()
        self._load_selected()

    def _refresh_mode(self):
        new = self.mode.get() == sheets.MODE_NEW
        self.btn_add.configure(state="normal" if new else "disabled")
        self.btn_find.configure(state="disabled" if new else "normal")
        access_ok = new and os.name == "nt"  # Access-драйвер ODBC есть только в Windows
        for w in (self.access_check, self.access_label, self.access_entry, self.access_btn):
            w.configure(state="normal" if access_ok else "disabled")
        if not access_ok:
            self.make_access.set(False)

    def _on_mode_change(self):
        if self.entries and not messagebox.askyesno(
                "Смена режима", "Таблица будет очищена. Продолжить?"):
            self.mode.set(self._prev_mode)
            return
        self._prev_mode = self.mode.get()
        self.entries = []
        self._rebuild_tree()
        self._refresh_mode()

    def _key_format_code(self) -> str:
        return next(k for k, v in KEY_FORMATS.items() if v == self.key_format.get())

    # ------------------------------------------------------------------ table
    def _rebuild_tree(self, select: int = None):
        self.tree.delete(*self.tree.get_children())
        for i, e in enumerate(self.entries):
            self.tree.insert("", "end", iid=str(i), values=self._tree_values(e),
                             tags=("odd",) if i % 2 else ())
        if select is not None and 0 <= select < len(self.entries):
            self.tree.selection_set(str(select))
            self.tree.see(str(select))
        self._load_selected()

    @staticmethod
    def _tree_values(e: Entry):
        return (e.fullname, e.keys, e.position, e.group, "да" if e.photo else "")

    def _selected_index(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _load_selected(self):
        idx = self._selected_index()
        self._loading = True
        try:
            e = self.entries[idx] if idx is not None else Entry()
            self.v_name.set(e.fullname)
            self.v_keys.set(e.keys)
            self.v_pos.set(e.position)
            self.v_grp.set(e.group)
            self._show_thumb(e.photo)
        finally:
            self._loading = False
        new = self.mode.get() == sheets.MODE_NEW
        has = idx is not None
        locked = has and self.entries[idx].employee_id is not None
        self.e_name.configure(state="disabled" if (not has or locked) else "normal")
        self.c_pos.configure(state="normal" if (has and new) else "disabled")
        self.c_grp.configure(state="normal" if (has and new) else "disabled")

    def _on_edit(self):
        if self._loading:
            return
        idx = self._selected_index()
        if idx is None:
            return
        e = self.entries[idx]
        e.fullname, e.keys = self.v_name.get(), self.v_keys.get()
        e.position, e.group = self.v_pos.get(), self.v_grp.get()
        self.tree.item(str(idx), values=self._tree_values(e))

    def _show_thumb(self, photo):
        if not photo:
            self._thumb = None
            self.thumb_label.configure(image="", text="нет фото")
            return
        im = Image.open(io.BytesIO(photo))
        im.thumbnail((150, 200))
        self._thumb = ImageTk.PhotoImage(im)
        self.thumb_label.configure(image=self._thumb, text="")

    def _on_add_row(self):
        self.entries.append(Entry())
        self._rebuild_tree(select=len(self.entries) - 1)
        self.e_name.focus_set()

    def _on_find(self):
        dlg = SearchDialog(self.root, self.cfg)
        self.root.wait_window(dlg)
        known = {e.employee_id for e in self.entries if e.employee_id}
        for emp_id, fio, _position in dlg.result:
            if emp_id not in known:
                self.entries.append(Entry(fullname=fio, employee_id=emp_id))
        if dlg.result:
            self._rebuild_tree(select=len(self.entries) - 1)

    def _on_delete(self):
        idx = self._selected_index()
        if idx is None:
            return
        del self.entries[idx]
        self._rebuild_tree(select=min(idx, len(self.entries) - 1))

    # ------------------------------------------------------------------ photo
    def _on_pick_photo(self):
        idx = self._selected_index()
        if idx is None:
            messagebox.showinfo("Фото", "Сначала выберите или добавьте строку.")
            return
        path = filedialog.askopenfilename(title="Фото", filetypes=_IMAGE_TYPES)
        if not path:
            return
        side = int(self.cfg["photo_max_side"]) if str(self.cfg["photo_max_side"]).isdigit() else 800
        data, ratio = cropper.pick_and_crop(self.root, path, self.cfg["photo_ratio"], side)
        if ratio != self.cfg["photo_ratio"]:
            self.cfg["photo_ratio"] = ratio
            try:
                settings.save(self.cfg)
            except OSError:
                pass
        if data:
            self.entries[idx].photo = data
            self.tree.item(str(idx), values=self._tree_values(self.entries[idx]))
            self._show_thumb(data)

    def _on_clear_photo(self):
        idx = self._selected_index()
        if idx is not None:
            self.entries[idx].photo = None
            self.tree.item(str(idx), values=self._tree_values(self.entries[idx]))
            self._show_thumb(None)

    # ------------------------------------------------------------------ excel
    def _on_load_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            loaded = sheets.load_xlsx_entries(path)
        except Exception as exc:
            messagebox.showerror("Excel", str(exc))
            return
        self.entries.extend(loaded)
        self._rebuild_tree(select=len(self.entries) - 1)
        self.status.set(f"Загружено строк из Excel: {len(loaded)}")
        self.status_label.configure(foreground="#374151")

    def _on_save_template(self):
        mode = self.mode.get()
        dest = filedialog.asksaveasfilename(
            title="Куда сохранить шаблон", defaultextension=".xlsx",
            initialfile="add_keys.xlsx" if mode == sheets.MODE_ADD else "new_profiles.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not dest:
            return
        try:
            sheets.create_from_template(mode, dest, self.positions, self.groups)
        except Exception as exc:
            messagebox.showerror("Не удалось создать файл", str(exc))
            return
        if not sheets.open_in_editor(dest):
            messagebox.showinfo("Шаблон", f"Файл создан: {dest}")

    def _on_pick_access(self):
        path = filedialog.asksaveasfilename(title="Файл Access", defaultextension=".accdb",
                                            filetypes=[("Access", "*.accdb")])
        if path:
            self.access_out.set(path)

    # ------------------------------------------------------------------ lookups
    def _load_lookups_async(self, show_errors: bool = False):
        if not self.cfg.get("dbname"):
            if show_errors:
                messagebox.showinfo("Списки", "Сначала заполните подключение на вкладке «Настройки».")
            return

        def worker():
            try:
                self.queue.put(("lookups", db.fetch_lookups(self.cfg)))
            except Exception as exc:
                self.queue.put(("lookups_error", (str(exc), show_errors)))

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------ setup tab
    def _build_setup(self, tab: ttk.Frame):
        tab.columnconfigure(0, weight=1)
        self.vars = {}

        def section(title, fields, row):
            box = ttk.LabelFrame(tab, text=title, padding=12)
            box.grid(row=row, column=0, sticky="we", pady=(0, 10))
            box.columnconfigure(1, weight=1)
            for r, (key, label) in enumerate(fields):
                self.vars[key] = tk.StringVar(value=str(self.cfg[key]))
                ttk.Label(box, text=label).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 12))
                ttk.Entry(box, textvariable=self.vars[key], show="*" if key == "password" else "",
                          width=36).grid(row=r, column=1, sticky="w")
            return box

        section("Подключение к БД RusGuard", [
            ("host", "Хост"), ("port", "Порт"), ("dbname", "База данных"),
            ("user", "Пользователь"), ("password", "Пароль"),
        ], 0)
        section("Фото", [("photo_max_side", "Макс. сторона, px")], 1)
        access = section("Access", [
            ("access_table", "Таблица"), ("access_max_photo_dimension", "Макс. размер фото, px"),
        ], 2)
        self.vars["access_position_case"] = tk.StringVar(value=self.cfg["access_position_case"])
        ttk.Label(access, text="Регистр должности").grid(row=2, column=0, sticky="w", pady=4, padx=(0, 12))
        ttk.Combobox(access, textvariable=self.vars["access_position_case"], values=["lower", "asis"],
                     state="readonly", width=10).grid(row=2, column=1, sticky="w")

        keys_box = ttk.LabelFrame(tab, text="Ключи", padding=12)
        keys_box.grid(row=3, column=0, sticky="we", pady=(0, 10))
        self.default_format = tk.StringVar(value=KEY_FORMATS.get(self.cfg["key_format"], KEY_FORMATS["dec"]))
        ttk.Label(keys_box, text="Формат по умолчанию").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Combobox(keys_box, textvariable=self.default_format, values=list(KEY_FORMATS.values()),
                     state="readonly", width=32).grid(row=0, column=1, sticky="w")

        row = ttk.Frame(tab)
        row.grid(row=4, column=0, sticky="w")
        ttk.Button(row, text="Сохранить", style="Accent.TButton", command=self._on_save_settings).pack(side="left")
        ttk.Button(row, text="Проверить подключение", command=self._on_test).pack(side="left", padx=8)
        ttk.Label(tab, text=f"Файл настроек: {settings.settings_path()}", style="Muted.TLabel").grid(
            row=5, column=0, sticky="w", pady=(10, 0))

    def _collect_settings(self) -> dict:
        cfg = dict(self.cfg)
        cfg.update({k: v.get().strip() if k != "password" else v.get() for k, v in self.vars.items()})
        cfg["key_format"] = next(k for k, v in KEY_FORMATS.items() if v == self.default_format.get())
        return cfg

    def _on_save_settings(self):
        self.cfg = self._collect_settings()
        try:
            settings.save(self.cfg)
        except OSError as exc:
            messagebox.showerror("Не удалось сохранить", str(exc))
            return
        messagebox.showinfo("Настройки", "Сохранено.")
        self._load_lookups_async()

    def _on_test(self):
        try:
            positions, groups = db.fetch_lookups(self._collect_settings())
        except Exception as exc:
            messagebox.showerror("Подключение", str(exc))
            return
        messagebox.showinfo("Подключение", f"Успешно. Должностей: {len(positions)}, групп: {len(groups)}.")

    # ------------------------------------------------------------------ import
    def _on_import(self):
        if self.busy:
            return
        entries = [Entry(e.fullname, e.keys, e.position, e.group, e.photo, e.employee_id)
                   for e in self.entries]
        if not entries:
            messagebox.showerror("Проверьте таблицу", "Таблица пуста.")
            return
        access_out = ""
        if self.mode.get() == sheets.MODE_NEW and self.make_access.get():
            access_out = self.access_out.get().strip()
            if not access_out:
                messagebox.showerror("Проверьте поля", "Укажите, куда сохранить Access-базу.")
                return
        args = (self.cfg, self.mode.get(), entries, self._key_format_code(), access_out or None,
                self.dry_run.get(), self.allow_partial.get())

        self.busy = True
        self.run_button.configure(state="disabled")
        self.status.set("Выполняется…")
        self.status_label.configure(foreground="#374151")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        def worker():
            try:
                code = importer.run_import(*args, log=lambda m: self.queue.put(("line", m)))
            except Exception as exc:
                self.queue.put(("line", f"НЕОЖИДАННАЯ ОШИБКА: {exc}"))
                code = importer.FAILED_TO_START
            self.queue.put(("done", code))

        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "line":
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", payload + "\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "done":
                    self.busy = False
                    self.run_button.configure(state="normal")
                    text, color = _STATUS.get(payload, ("Завершено", "#374151"))
                    self.status.set(text)
                    self.status_label.configure(foreground=color)
                elif kind == "lookups":
                    self.positions, self.groups = payload
                    self.c_pos.configure(values=sorted({n for _, n in self.positions if n}))
                    self.c_grp.configure(values=sorted({n for _, n in self.groups if n}))
                elif kind == "lookups_error":
                    message, show = payload
                    self.status.set(f"Не удалось получить должности/группы: {message}")
                    self.status_label.configure(foreground="#b45309")
                    if show:
                        messagebox.showerror("Списки", message)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)


def main():
    root = tk.Tk()
    theme.apply(root)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
