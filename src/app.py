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
from keys import KEY_FORMATS, parse_keys
from sheets import Entry

_MODES = {
    sheets.MODE_ADD: "Редактирование профиля",
    sheets.MODE_NEW: "Создание профиля",
}
_STATUS = {
    importer.OK: ("Готово", "#1a7f37"),
    importer.FAILED_TO_START: ("Не удалось начать — см. лог", "#b91c1c"),
    importer.PROBLEMS: ("Есть проблемы — см. лог", "#b45309"),
}
_IMAGE_TYPES = [("Изображения", "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp"), ("Все файлы", "*.*")]


class SearchDialog(tk.Toplevel):
    """Поиск сотрудников в БД по ФИО или номеру ключа; result — выбранные (id, ФИО, должность, ключи)."""

    def __init__(self, parent, cfg: dict, key_format: str):
        super().__init__(parent)
        self.title("Найти сотрудника в БД")
        self.cfg = cfg
        self.key_format = key_format
        self.result: list = []
        self.found: list = []
        self.transient(parent)
        self.configure(background=theme.BG)

        top = ttk.Frame(self, padding=(10, 10, 10, 0))
        top.pack(fill="x")
        self.query = tk.StringVar()
        entry = ttk.Entry(top, textvariable=self.query, width=40)
        entry.pack(side="left", fill="x", expand=True)
        entry.bind("<Return>", lambda e: self._search())
        ttk.Button(top, text="Найти", command=self._search).pack(side="left", padx=(6, 0))

        ttk.Label(self, text="ФИО (часть слов) или номер ключа в выбранном формате", style="Muted.TLabel").pack(
            anchor="w", padx=10, pady=(4, 8))
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
            self.found = db.search_employees(self.cfg, text, self._key_candidates(text))
        except Exception as exc:
            messagebox.showerror("Поиск", str(exc), parent=self)
            return
        self.listbox.delete(0, "end")
        for _id, fio, position, keys in self.found:
            line = fio + (f" — {position}" if position else "")
            if keys:
                line += "   [ключи: " + ", ".join(str(k) for k in keys) + "]"
            self.listbox.insert("end", line)
        if not self.found:
            self.listbox.insert("end", "Ничего не найдено")

    def _key_candidates(self, text: str) -> set:
        numbers: set = set()
        for fmt in {self.key_format, "dec"}:
            parsed, errors = parse_keys(text, fmt)
            if not errors:
                numbers.update(parsed)
        return numbers

    def _accept(self):
        self.result = [self.found[i] for i in self.listbox.curselection() if i < len(self.found)]
        if self.result:
            self.destroy()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("RusGuard Importer")
        root.geometry(f"{theme.px(1280)}x{theme.px(900)}")
        root.minsize(theme.px(1000), theme.px(720))
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
        tab.rowconfigure(2, weight=1)
        pad = theme.px

        # --- шапка: режим и формат ключей
        head = ttk.Frame(tab)
        head.grid(row=0, column=0, sticky="we")
        self.mode = tk.StringVar(value=sheets.MODE_NEW)
        for value, text in _MODES.items():
            ttk.Radiobutton(head, text=text, value=value, variable=self.mode, style="Seg.Toolbutton",
                            command=self._on_mode_change).pack(side="left", padx=(0, 2))
        self.key_format = tk.StringVar(value=KEY_FORMATS.get(self.cfg["key_format"], KEY_FORMATS["dec"]))
        ttk.Combobox(head, textvariable=self.key_format, values=list(KEY_FORMATS.values()),
                     state="readonly", width=26).pack(side="right")
        ttk.Label(head, text="Формат ключей:", style="Muted.TLabel").pack(side="right", padx=(0, 8))

        # --- панель инструментов
        bar = ttk.Frame(tab)
        bar.grid(row=1, column=0, sticky="we", pady=(pad(12), pad(8)))
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

        # --- вертикальный разделитель: сверху таблица+карточка, снизу запуск и лог
        vpane = ttk.PanedWindow(tab, orient="vertical")
        vpane.grid(row=2, column=0, sticky="nsew")
        hpane = ttk.PanedWindow(vpane, orient="horizontal")
        vpane.add(hpane, weight=4)

        tree_card = ttk.Frame(hpane, style="Card.TFrame")
        tree_card.columnconfigure(0, weight=1)
        tree_card.rowconfigure(0, weight=1)
        hpane.add(tree_card, weight=3)
        cols = ("fio", "keys", "position", "group", "photo")
        self.tree = ttk.Treeview(tree_card, columns=cols, show="headings", selectmode="browse")
        for col, title in (("fio", "ФИО"), ("keys", "Ключи"), ("position", "Должность"),
                           ("group", "Группа"), ("photo", "Фото")):
            self.tree.heading(col, text=title, anchor="w")
            self.tree.column(col, width=pad(80), minwidth=20, anchor="w", stretch=False)
        self.tree.tag_configure("odd", background=theme.ZEBRA)
        sb = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._load_selected())
        self.tree.bind("<Configure>", self._fit_columns)
        for seq in ("<Button-1>", "<Double-Button-1>"):
            self.tree.bind(seq, self._block_resize)

        panel = ttk.Frame(hpane, style="Card.TFrame", padding=pad(16))
        panel.columnconfigure(0, weight=1)
        hpane.add(panel, weight=1)
        self.panel_title = tk.StringVar()
        ttk.Label(panel, textvariable=self.panel_title, style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, pad(10)))

        photo_frame = ttk.Frame(panel, style="Photo.TFrame", width=pad(180), height=pad(240))
        photo_frame.grid(row=1, column=0, pady=(0, pad(8)))
        photo_frame.grid_propagate(False)
        photo_frame.pack_propagate(False)
        self.thumb_label = ttk.Label(photo_frame, text="нет фото", style="Photo.TLabel")
        self.thumb_label.pack(fill="both", expand=True)
        photo_btns = ttk.Frame(panel, style="Card.TFrame")
        photo_btns.grid(row=2, column=0, sticky="we", pady=(0, pad(12)))
        ttk.Button(photo_btns, text="Выбрать фото…", command=self._on_pick_photo).pack(fill="x")
        self.btn_clear = ttk.Button(photo_btns, text="Убрать фото", command=self._on_clear_photo)
        self.btn_clear.pack(fill="x", pady=(4, 0))

        self.v_name, self.v_keys = tk.StringVar(), tk.StringVar()
        self.v_pos, self.v_grp = tk.StringVar(), tk.StringVar()
        field_pad = (2, pad(10))
        ttk.Label(panel, text="ФИО", style="Card.TLabel").grid(row=3, column=0, sticky="w")
        self.e_name = ttk.Entry(panel, textvariable=self.v_name)
        self.e_name.grid(row=4, column=0, sticky="we", pady=field_pad)
        self.lbl_keys = ttk.Label(panel, style="Card.TLabel")
        self.lbl_keys.grid(row=5, column=0, sticky="w")
        self.e_keys = ttk.Entry(panel, textvariable=self.v_keys)
        self.e_keys.grid(row=6, column=0, sticky="we", pady=field_pad)
        self.lbl_cur_keys = ttk.Label(panel, style="Card.TLabel", foreground=theme.MUTED, wraplength=pad(300))
        self.lbl_cur_keys.grid(row=7, column=0, sticky="w", pady=(0, pad(10)))
        ttk.Label(panel, text="Должность", style="Card.TLabel").grid(row=8, column=0, sticky="w")
        self.c_pos = ttk.Combobox(panel, textvariable=self.v_pos)
        self.c_pos.grid(row=9, column=0, sticky="we", pady=field_pad)
        ttk.Label(panel, text="Группа", style="Card.TLabel").grid(row=10, column=0, sticky="w")
        self.c_grp = ttk.Combobox(panel, textvariable=self.v_grp)
        self.c_grp.grid(row=11, column=0, sticky="we", pady=(2, 0))
        for var in (self.v_name, self.v_keys, self.v_pos, self.v_grp):
            var.trace_add("write", lambda *a: self._on_edit())

        # --- нижняя часть: параметры, запуск, лог
        bottom = ttk.Frame(vpane)
        bottom.columnconfigure(0, weight=1)
        bottom.rowconfigure(2, weight=1)
        vpane.add(bottom, weight=1)

        opts = ttk.Frame(bottom)
        opts.grid(row=0, column=0, sticky="we", pady=(pad(10), 0))
        self.dry_run = tk.BooleanVar(value=True)
        self.allow_partial = tk.BooleanVar()
        ttk.Checkbutton(opts, text="Dry-run (проверить, ничего не сохранять)",
                        variable=self.dry_run).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(opts, text="Сохранять валидные строки при ошибках в других",
                        variable=self.allow_partial).grid(row=0, column=1, sticky="w", padx=(pad(18), 0))
        self.make_access = tk.BooleanVar()
        self.access_check = ttk.Checkbutton(opts, text="Создать Access-базу из шаблона", variable=self.make_access)
        self.access_check.grid(row=1, column=0, sticky="w")
        if os.name != "nt":
            self.access_check.configure(text="Создать Access-базу из шаблона (только Windows)")
        self.access_out = tk.StringVar()
        access_row = ttk.Frame(opts)
        access_row.grid(row=2, column=0, columnspan=2, sticky="we", pady=(2, 0))
        access_row.columnconfigure(1, weight=1)
        self.access_label = ttk.Label(access_row, text="Файл Access:", style="Muted.TLabel")
        self.access_entry = ttk.Entry(access_row, textvariable=self.access_out)
        self.access_btn = ttk.Button(access_row, text="Обзор…", command=self._on_pick_access)
        self.access_label.grid(row=0, column=0, padx=(0, 8))
        self.access_entry.grid(row=0, column=1, sticky="we")
        self.access_btn.grid(row=0, column=2, padx=(8, 0))
        opts.columnconfigure(1, weight=1)

        run_row = ttk.Frame(bottom)
        run_row.grid(row=1, column=0, sticky="we", pady=pad(10))
        self.run_button = ttk.Button(run_row, text="Импортировать", style="Accent.TButton",
                                     command=self._on_import)
        self.run_button.pack(side="left")
        self.status = tk.StringVar(value="Готово к запуску")
        self.status_label = ttk.Label(run_row, textvariable=self.status, style="Muted.TLabel")
        self.status_label.pack(side="left", padx=14)

        self.log_text = scrolledtext.ScrolledText(
            bottom, height=5, wrap="word", state="disabled", relief="flat", borderwidth=0,
            background=theme.CARD, highlightthickness=1, highlightbackground=theme.BORDER,
            font=(theme.MONO_FAMILY, max(8, theme.font(-1)[1])), padx=8, pady=6,
        )
        self.log_text.grid(row=2, column=0, sticky="nsew")
        self._refresh_mode()
        self._load_selected()

    # ширина столбцов — фиксированные доли ширины таблицы
    _COL_PERCENT = (("fio", 0.38), ("keys", 0.22), ("position", 0.17), ("group", 0.15), ("photo", 0.08))

    def _fit_columns(self, event):
        total = max(event.width - 4, 300)
        for col, share in self._COL_PERCENT:
            self.tree.column(col, width=int(total * share))

    def _block_resize(self, event):
        if self.tree.identify_region(event.x, event.y) == "separator":
            return "break"

    def _refresh_mode(self):
        new = self.mode.get() == sheets.MODE_NEW
        self.btn_add.configure(state="normal" if new else "disabled")
        self.btn_find.configure(state="disabled" if new else "normal")
        self.panel_title.set("Новый профиль" if new else "Редактирование профиля")
        self.lbl_keys.configure(text="Ключ(и) — через запятую" if new else "Добавить ключ(и) — через запятую")
        self.btn_clear.configure(text="Убрать фото" if new else "Вернуть исходное фото")
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
        entry = self.entries[idx] if has else None
        locked = has and entry.employee_id is not None
        self.e_name.configure(state="disabled" if (not has or locked) else "normal")
        self.e_keys.configure(state="normal" if has else "disabled")
        self.c_pos.configure(state="normal" if has else "disabled")
        self.c_grp.configure(state="normal" if has else "disabled")
        if not new and has and entry.current_keys:
            self.lbl_cur_keys.configure(
                text="Текущие ключи: " + ", ".join(str(k) for k in entry.current_keys))
            self.lbl_cur_keys.grid()
        elif not new:
            self.lbl_cur_keys.configure(text="Текущих ключей нет" if has else "")
            self.lbl_cur_keys.grid()
        else:
            self.lbl_cur_keys.grid_remove()

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
        try:
            im = Image.open(io.BytesIO(photo))
            im.thumbnail((theme.px(180), theme.px(240)))
            self._thumb = ImageTk.PhotoImage(im)
        except Exception:
            self._thumb = None
            self.thumb_label.configure(image="", text="фото не читается")
            return
        self.thumb_label.configure(image=self._thumb, text="")

    def _on_add_row(self):
        self.entries.append(Entry())
        self._rebuild_tree(select=len(self.entries) - 1)
        self.e_name.focus_set()

    def _on_find(self):
        dlg = SearchDialog(self.root, self.cfg, self._key_format_code())
        self.root.wait_window(dlg)
        known = {e.employee_id for e in self.entries if e.employee_id}
        chosen = [(i, fio) for i, fio, _p, _k in dlg.result if i not in known]
        if not chosen:
            return
        try:
            profiles = db.fetch_profiles(self.cfg, [i for i, _ in chosen])
        except Exception as exc:
            messagebox.showerror("Профили", f"Не удалось загрузить данные сотрудников:\n{exc}")
            return
        for emp_id, fio in chosen:
            p = profiles[emp_id]
            self.entries.append(Entry(
                fullname=fio, employee_id=emp_id, photo=p["photo"], photo_original=p["photo"],
                position=p["position"], position_original=p["position"],
                group=p["group"], group_original=p["group"], current_keys=p["keys"],
            ))
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
            e = self.entries[idx]
            e.photo = e.photo_original  # у нового профиля исходного нет — фото убирается
            self.tree.item(str(idx), values=self._tree_values(e))
            self._show_thumb(e.photo)

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
        section("Фото и интерфейс", [
            ("photo_max_side", "Макс. сторона фото, px"),
            ("ui_scale", "Масштаб интерфейса, % (после перезапуска)"),
        ], 1)
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
        entries = list(self.entries)
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
    scale = str(settings.load().get("ui_scale", "100")).strip()
    theme.apply(root, int(scale) if scale.isdigit() else 100)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
