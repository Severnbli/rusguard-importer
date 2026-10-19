#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui_common.py — общая часть графических обёрток для rusguard.py и
access.py.

Это НЕ импортирует сами скрипты и не дублирует их логику: GUI просто
собирает аргументы командной строки и запускает нужный .py как отдельный
процесс (через тот же интерпретатор Python, что запустил GUI), построчно
показывая его вывод в окне. Сами скрипты остаются полностью
самостоятельными CLI-инструментами — GUI ничего в их поведении не меняет.

Пароль (для rusguard_gui.py) передаётся дочернему процессу через
переменную окружения PGPASSWORD, а не аргументом командной строки — так
он не виден в списке процессов (Диспетчер задач / ps).
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Callable, Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

_EXIT_CODE_STATUS = {
    0: ("Готово: сохранено без проблем", "#1a7f37"),
    1: ("Не удалось начать — см. лог", "#b91c1c"),
    2: ("Есть проблемы / ничего не сохранено — см. лог", "#b45309"),
}


def default_status_for_code(code: Optional[int]):
    if code is None:
        return ("Выполняется…", "#374151")
    if code in _EXIT_CODE_STATUS:
        return _EXIT_CODE_STATUS[code]
    return (f"Завершилось с кодом {code} — см. лог", "#b91c1c")


class LogPanel:
    """Прокручиваемая текстовая область + потоковый запуск subprocess.
    Вся работа с виджетом идёт из основного потока Tk через очередь —
    из фонового потока (чтение stdout процесса) в виджет напрямую не
    пишем, это небезопасно в Tkinter."""

    def __init__(self, parent: tk.Widget, root: tk.Tk, height: int = 22):
        self.root = root
        self.frame = ttk.Frame(parent)
        self.text = scrolledtext.ScrolledText(
            self.frame, height=height, wrap="word", state="disabled",
            font=("Consolas", 9) if os.name == "nt" else ("Monospace", 9),
        )
        self.text.pack(fill="both", expand=True)
        self._queue: "queue.Queue[tuple]" = queue.Queue()
        self._proc: Optional[subprocess.Popen] = None
        self._running = False
        self._on_done: Optional[Callable[[int], None]] = None
        self.root.after(80, self._poll)

    # ------------------------------------------------------------------
    def clear(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def get_text(self) -> str:
        return self.text.get("1.0", "end")

    def _append(self, line: str):
        self.text.configure(state="normal")
        self.text.insert("end", line)
        self.text.see("end")
        self.text.configure(state="disabled")

    def _poll(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "line":
                    self._append(payload)
                elif kind == "done":
                    self._proc = None
                    self._running = False
                    if self._on_done:
                        self._on_done(payload)
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    # ------------------------------------------------------------------
    def is_running(self) -> bool:
        return self._running

    def stop(self):
        """Принудительно остановить запущенный процесс (если есть)."""
        if self._proc is not None:
            try:
                self._proc.terminate()
            except OSError:
                pass

    def run(self, cmd: list, env: dict, on_done: Callable[[int], None]):
        """Запускает cmd как процесс, стримит stdout+stderr в текстовое
        поле построчно, вызывает on_done(returncode) по завершении
        (уже в основном потоке Tk)."""
        if self.is_running():
            raise RuntimeError("процесс уже выполняется")
        self.clear()
        self._running = True
        self._on_done = on_done
        self._append("$ " + " ".join(_quote_for_display(c) for c in cmd) + "\n\n")

        run_env = os.environ.copy()
        run_env.update(env or {})
        # чтобы кириллица из дочернего процесса всегда читалась как UTF-8,
        # независимо от системной кодовой страницы Windows
        run_env["PYTHONIOENCODING"] = "utf-8"
        run_env["PYTHONUTF8"] = "1"

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

        def worker():
            try:
                proc = subprocess.Popen(
                    cmd, env=run_env, cwd=SCRIPT_DIR,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace",
                    bufsize=1, creationflags=creationflags,
                )
            except Exception as exc:
                self._queue.put(("line", f"Не удалось запустить процесс: {exc}\n"))
                self._queue.put(("done", -1))
                return
            self._proc = proc
            for line in proc.stdout:
                self._queue.put(("line", line))
            proc.wait()
            self._queue.put(("done", proc.returncode))

        threading.Thread(target=worker, daemon=True).start()


def _quote_for_display(arg: str) -> str:
    return f'"{arg}"' if " " in arg else arg


# ======================================================================
#  Общие виджеты форм
# ======================================================================

class FormBuilder:
    """Небольшой помощник для однотипных строк формы на ttk.Frame с grid."""

    def __init__(self, parent: tk.Widget):
        self.parent = parent
        self.row = 0
        parent.columnconfigure(1, weight=1)

    def _next_row(self) -> int:
        r = self.row
        self.row += 1
        return r

    def entry(self, label: str, default: str = "", width: int = 40, show: Optional[str] = None) -> tk.StringVar:
        var = tk.StringVar(value=default)
        r = self._next_row()
        ttk.Label(self.parent, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(self.parent, textvariable=var, width=width, show=show or "").grid(
            row=r, column=1, sticky="we", pady=3, columnspan=2,
        )
        return var

    def password_entry(self, label: str, width: int = 30) -> tk.StringVar:
        """Поле пароля с чекбоксом 'показать' — сразу в одной строке,
        без необходимости потом искать виджет по StringVar."""
        var = tk.StringVar()
        r = self._next_row()
        ttk.Label(self.parent, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        entry = ttk.Entry(self.parent, textvariable=var, width=width, show="*")
        entry.grid(row=r, column=1, sticky="w", pady=3)
        show_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.parent, text="показать", variable=show_var,
            command=lambda: entry.configure(show="" if show_var.get() else "*"),
        ).grid(row=r, column=2, sticky="w", padx=(6, 0), pady=3)
        return var

    def entry_with_button(self, label: str, button_text: str, on_click, default: str = "") -> tk.StringVar:
        var = tk.StringVar(value=default)
        r = self._next_row()
        ttk.Label(self.parent, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(self.parent, textvariable=var).grid(row=r, column=1, sticky="we", pady=3)
        ttk.Button(self.parent, text=button_text, command=lambda: on_click(var)).grid(
            row=r, column=2, sticky="e", padx=(6, 0), pady=3,
        )
        return var

    def file_picker(self, label: str, filetypes) -> tk.StringVar:
        def pick(var):
            path = filedialog.askopenfilename(title=label, filetypes=filetypes)
            if path:
                var.set(path)
        return self.entry_with_button(label, "Обзор…", pick)

    def folder_picker(self, label: str) -> tk.StringVar:
        def pick(var):
            path = filedialog.askdirectory(title=label)
            if path:
                var.set(path)
        return self.entry_with_button(label, "Обзор…", pick)

    def checkbox(self, label: str, default: bool = False) -> tk.BooleanVar:
        var = tk.BooleanVar(value=default)
        r = self._next_row()
        ttk.Checkbutton(self.parent, text=label, variable=var).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=2,
        )
        return var

    def combobox(self, label: str, values: list, default: str) -> tk.StringVar:
        var = tk.StringVar(value=default)
        r = self._next_row()
        ttk.Label(self.parent, text=label).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        cb = ttk.Combobox(self.parent, textvariable=var, values=values, state="readonly", width=15)
        cb.grid(row=r, column=1, sticky="w", pady=3)
        return var

    def two_entries(self, label1: str, default1: str, label2: str, default2: str, width: int = 18):
        var1, var2 = tk.StringVar(value=default1), tk.StringVar(value=default2)
        r = self._next_row()
        ttk.Label(self.parent, text=label1).grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        sub = ttk.Frame(self.parent)
        sub.grid(row=r, column=1, columnspan=2, sticky="w", pady=3)
        ttk.Entry(sub, textvariable=var1, width=width).pack(side="left")
        ttk.Label(sub, text=f"   {label2}").pack(side="left", padx=(10, 8))
        ttk.Entry(sub, textvariable=var2, width=width).pack(side="left")
        return var1, var2

    def separator(self):
        r = self._next_row()
        ttk.Separator(self.parent, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="we", pady=8,
        )


# ======================================================================
#  Базовое приложение
# ======================================================================

class BaseImporterGUI:
    """Общий каркас: поля Excel/страницы/фото + dry-run/allow-partial/
    verbose + кнопки запуска/сохранения лога + панель лога + статус-бар.
    Наследник добавляет свои поля через add_specific_fields() и собирает
    итоговые аргументы через build_extra_args()/build_env()."""

    WINDOW_TITLE = "Импорт"
    SCRIPT_NAME = ""  # переопределяется наследником, напр. "rusguard.py"

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(self.WINDOW_TITLE)
        root.geometry("880x720")
        root.minsize(700, 560)

        outer = ttk.Frame(root, padding=12)
        outer.pack(fill="both", expand=True)

        form_frame = ttk.Frame(outer)
        form_frame.pack(fill="x")
        self.form = FormBuilder(form_frame)

        self.excel_file = self.form.file_picker(
            "Excel-файл:", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")],
        )

        self.add_specific_fields(self.form)

        self.sheet_positions = self.form.entry("Лист-справочник должностей:", default="Должности", width=30)
        self.photos_root = self.form.folder_picker("Папка с фото (--photos-root):")
        self.only_page, self.only_person = self.form.two_entries(
            "Только страница (--only-page):", "", "Только № человека:", "", width=16,
        )

        self.form.separator()
        self.dry_run = self.form.checkbox("Dry-run — проверить, но ничего не сохранять", default=True)
        self.allow_partial = self.form.checkbox(
            "Allow-partial — сохранить валидные строки, даже если другие с проблемами",
        )
        self.verbose = self.form.checkbox("Подробный вывод (verbose)")

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill="x", pady=(10, 6))
        self.run_button = ttk.Button(btn_row, text="Запустить", command=self._on_run)
        self.run_button.pack(side="left")
        self.stop_button = ttk.Button(btn_row, text="Остановить", command=self._on_stop, state="disabled")
        self.stop_button.pack(side="left", padx=(6, 0))
        ttk.Button(btn_row, text="Сохранить лог…", command=self._on_save_log).pack(side="left", padx=(6, 0))

        self.status_var = tk.StringVar(value="Готово к запуску")
        self.status_label = ttk.Label(outer, textvariable=self.status_var, foreground="#374151")
        self.status_label.pack(fill="x", pady=(0, 6))

        self.log_panel = LogPanel(outer, root)
        self.log_panel.frame.pack(fill="both", expand=True)

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---- переопределяется наследником -------------------------------
    def add_specific_fields(self, form: FormBuilder):
        """Добавить свои поля МЕЖДУ Excel-файлом и общими полями (страницы,
        фото, only-page/only-person)."""
        pass

    def build_extra_args(self) -> list:
        """Аргументы командной строки, специфичные для конкретного скрипта."""
        raise NotImplementedError

    def build_env(self) -> dict:
        """Дополнительные переменные окружения для дочернего процесса
        (например, PGPASSWORD)."""
        return {}

    def validate(self) -> Optional[str]:
        """Вернуть текст ошибки, если поля невалидны, иначе None.
        Наследник может расширить (вызвать super().validate() первым)."""
        if not self.excel_file.get().strip():
            return "Укажите Excel-файл."
        if not os.path.exists(self.excel_file.get().strip()):
            return f"Excel-файл не найден: {self.excel_file.get().strip()}"
        only_page = self.only_page.get().strip()
        only_person = self.only_person.get().strip()
        if only_person and not only_page:
            return "«Только № человека» требует также указать «Только страница»."
        if only_person and not only_person.lstrip("-").isdigit():
            return "«Только № человека» должно быть числом."
        return None

    # ------------------------------------------------------------------
    def _common_args(self) -> list:
        args = [self.excel_file.get().strip()]
        args += self.build_extra_args()
        sp = self.sheet_positions.get().strip()
        if sp:
            args += ["--sheet-positions", sp]
        pr = self.photos_root.get().strip()
        if pr:
            args += ["--photos-root", pr]
        op = self.only_page.get().strip()
        if op:
            args += ["--only-page", op]
        opn = self.only_person.get().strip()
        if opn:
            args += ["--only-person", opn]
        if self.dry_run.get():
            args.append("--dry-run")
        if self.allow_partial.get():
            args.append("--allow-partial")
        if self.verbose.get():
            args.append("--verbose")
        return args

    def _on_run(self):
        error = self.validate()
        if error:
            messagebox.showerror("Проверьте поля", error)
            return
        script_path = os.path.join(SCRIPT_DIR, self.SCRIPT_NAME)
        if not os.path.exists(script_path):
            messagebox.showerror(
                "Скрипт не найден",
                f"Не найден {self.SCRIPT_NAME} рядом с этой программой:\n{script_path}\n\n"
                f"Положите {self.SCRIPT_NAME} в ту же папку.",
            )
            return
        cmd = [sys.executable, script_path] + self._common_args()
        env = self.build_env()

        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.status_var.set("Выполняется…")
        self.status_label.configure(foreground="#374151")

        self.log_panel.run(cmd, env, on_done=self._on_finished)

    def _on_stop(self):
        self.log_panel.stop()

    def _on_finished(self, code: int):
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        text, color = default_status_for_code(code)
        self.status_var.set(text)
        self.status_label.configure(foreground=color)

    def _on_save_log(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить лог", defaultextension=".txt",
            filetypes=[("Текстовый файл", "*.txt"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.log_panel.get_text())
        except OSError as exc:
            messagebox.showerror("Не удалось сохранить", str(exc))

    def _on_close(self):
        if self.log_panel.is_running():
            if not messagebox.askyesno(
                "Процесс ещё выполняется",
                "Импорт ещё не завершён. Закрыть окно и остановить его?",
            ):
                return
            self.log_panel.stop()
        self.root.destroy()


def run_app(app_cls):
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if os.name == "nt" else "clam")
    except tk.TclError:
        pass
    app_cls(root)
    root.mainloop()
