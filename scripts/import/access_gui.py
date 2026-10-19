#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
access_gui.py — графическая обёртка над access.py.

Просто собирает аргументы командной строки и запускает access.py
(должен лежать в той же папке) как отдельный процесс, показывая его вывод
в окне построчно. Логика не дублируется и не меняется — режим "всё или
ничего", формат [Фото] и т.п. работают точно так же, как при запуске
access.py напрямую из консоли.

Требует: gui_common.py и access.py в той же папке.
Запуск: python access_gui.py
"""

from __future__ import annotations

import os

try:
    import gui_common as gc
except ImportError:
    raise SystemExit(
        "Не найден gui_common.py. Положите его в ту же папку, что и access_gui.py."
    )


class AccessGUI(gc.BaseImporterGUI):
    WINDOW_TITLE = "Карточки — добавление в Access"
    SCRIPT_NAME = "access.py"

    def add_specific_fields(self, form: "gc.FormBuilder"):
        self.access_db = form.file_picker(
            "Файл Access (.accdb):",
            filetypes=[("Access", "*.accdb *.mdb"), ("Все файлы", "*.*")],
        )
        self.access_table = form.entry("Имя таблицы:", default="Import", width=20)
        self.max_dimension = form.entry(
            "Макс. размер фото, px (--access-max-photo-dimension):", default="800", width=10,
        )
        self.position_case = form.combobox(
            "Регистр должности:", values=["lower", "asis"], default="lower",
        )

    def validate(self):
        error = super().validate()
        if error:
            return error
        access_db = self.access_db.get().strip()
        if not access_db:
            return "Укажите файл Access (.accdb)."
        if not os.path.exists(access_db):
            return f"Файл Access не найден: {access_db}"
        md = self.max_dimension.get().strip()
        if md and not md.isdigit():
            return "«Макс. размер фото» должно быть числом (или пусто)."
        return None

    def build_extra_args(self) -> list:
        args = [self.access_db.get().strip()]
        table = self.access_table.get().strip()
        if table:
            args += ["--access-table", table]
        md = self.max_dimension.get().strip()
        if md:
            args += ["--access-max-photo-dimension", md]
        args += ["--access-position-case", self.position_case.get()]
        return args


if __name__ == "__main__":
    gc.run_app(AccessGUI)
