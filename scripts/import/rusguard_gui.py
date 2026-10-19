#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rusguard_gui.py — графическая обёртка над rusguard.py.

Просто собирает аргументы командной строки и запускает rusguard.py
(должен лежать в той же папке) как отдельный процесс, показывая его вывод
в окне построчно. Логика импорта не дублируется и не меняется — вся
валидация, режим "всё или ничего" и т.п. работают точно так же, как при
запуске rusguard.py напрямую из консоли.

Пароль передаётся дочернему процессу через переменную окружения
PGPASSWORD, а не аргументом командной строки — не виден в списке
процессов (Диспетчер задач).

Требует: gui_common.py и rusguard.py в той же папке.
Запуск: python rusguard_gui.py
"""

from __future__ import annotations

try:
    import gui_common as gc
except ImportError:
    raise SystemExit(
        "Не найден gui_common.py. Положите его в ту же папку, что и rusguard_gui.py."
    )


class RusguardGUI(gc.BaseImporterGUI):
    WINDOW_TITLE = "RusGuard — импорт в PostgreSQL"
    SCRIPT_NAME = "rusguard.py"

    def add_specific_fields(self, form: "gc.FormBuilder"):
        self.host, self.port = form.two_entries(
            "Хост БД:", "localhost", "Порт:", "5432", width=18,
        )
        self.dbname = form.entry("Имя базы данных:", width=30)
        self.user = form.entry("Пользователь БД:", width=30)
        self.password = form.password_entry("Пароль:", width=30)

    def validate(self):
        error = super().validate()
        if error:
            return error
        if not self.dbname.get().strip():
            return "Укажите имя базы данных."
        if not self.user.get().strip():
            return "Укажите пользователя БД."
        if not self.password.get():
            return "Укажите пароль (иначе дочерний процесс не сможет его запросить в этом окне)."
        try:
            int(self.port.get().strip())
        except ValueError:
            return "Порт должен быть числом."
        return None

    def build_extra_args(self) -> list:
        return [
            "--host", self.host.get().strip() or "localhost",
            "--port", self.port.get().strip() or "5432",
            "--dbname", self.dbname.get().strip(),
            "--user", self.user.get().strip(),
        ]

    def build_env(self) -> dict:
        # пароль уходит через переменную окружения, а не аргумент командной
        # строки — не будет виден в списке процессов
        return {"PGPASSWORD": self.password.get()}


if __name__ == "__main__":
    gc.run_app(RusguardGUI)
