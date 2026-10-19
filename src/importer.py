# -*- coding: utf-8 -*-
"""Оркестрация импорта строк таблицы в RusGuard (PostgreSQL)."""

from __future__ import annotations

import db
import sheets

OK, FAILED_TO_START, PROBLEMS = 0, 1, 2

_KEY_RU = {
    "created": "создан и назначен",
    "reused": "карта уже была в acskeys, назначена сотруднику",
    "already": "уже был назначен этому сотруднику",
}


def run_import(cfg: dict, mode: str, entries: list, key_format: str,
               dry_run: bool, allow_partial: bool, log) -> int:
    """entries — строки таблицы (sheets.Entry); log — функция str -> None. Возвращает код результата."""
    try:
        conn = db.connect(cfg)
    except RuntimeError as exc:
        log(f"ОШИБКА: {exc}")
        return FAILED_TO_START

    created_ids, problems = [], 0
    try:
        with conn.cursor() as cur:
            positions, groups = db.fetch_positions(cur), db.fetch_groups(cur)
            log(f"С сервера получено: должностей {len(positions)}, групп {len(groups)}")
            rows = sheets.build_rows(entries, mode, key_format, positions, groups)
            if not rows:
                log("ОШИБКА: таблица пуста.")
                return FAILED_TO_START

            for row in rows:
                for w in row.warnings:
                    log(f"ВНИМАНИЕ {row.label}: {w}")
                for e in row.errors:
                    log(f"ОШИБКА {row.label}: {e}")
            problems += sum(1 for r in rows if r.errors)
            valid = [r for r in rows if not r.errors]

            seen_keys: dict = {}
            for row in valid:
                for k in row.keys:
                    if k in seen_keys and seen_keys[k] != row.fullname:
                        log(f"ВНИМАНИЕ: ключ {k} указан у нескольких людей: {seen_keys[k]}, {row.fullname}")
                    seen_keys.setdefault(k, row.fullname)

            log("-" * 60)
            for row in valid:
                cur.execute("SAVEPOINT row_sp")
                try:
                    created, conflict = _process_row(cur, row, mode, log)
                    problems += conflict
                    if created is not None:
                        created_ids.append(created)
                    cur.execute("RELEASE SAVEPOINT row_sp")
                except Exception as exc:
                    cur.execute("ROLLBACK TO SAVEPOINT row_sp")
                    cur.execute("RELEASE SAVEPOINT row_sp")
                    problems += 1
                    log(f"ОШИБКА {row.label}: {exc}")

            abort = bool(problems) and not allow_partial and not dry_run

        if dry_run:
            conn.rollback()
            log("DRY RUN: в БД ничего не сохранено (ROLLBACK)." + (" Есть проблемы — см. выше." if problems else ""))
        elif abort:
            conn.rollback()
            log("=" * 60)
            log("ИМПОРТ ОТМЕНЁН: найдены проблемы (см. выше), в БД не сохранено ничего. "
                "Исправьте строки или включите «Сохранять валидные строки».")
        else:
            conn.commit()
            log("=" * 60)
            log(f"Сохранено в БД. Создано профилей: {len(created_ids)}. Проблем: {problems}.")
        return PROBLEMS if problems else OK
    finally:
        conn.close()


def _process_row(cur, row, mode: str, log):
    """-> (id созданного сотрудника или None, число конфликтов ключей)."""
    created_id = None
    if mode == sheets.MODE_ADD:
        if row.employee_id:
            employee_id = row.employee_id
        else:
            matches = db.find_employees(cur, row.lastname, row.firstname, row.secondname)
            if not matches:
                raise RuntimeError("сотрудник с таким ФИО не найден в БД")
            if len(matches) > 1:
                raise RuntimeError("в БД несколько сотрудников с таким ФИО — выберите нужного через «Найти в БД»")
            employee_id = matches[0]
    else:
        if db.find_employees(cur, row.lastname, row.firstname, row.secondname):
            raise RuntimeError("сотрудник с таким ФИО уже есть в БД — для добавления ключа используйте другой режим")
        employee_id = created_id = db.create_employee(cur, row)
        log(f"{row.label}: профиль создан")
    if mode == sheets.MODE_ADD and (row.position_id or row.group_id):
        db.update_employee(cur, employee_id, row.position_id, row.group_id)
        log(f"{row.label}: должность/группа обновлены")
    if row.photo:
        db.upsert_photo(cur, employee_id, row.photo)
        log(f"{row.label}: фото сохранено")

    for key in row.remove_keys:
        if db.unassign_key(cur, employee_id, key):
            log(f"{row.label}: ключ {key} снят")
        else:
            log(f"ВНИМАНИЕ {row.label}: ключ {key} не был назначен — пропущено")

    conflicts = 0
    for key in row.keys:
        result = db.ensure_key_assigned(cur, employee_id, key, row.fullname)
        if result == "conflict":
            conflicts += 1
            log(f"ВНИМАНИЕ {row.label}: ключ {key} уже назначен ДРУГОМУ сотруднику — пропущен")
        else:
            log(f"{row.label}: ключ {key}: {_KEY_RU[result]}")
    return created_id, conflicts
