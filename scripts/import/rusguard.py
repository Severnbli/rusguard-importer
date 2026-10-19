#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rusguard.py — импорт сотрудников/учащихся, ключей СКУД (acskeys) и
карточных фотографий в БД RusGuard из Excel-файла.

СТРУКТУРА ВХОДНОГО EXCEL-ФАЙЛА
  Один лист — справочник должностей (по умолчанию называется "Должности",
  см. --sheet-positions): столбцы №, Наименование, Айди. Айди может быть
  пустым — тогда для этой должности employeepositionid будет NULL (это не
  ошибка, см. например "Учащийся").

  ВСЕ ОСТАЛЬНЫЕ листы — «страницы» с людьми (сотрудники, конкретный класс
  и т.п.), каждая обрабатывается независимо. Столбцы (ищутся по заголовку,
  регистр не важен):
      №                    — порядковый номер человека на странице.
                             ОБЯЗАТЕЛЕН, должен быть уникален на странице
                             (используется для сопоставления с файлом фото).
      ФИО                  — "Фамилия Имя Отчество". ОБЯЗАТЕЛЬНО.
      Должность            — сверяется со справочником должностей. Можно
                             оставить пустым. Если указано — должно
                             найтись в справочнике (иначе это ошибка данных).
      Номер карты          — номер(а) карты, можно несколько через , ; / или
                             перенос строки. Необязательно.
      Карточка в системе   — "Да"/"Нет". Импортируются в БД ТОЛЬКО строки
                             со значением "Нет". ОБЯЗАТЕЛЬНЫЙ столбец.
      Айди группы          — UUID группы. Одно значение НА ВСЮ СТРАНИЦУ
                             (по факту читается из любой заполненной ячейки
                             этого столбца на странице; если заполнено
                             по-разному в разных строках — ошибка).
                             ОБЯЗАТЕЛЬНО (employeegroupid — NOT NULL).
      Путь к фото          — папка с фото, тоже ОДНА НА ВСЮ СТРАНИЦУ.
                             Необязательно (нет столбца/значения — фото не
                             импортируются для этой страницы). Абсолютный
                             путь Windows (с буквой диска, напр. "C:\\...")
                             используется как есть; любой другой (в т.ч.
                             "/photos/employee") считается ОТНОСИТЕЛЬНО
                             --photos-root. В папке фото должны называться
                             "<№>.<любое расширение>" (jpg/jpeg/png/bmp/
                             gif/tif/tiff/webp) — сравнение по числу, не по
                             строке, ведущие нули не важны.

ПОЛИТИКА "ВСЁ ИЛИ НИЧЕГО" (по умолчанию, как и раньше)
  Если хотя бы одна ПОДЛЕЖАЩАЯ импорту строка (см. фильтр по "Карточка в
  системе" выше) не проходит проверку — в БД не сохраняется вообще ничего.
  Флаг --allow-partial возвращает старое поведение (сохранить валидные,
  проблемные пропустить).

ЧТО СКРИПТ ДЕЛАЕТ И ЧЕГО НЕ ДЕЛАЕТ (без изменений от прошлой версии)
  - Создаёт/дополняет employee, идемпотентен (поиск по ФИО).
  - Создаёт карты в acskeys БЕЗ биометрических полей, привязывает через
    acskey2employeeassignment. Конфликт активного назначения карты на
    ДРУГОГО сотрудника — не трогает, только сообщает (решать вручную).
  - Фото — в employeephoto, employeeimagetype = 0. Биометрию не трогает.

ВЫБОРОЧНЫЙ ЗАПУСК
  --only-page "2А"        — обработать только эту страницу.
  --only-person 5         — обработать только человека с №5 (требует
                             --only-page, т.к. номера не уникальны между
                             страницами).

ПРИМЕРЫ ЗАПУСКА
    python rusguard.py import.xlsm --host 192.168.1.10 \\
        --dbname rusguard --user rg_import --photos-root "C:\\Data" --dry-run

    python rusguard.py import.xlsm --host 192.168.1.10 \\
        --dbname rusguard --user rg_import --photos-root "C:\\Data" \\
        --only-page "2А" --only-person 5 --dry-run

Пароль (в порядке приоритета): флаг --password/--pass, переменная
окружения PGPASSWORD, интерактивный скрытый ввод.

Для добавления ФИО+Должность+Фото в Access (карточки.accdb) используйте
отдельный скрипт access_import.py — он не связан с этим и запускается
отдельно.
"""

from __future__ import annotations

import argparse
import getpass
import logging
import os
import re
import sys
import uuid
from dataclasses import dataclass, field
from typing import Optional

try:
    import psycopg2
except ImportError:
    sys.exit("Не найден пакет psycopg2. Установите: pip install psycopg2-binary")

try:
    import common
except ImportError:
    sys.exit("Не найден common.py. Положите common.py в ту же папку, что и rusguard.py.")

log = logging.getLogger("rusguard")

# employeephoto.employeeimagetype: 0 = обычное фото для карточки/пропуска.
EMPLOYEE_IMAGE_TYPE_CARD = 0
PHOTO_NUMBER_DEFAULT = 1

_KEY_SPLIT_RE = re.compile(r"[,;/\n]+")


# ======================================================================
#  Структуры данных
# ======================================================================

@dataclass
class EmployeeRow(common.BasePersonRow):
    lastname: str = ""
    firstname: str = ""
    secondname: Optional[str] = None
    position_name: Optional[str] = None
    position_id: Optional[str] = None
    group_id: Optional[str] = None
    key_numbers: list = field(default_factory=list)


@dataclass
class Stats:
    employees_created: int = 0
    employees_updated: int = 0
    employees_unchanged: int = 0
    keys_created: int = 0
    keys_reused: int = 0
    keys_already: int = 0
    keys_conflict: int = 0
    photos_inserted: int = 0
    photos_updated: int = 0
    row_errors: list = field(default_factory=list)


# ======================================================================
#  Разбор Excel
# ======================================================================

def split_fio(fullname: str):
    """'Фамилия Имя Отчество' -> (lastname, firstname, secondname, warnings)."""
    warnings = []
    cleaned = re.sub(r"\s+", " ", fullname.strip())
    tokens = cleaned.split(" ") if cleaned else []
    if len(tokens) == 3:
        lastname, firstname, secondname = tokens
    elif len(tokens) == 2:
        lastname, firstname = tokens
        secondname = None
        warnings.append("отчество не указано (в ФИО только 2 слова)")
    elif len(tokens) >= 4:
        lastname, firstname = tokens[0], tokens[1]
        secondname = " ".join(tokens[2:])
        warnings.append(
            f"в ФИО {len(tokens)} слов — разобрано как фамилия='{lastname}', "
            f"имя='{firstname}', отчество='{secondname}'; проверьте вручную"
        )
    elif len(tokens) == 1:
        lastname, firstname, secondname = tokens[0], "", None
        warnings.append("в ФИО только одно слово — записано как фамилия; проверьте вручную")
    else:
        lastname, firstname, secondname = "", "", None
        warnings.append("пустое ФИО")
    return lastname, firstname, secondname, warnings


def load_positions_dict(ws, sheet_label: str) -> dict:
    """Лист №/Наименование/Айди -> {наименование: uuid_str_или_None}.
    Пустой Айди — валидное значение (employeepositionid будет NULL)."""
    header = next(ws.iter_rows(min_row=1, max_row=1))
    name_col = common.find_col(header, "наименование", "name", "название")
    id_col = common.find_col(header, "айди", "id", "идентификатор")
    if name_col is None or id_col is None:
        raise ValueError(f"На листе '{sheet_label}' не найдены столбцы Наименование/Айди.")
    result: dict[str, Optional[str]] = {}
    dupes = []
    for row in ws.iter_rows(min_row=2):
        name_cell = row[name_col - 1]
        name = str(name_cell.value).strip() if name_cell.value not in (None, "") else None
        if not name:
            continue
        id_cell = row[id_col - 1]
        raw_id = str(id_cell.value).strip() if id_cell.value not in (None, "") else None
        if raw_id is None:
            norm_id = None
        else:
            try:
                norm_id = str(uuid.UUID(raw_id))
            except ValueError:
                raise ValueError(
                    f"Лист '{sheet_label}', строка {id_cell.row}: '{raw_id}' не похож на UUID."
                )
        if name in result:
            dupes.append(name)
        result[name] = norm_id
    if dupes:
        log.warning(
            "Лист '%s': повторяющиеся названия должностей (учтён последний вариант): %s",
            sheet_label, ", ".join(sorted(set(dupes))),
        )
    return result


def parse_key_cell(value, warnings_out: list) -> list:
    """Содержимое ячейки 'Номер карты' -> список номеров карт (int).
    Поддерживает несколько номеров через , ; / или перенос строки."""
    if value is None:
        return []
    if isinstance(value, bool):
        return []
    if isinstance(value, int):
        return [value] if value > 0 else []
    if isinstance(value, float):
        if value.is_integer() and value > 0:
            return [int(value)]
        warnings_out.append(f"нецелое значение в столбце 'Номер карты': {value!r} — пропущено")
        return []
    text = str(value).strip()
    if not text:
        return []
    parts = [p.strip() for p in _KEY_SPLIT_RE.split(text) if p.strip()]
    numbers = []
    bad_parts = []
    for p in parts:
        digits = p.replace(" ", "")
        if digits.isdigit():
            numbers.append(int(digits))
        else:
            bad_parts.append(p)
    if bad_parts:
        warnings_out.append(
            f"в столбце 'Номер карты' не распознано: {', '.join(repr(b) for b in bad_parts)} "
            f"(исходное значение: '{text}') — пропущено"
        )
    seen = set()
    unique_numbers = []
    for n in numbers:
        if n not in seen:
            seen.add(n)
            unique_numbers.append(n)
    return unique_numbers


def load_page_rows(ws, page_name: str, positions: dict, photos_root: Optional[str]):
    """-> (rows: list[EmployeeRow], page_errors: list[str])."""
    header = next(ws.iter_rows(min_row=1, max_row=1))
    col_num = common.find_col(header, "№", "номер", "n")
    col_fio = common.find_col(header, "фио", "fullname")
    col_card_number = common.find_col(header, "номер карты", "key", "карта")
    col_in_system = common.find_col(header, "карточка в системе")
    col_position = common.find_col(header, "должность", "position")
    col_photo_path = common.find_col(header, "путь к фото")
    col_group_id = common.find_col(header, "айди группы")

    required = [
        ("№", col_num), ("ФИО", col_fio), ("Карточка в системе", col_in_system),
        ("Должность", col_position), ("Айди группы", col_group_id),
    ]
    missing = [n for n, c in required if c is None]
    if missing:
        return [], [f"на странице '{page_name}' не найдены обязательные столбцы: {', '.join(missing)}"]

    group_id_raw, errs = common.collect_page_wide_value(ws, col_group_id, "Айди группы", page_name)
    if errs:
        return [], errs
    if group_id_raw is None:
        return [], [f"на странице '{page_name}' не заполнен ни один 'Айди группы'"]
    try:
        group_id = str(uuid.UUID(group_id_raw))
    except ValueError:
        return [], [f"на странице '{page_name}': Айди группы '{group_id_raw}' не похож на UUID"]

    photo_index: dict = {}
    if col_photo_path:
        path_raw, errs = common.collect_page_wide_value(ws, col_photo_path, "Путь к фото", page_name)
        if errs:
            return [], errs
        if path_raw:
            try:
                folder = common.resolve_photo_folder(path_raw, photos_root)
            except ValueError as exc:
                return [], [f"на странице '{page_name}': {exc}"]
            photo_index = common.index_photo_folder(folder)

    rows: list[EmployeeRow] = []
    seen_numbers: dict[int, int] = {}
    skipped_hash_count = 0
    for row in ws.iter_rows(min_row=2):
        if all(c.value in (None, "") for c in row):
            continue

        fio_cell = row[col_fio - 1]
        excel_row = fio_cell.row
        fullname_raw = str(fio_cell.value).strip() if fio_cell.value not in (None, "") else ""

        num_val = row[col_num - 1].value
        try:
            person_number = int(num_val) if num_val not in (None, "") else None
        except (ValueError, TypeError):
            person_number = None

        if fullname_raw.startswith("#"):
            # строка сознательно закомментирована — саму запись пропускаем
            # (не создаём EmployeeRow, не трогаем фото), но № за ней остаётся
            # "занят": не считается осиротевшим в папке с фото и не может
            # быть случайно переиспользован другой строкой без предупреждения.
            skipped_hash_count += 1
            if person_number is not None:
                seen_numbers[person_number] = excel_row
            continue

        emp = EmployeeRow(
            page=page_name, excel_row=excel_row, person_number=person_number,
            fullname_raw=fullname_raw, group_id=group_id,
        )

        if person_number is None:
            emp.fatal_errors.append("не заполнен или нечисловой №")
        elif person_number in seen_numbers:
            emp.fatal_errors.append(
                f"№ {person_number} повторяется на этой странице (первый раз — "
                f"строка {seen_numbers[person_number]})"
            )
        else:
            seen_numbers[person_number] = excel_row

        if not fullname_raw:
            emp.fatal_errors.append("не заполнен ФИО")
            rows.append(emp)
            continue

        lastname, firstname, secondname, warns = split_fio(fullname_raw)
        emp.lastname, emp.firstname, emp.secondname = lastname, firstname, secondname
        emp.warnings.extend(warns)

        pos_cell = row[col_position - 1]
        emp.position_name = str(pos_cell.value).strip() if pos_cell.value not in (None, "") else None
        if emp.position_name:
            if emp.position_name not in positions:
                emp.fatal_errors.append(f"должность '{emp.position_name}' не найдена в справочнике")
            else:
                emp.position_id = positions[emp.position_name]  # может быть None — это ок

        if col_card_number:
            emp.key_numbers = parse_key_cell(row[col_card_number - 1].value, emp.warnings)

        emp.eligible = common.is_flag_no(row[col_in_system - 1].value, emp.warnings, "Карточка в системе")

        if person_number is not None and photo_index:
            photo_path = photo_index.get(person_number)
            if photo_path:
                try:
                    data = photo_path.read_bytes()
                except OSError as exc:
                    emp.warnings.append(f"не удалось прочитать файл фото {photo_path}: {exc}")
                    data = None
                if data:
                    problem = common.validate_photo_bytes(data)
                    if problem:
                        emp.warnings.append(
                            f"файл фото {photo_path.name} повреждён/не распознан ({problem}) — фото пропущено"
                        )
                    else:
                        emp.photo_bytes = data

        rows.append(emp)

    if skipped_hash_count:
        log.info(
            "Страница '%s': пропущено %d строк(и), ФИО которых начинается с '#'",
            page_name, skipped_hash_count,
        )

    if photo_index:
        extra = sorted(set(photo_index) - set(seen_numbers))
        if extra:
            shown = ", ".join(str(n) for n in extra[:20]) + (" …" if len(extra) > 20 else "")
            log.warning(
                "Страница '%s': в папке с фото есть файлы для номеров без строк на странице: %s",
                page_name, shown,
            )

    return rows, []


def load_all_rows(path: str, sheet_positions: str, photos_root: Optional[str], only_page: Optional[str]):
    log.info("Читаю файл: %s", path)
    wb, page_names = common.open_workbook_and_get_pages(path, sheet_positions, only_page)

    positions = load_positions_dict(wb[sheet_positions], sheet_positions)
    log.info("Справочник должностей: %d записей", len(positions))

    all_rows: list[EmployeeRow] = []
    for page_name in page_names:
        rows, page_errors = load_page_rows(wb[page_name], page_name, positions, photos_root)
        if page_errors:
            for e in page_errors:
                log.error("%s", e)
            all_rows.append(EmployeeRow(
                page=page_name, excel_row=1, person_number=None, fullname_raw="",
                fatal_errors=list(page_errors),
            ))
            continue
        log.info("Страница '%s': %d строк", page_name, len(rows))
        all_rows.extend(rows)

    return all_rows, positions


def find_duplicate_keys_in_rows(rows) -> dict:
    seen: dict[int, list] = {}
    for emp in rows:
        if not emp.is_valid:
            continue
        for k in emp.key_numbers:
            seen.setdefault(k, []).append(emp.display_name)
    return {k: names for k, names in seen.items() if len(names) > 1}


# ======================================================================
#  Слой БД
# ======================================================================

def get_connection(args):
    password = args.password or os.environ.get("PGPASSWORD")
    if password is None:
        try:
            password = getpass.getpass(
                f"Пароль для {args.user}@{args.host}:{args.port}/{args.dbname}: "
            )
        except (EOFError, KeyboardInterrupt):
            log.error(
                "Пароль не введён. Укажите --password/--pass или переменную PGPASSWORD."
            )
            sys.exit(1)
    try:
        conn = psycopg2.connect(
            host=args.host, port=args.port, dbname=args.dbname,
            user=args.user, password=password, connect_timeout=10,
        )
    except psycopg2.OperationalError as exc:
        log.error("Не удалось подключиться к БД: %s", exc)
        sys.exit(1)
    conn.autocommit = False
    return conn


def preflight_check_ids(cur, position_ids: set, group_ids: set):
    """Возвращает (missing_group_ids, missing_position_ids) — Id, которые
    реально используются в подлежащих импорту строках, но не существуют
    в целевой БД."""
    missing_group_ids = set()
    if group_ids:
        ids = list(group_ids)
        cur.execute("SELECT _id::text FROM employeegroup WHERE _id = ANY(%s::uuid[])", (ids,))
        found = {r[0] for r in cur.fetchall()}
        missing_group_ids = set(ids) - found

    missing_position_ids = set()
    if position_ids:
        ids = list(position_ids)
        cur.execute("SELECT _id::text FROM employeeposition WHERE _id = ANY(%s::uuid[])", (ids,))
        found = {r[0] for r in cur.fetchall()}
        missing_position_ids = set(ids) - found

    return missing_group_ids, missing_position_ids


def find_existing_employee(cur, lastname, firstname, secondname):
    cur.execute(
        """
        SELECT _id, employeegroupid, employeepositionid
        FROM employee
        WHERE isremoved = false
          AND lastname IS NOT DISTINCT FROM %s
          AND firstname IS NOT DISTINCT FROM %s
          AND secondname IS NOT DISTINCT FROM %s
        """,
        (lastname or None, firstname or None, secondname),
    )
    return cur.fetchall()


def upsert_employee(cur, emp: EmployeeRow):
    """-> (employee_id, action) где action: created|updated|unchanged|ambiguous."""
    matches = find_existing_employee(cur, emp.lastname, emp.firstname, emp.secondname)
    if len(matches) > 1:
        return None, "ambiguous"
    if len(matches) == 1:
        emp_id, cur_group, cur_pos = matches[0]
        cur_group = str(cur_group) if cur_group else None
        cur_pos = str(cur_pos) if cur_pos else None
        if cur_group != emp.group_id or cur_pos != emp.position_id:
            cur.execute(
                "UPDATE employee SET employeegroupid = %s, employeepositionid = %s WHERE _id = %s",
                (emp.group_id, emp.position_id, emp_id),
            )
            return emp_id, "updated"
        return emp_id, "unchanged"

    cur.execute(
        """
        INSERT INTO employee
            (lastname, firstname, secondname, employeegroupid, employeepositionid,
             isaccesslevelsinherited)
        VALUES (%s, %s, %s, %s, %s, true)
        RETURNING _id
        """,
        (emp.lastname or None, emp.firstname or None, emp.secondname, emp.group_id, emp.position_id),
    )
    return cur.fetchone()[0], "created"


def ensure_key_assigned(cur, employee_id, keynumber: int, key_label: str) -> str:
    """-> created|reused|already_assigned|conflict."""
    cur.execute("SELECT keynumber FROM acskeys WHERE keynumber = %s", (keynumber,))
    already_in_acskeys = cur.fetchone() is not None
    if not already_in_acskeys:
        cur.execute("INSERT INTO acskeys (keynumber, name) VALUES (%s, %s)", (keynumber, key_label))

    cur.execute(
        """
        SELECT employeeid FROM acskey2employeeassignment
        WHERE acskeyid = %s AND assignmentmodificationtype = 0
        """,
        (keynumber,),
    )
    row = cur.fetchone()
    if row is not None:
        return "already_assigned" if str(row[0]) == str(employee_id) else "conflict"

    cur.execute(
        """
        SELECT COALESCE(MAX(indexnumberofassignedkey), 0) + 1
        FROM acskey2employeeassignment
        WHERE employeeid = %s AND assignmentmodificationtype = 0
        """,
        (employee_id,),
    )
    next_index = cur.fetchone()[0]
    cur.execute(
        """
        INSERT INTO acskey2employeeassignment
            (employeeid, acskeyid, assignmentmodificationtype, indexnumberofassignedkey)
        VALUES (%s, %s, 0, %s)
        """,
        (employee_id, keynumber, next_index),
    )
    return "reused" if already_in_acskeys else "created"


def upsert_photo(cur, employee_id, photo_bytes: bytes,
                  photonumber: int = PHOTO_NUMBER_DEFAULT,
                  employeeimagetype: int = EMPLOYEE_IMAGE_TYPE_CARD) -> str:
    cur.execute(
        """
        SELECT _id FROM employeephoto
        WHERE employeeid = %s AND photonumber = %s
        ORDER BY modificationdatetime DESC LIMIT 1
        """,
        (employee_id, photonumber),
    )
    row = cur.fetchone()
    if row is not None:
        cur.execute(
            "UPDATE employeephoto SET photo = %s, employeeimagetype = %s WHERE _id = %s",
            (psycopg2.Binary(photo_bytes), employeeimagetype, row[0]),
        )
        return "updated"
    cur.execute(
        """
        INSERT INTO employeephoto (employeeid, photonumber, photo, employeeimagetype)
        VALUES (%s, %s, %s, %s)
        """,
        (employee_id, photonumber, psycopg2.Binary(photo_bytes), employeeimagetype),
    )
    return "inserted"


_ACTION_RU = {
    "created": "создан",
    "updated": "обновлён (группа/должность)",
    "unchanged": "уже существовал, без изменений",
}


def process_row(cur, emp: EmployeeRow, stats: Stats):
    employee_id, action = upsert_employee(cur, emp)
    if action == "ambiguous":
        raise RuntimeError(
            "в БД несколько активных сотрудников с таким же ФИО — "
            "автоматическое сопоставление невозможно, разберите вручную"
        )

    if action == "created":
        stats.employees_created += 1
    elif action == "updated":
        stats.employees_updated += 1
    else:
        stats.employees_unchanged += 1
    log.info("  %s — сотрудник: %s", emp.display_name, _ACTION_RU[action])

    for keynum in emp.key_numbers:
        key_action = ensure_key_assigned(cur, employee_id, keynum, emp.fullname_raw)
        if key_action == "created":
            stats.keys_created += 1
            log.info("      ключ %s: создан и назначен", keynum)
        elif key_action == "reused":
            stats.keys_reused += 1
            log.info("      ключ %s: карта уже была в acskeys, назначена сотруднику", keynum)
        elif key_action == "already_assigned":
            stats.keys_already += 1
            log.info("      ключ %s: уже был назначен этому сотруднику", keynum)
        else:
            stats.keys_conflict += 1
            log.warning(
                "      ключ %s: УЖЕ активно назначен ДРУГОМУ сотруднику — пропущено, "
                "разберите конфликт вручную в RusGuard",
                keynum,
            )

    if emp.photo_bytes:
        photo_action = upsert_photo(cur, employee_id, emp.photo_bytes)
        if photo_action == "inserted":
            stats.photos_inserted += 1
            log.info("      фото: сохранено")
        else:
            stats.photos_updated += 1
            log.info("      фото: обновлено")


# ======================================================================
#  CLI / main
# ======================================================================

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="rusguard.py",
        description="Импорт сотрудников/учащихся, ключей(карт), фото RusGuard из Excel в PostgreSQL.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    common.add_common_cli_args(p)
    p.add_argument("--host", default="localhost", help="адрес сервера БД (по умолчанию localhost)")
    p.add_argument("--port", type=int, default=5432, help="порт БД (по умолчанию 5432)")
    p.add_argument("--dbname", "--db", dest="dbname", required=True, help="имя базы данных")
    p.add_argument("--user", required=True, help="пользователь БД")
    p.add_argument(
        "--password", "--pass", dest="password", default=None,
        help="пароль (виден в истории shell/ps — лучше PGPASSWORD или интерактивный ввод)",
    )
    return p


def print_summary(stats: Stats, skipped_rows: int, already_in_system: int, dry_run: bool, aborted: bool):
    log.info("=" * 66)
    if dry_run:
        log.info("ИТОГ (DRY RUN — в БД НИЧЕГО не сохранено)")
    elif aborted:
        log.info("ИТОГ (ИМПОРТ ОТМЕНЁН — в БД НИЧЕГО не сохранено)")
    else:
        log.info("ИТОГ (сохранено в БД)")
    log.info(
        "Сотрудники: создано %d, обновлено %d, без изменений %d, пропущено (ошибки) %d, "
        "уже 'в системе' (не трогали) %d",
        stats.employees_created, stats.employees_updated, stats.employees_unchanged,
        skipped_rows, already_in_system,
    )
    log.info(
        "Ключи: новых %d, переиспользовано карт %d, уже назначены %d, КОНФЛИКТОВ %d",
        stats.keys_created, stats.keys_reused, stats.keys_already, stats.keys_conflict,
    )
    log.info("Фото: сохранено новых %d, обновлено %d", stats.photos_inserted, stats.photos_updated)
    if stats.row_errors:
        log.warning("Ошибки при обработке отдельных строк (%d):", len(stats.row_errors))
        for row_num, name, msg in stats.row_errors:
            log.warning("  %s: %s", name, msg)


def main():
    args = build_arg_parser().parse_args()
    common.setup_logging(args.verbose, args.log_file)
    common.validate_only_person_requires_only_page(args)

    if not os.path.exists(args.excel_file):
        log.error("Файл не найден: %s", args.excel_file)
        sys.exit(1)

    try:
        rows, positions = load_all_rows(args.excel_file, args.sheet_positions, args.photos_root, args.only_page)
    except ValueError as exc:
        log.error(str(exc))
        sys.exit(1)

    if args.only_person is not None:
        rows = [r for r in rows if r.person_number == args.only_person]
        if not rows:
            log.error("На странице '%s' человек с №%s не найден", args.only_page, args.only_person)
            sys.exit(1)

    for emp in rows:
        for w in emp.warnings:
            log.warning("%s: %s", emp.display_name, w)
        for e in emp.fatal_errors:
            log.error("%s: %s — строка пропущена", emp.display_name, e)

    valid_rows = [r for r in rows if r.is_valid]
    skipped_rows = len(rows) - len(valid_rows)

    eligible_rows = [r for r in valid_rows if r.eligible]
    already_in_system = [r for r in valid_rows if not r.eligible]

    for k, names in find_duplicate_keys_in_rows(eligible_rows).items():
        log.warning("Номер карты %s указан у нескольких человек в этом запуске: %s", k, ", ".join(names))

    log.info(
        "Разобрано строк: %d, к импорту (Карточка в системе = Нет): %d, "
        "уже в системе: %d, пропущено из-за ошибок: %d",
        len(rows), len(eligible_rows), len(already_in_system), skipped_rows,
    )
    if not args.allow_partial:
        log.info(
            "Режим «всё или ничего»: если хоть одна из подлежащих импорту строк не "
            "пройдёт проверку, в БД не сохранится ничего. --allow-partial отключает это."
        )
    if not eligible_rows:
        log.error("Нет ни одной строки для импорта (для всех 'Карточка в системе' != 'Нет', либо ошибки).")
        sys.exit(1)

    conn = get_connection(args)
    stats = Stats()
    aborted = False
    try:
        with conn.cursor() as cur:
            position_ids = {r.position_id for r in eligible_rows if r.position_id}
            group_ids = {r.group_id for r in eligible_rows if r.group_id}
            missing_group_ids, missing_position_ids = preflight_check_ids(cur, position_ids, group_ids)
            if missing_group_ids:
                pages = sorted({r.page for r in eligible_rows if r.group_id in missing_group_ids})
                log.error(
                    "В целевой БД НЕ найдены группы employeegroup с Id: %s (страницы: %s)",
                    ", ".join(sorted(missing_group_ids)), ", ".join(pages),
                )
            if missing_position_ids:
                bad_names = sorted({n for n, i in positions.items() if i in missing_position_ids})
                log.error(
                    "В целевой БД НЕ найдены должности employeeposition с Id для: %s",
                    ", ".join(bad_names),
                )

            filtered_rows = []
            for emp in eligible_rows:
                if emp.group_id in missing_group_ids:
                    log.error("%s: группа отсутствует в целевой БД — пропущена", emp.display_name)
                    skipped_rows += 1
                    continue
                if emp.position_id and emp.position_id in missing_position_ids:
                    log.error("%s: должность отсутствует в целевой БД — пропущена", emp.display_name)
                    skipped_rows += 1
                    continue
                filtered_rows.append(emp)

            log.info("-" * 66)
            for emp in filtered_rows:
                cur.execute("SAVEPOINT row_sp")
                try:
                    process_row(cur, emp, stats)
                    cur.execute("RELEASE SAVEPOINT row_sp")
                except Exception as exc:
                    cur.execute("ROLLBACK TO SAVEPOINT row_sp")
                    cur.execute("RELEASE SAVEPOINT row_sp")
                    stats.row_errors.append((emp.excel_row, emp.display_name, str(exc)))
                    log.error("%s: ОШИБКА — %s", emp.display_name, exc)

        has_problems = bool(skipped_rows or stats.row_errors or stats.keys_conflict)

        if args.dry_run:
            conn.rollback()
            if has_problems:
                log.info(
                    "DRY RUN: найдены проблемы (см. выше) — при обычном запуске (без "
                    "--allow-partial) импорт был бы отменён целиком. В БД ничего не изменилось."
                )
            else:
                log.info(
                    "DRY RUN: проблем не найдено, все строки прошли бы успешно. "
                    "В БД ничего не изменилось (ROLLBACK)."
                )
        elif has_problems and not args.allow_partial:
            conn.rollback()
            aborted = True
            log.error("=" * 66)
            log.error(
                "ИМПОРТ ОТМЕНЁН: найдены проблемы (см. выше) — в БД НЕ сохранено ничего, "
                "включая полностью корректные строки."
            )
            log.error(
                "Исправьте перечисленные строки в Excel и запустите заново, либо добавьте "
                "флаг --allow-partial, чтобы сохранить валидные строки, а проблемные пропустить."
            )
        else:
            conn.commit()
            log.info("Изменения сохранены (COMMIT).")
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    print_summary(stats, skipped_rows, len(already_in_system), args.dry_run, aborted)

    if stats.row_errors or stats.keys_conflict or skipped_rows:
        sys.exit(2)
    sys.exit(0)


if __name__ == "__main__":
    main()