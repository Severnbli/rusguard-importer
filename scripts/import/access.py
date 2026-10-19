#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
access.py — добавление ФИО, Должности и Фото из Excel в Access
(карточки.accdb). Самостоятельный скрипт, не зависит от rusguard_import.py
и от Postgres — запускается отдельно.

СТРУКТУРА ВХОДНОГО EXCEL-ФАЙЛА
  Один лист — справочник должностей (по умолчанию "Должности", см.
  --sheet-positions). Этот скрипт его СОДЕРЖИМОЕ не читает вообще (Access
  не хранит Id должности) — только использует имя листа, чтобы не принять
  его за страницу с людьми.

  ВСЕ ОСТАЛЬНЫЕ листы — «страницы» с людьми, каждая обрабатывается
  независимо. Столбцы (ищутся по заголовку, регистр не важен):
      №                     — порядковый номер человека на странице.
                              ОБЯЗАТЕЛЕН, уникален на странице (сопоставление
                              с файлом фото).
      ФИО                   — "Фамилия Имя Отчество". ОБЯЗАТЕЛЬНО.
      Должность             — записывается в Access КАК ЕСТЬ (текстом, без
                              проверки по справочнику). Можно оставить пустым.
      Карточка изготовлена  — "Да"/"Нет". Добавляются в Access ТОЛЬКО строки
                              со значением "Нет". ОБЯЗАТЕЛЬНЫЙ столбец.
      Путь к фото           — папка с фото, ОДНА НА ВСЮ СТРАНИЦУ (читается
                              из любой заполненной ячейки этого столбца).
                              Необязательно. Абсолютный Windows-путь (с
                              буквой диска или \\\\UNC) — как есть; любой
                              другой (в т.ч. "/photos/employee") — считается
                              ОТНОСИТЕЛЬНО --photos-root. Файлы фото:
                              "<№>.<расширение>" (jpg/jpeg/png/bmp/gif/tif/
                              tiff/webp), сравнение по числу.

Строка, у которой ФИО начинается с "#", полностью пропускается (как
закомментированная).

ЧТО ДЕЛАЕТ
  Добавляет в указанную Access-таблицу (по умолчанию Export) новые записи:
  [ФИО], [Должность], [Фото]. Только ДОБАВЛЕНИЕ: если ФИО уже есть в Access
  (без учёта регистра/пробелов) — строка целиком пропускается, ничего в
  существующей записи не меняется.

ПОЛИТИКА "ВСЁ ИЛИ НИЧЕГО" (по умолчанию)
  Если хотя бы одна ПОДЛЕЖАЩАЯ добавлению строка не проходит проверку
  (пустой Fullname, повторяющееся ФИО в файле, которого ещё нет в Access)
  — в Access не добавляется вообще ничего. --allow-partial отключает это.

ВЫБОРОЧНЫЙ ЗАПУСК
  --only-page "2А"        — обработать только эту страницу.
  --only-person 5         — обработать только человека с №5 (требует
                             --only-page).

ФОРМАТ ПОЛЯ [Фото]
  Реверс-инжинирен по 66 реальным записям присланного образца карточки.accdb
  — обычная Access-вставка "Bitmap Image" (Paint.Picture): 78-байтный
  заголовок + BMP + 12-байтный хвост. Побайтовый round-trip совпал на всех
  66 из 66. Фото любого формата на диске (jpg/png/bmp/…) конвертируется в
  BMP и оборачивается в тот же формат.

ПРИМЕРЫ ЗАПУСКА
    python access.py import.xlsm "C:\\путь\\карточки.accdb" \\
        --photos-root "C:\\Data" --dry-run

    python access.py import.xlsm "C:\\путь\\карточки.accdb" \\
        --photos-root "C:\\Data" --only-page "2А" --only-person 5 --dry-run

ТРЕБОВАНИЯ
  pip install openpyxl Pillow pyodbc
  Плюс драйвер "Microsoft Access Driver (*.mdb, *.accdb)" (Microsoft Access
  Database Engine Redistributable) — только Windows; разрядность (32/64-бит)
  должна совпадать с разрядностью Python.
"""

from __future__ import annotations

import argparse
import io
import logging
import os
import struct
import sys
from dataclasses import dataclass, field
from typing import Optional

try:
    import pyodbc
    _PYODBC_AVAILABLE = True
except ImportError:
    _PYODBC_AVAILABLE = False

try:
    from PIL import Image as PILImage
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False


try:
    import common
except ImportError:
    sys.exit("Не найден common.py. Положите common.py в ту же папку, что и access.py.")

log = logging.getLogger("access")

# ----------------------------------------------------------------------
# Формат поля [Фото] (OLE Object) в карточки.accdb, таблица [Export].
# Реверс-инжинирен по 66 существующим записям — 78-байтный заголовок
# (внедрённый объект класса "Paint.Picture" / "Точечный рисунок"), 4 байта
# (LE uint32) — длина встроенного BMP, сам BMP, 12-байтный хвост (8
# постоянных байт + 4 байта, не влияющих на чтение). Проверено побайтовым
# round-trip на всех 66 записях образца — совпадение 100%.
# ----------------------------------------------------------------------
_ACCESS_OLE_FIXED_HEADER = bytes.fromhex(
    "151c33000200000011000e0014002500ffffffffd2eef7e5f7edfbe920f0e8f1f3edeeea"
    "005061696e742e50696374757265000105000002000000070000005042727573680000"
    "00000000000000"
)
_ACCESS_OLE_TRAILER_PREFIX = bytes.fromhex("0105000000000000")


# ======================================================================
#  Разбор Excel
# ======================================================================

@dataclass
class PersonRow(common.BasePersonRow):
    position_raw: Optional[str] = None


def load_page_rows(ws, page_name: str, photos_root: Optional[str]):
    """-> (rows: list[PersonRow], page_errors: list[str])."""
    header = next(ws.iter_rows(min_row=1, max_row=1))
    col_num = common.find_col(header, "№", "номер", "n")
    col_fio = common.find_col(header, "фио", "fullname")
    col_card_made = common.find_col(header, "карточка изготовлена")
    col_position = common.find_col(header, "должность", "position")
    col_photo_path = common.find_col(header, "путь к фото")

    required = [
        ("№", col_num), ("ФИО", col_fio),
        ("Карточка изготовлена", col_card_made), ("Должность", col_position),
    ]
    missing = [n for n, c in required if c is None]
    if missing:
        return [], [f"на странице '{page_name}' не найдены обязательные столбцы: {', '.join(missing)}"]

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

    rows: list[PersonRow] = []
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
            # строка сознательно закомментирована — саму запись пропускаем, но
            # № остаётся "занят" (не осиротевшее фото, нельзя переиспользовать
            # без предупреждения).
            skipped_hash_count += 1
            if person_number is not None:
                seen_numbers[person_number] = excel_row
            continue

        r = PersonRow(page=page_name, excel_row=excel_row, person_number=person_number,
                       fullname_raw=fullname_raw)

        if person_number is None:
            r.fatal_errors.append("не заполнен или нечисловой №")
        elif person_number in seen_numbers:
            r.fatal_errors.append(
                f"№ {person_number} повторяется на этой странице (первый раз — "
                f"строка {seen_numbers[person_number]})"
            )
        else:
            seen_numbers[person_number] = excel_row

        if not fullname_raw:
            r.fatal_errors.append("не заполнен ФИО")
            rows.append(r)
            continue

        pos_cell = row[col_position - 1]
        r.position_raw = str(pos_cell.value).strip() if pos_cell.value not in (None, "") else None

        r.eligible = common.is_flag_no(row[col_card_made - 1].value, r.warnings, "Карточка изготовлена")

        if person_number is not None and photo_index:
            photo_path = photo_index.get(person_number)
            if photo_path:
                try:
                    data = photo_path.read_bytes()
                except OSError as exc:
                    r.warnings.append(f"не удалось прочитать файл фото {photo_path}: {exc}")
                    data = None
                if data:
                    problem = common.validate_photo_bytes(data)
                    if problem:
                        r.warnings.append(
                            f"файл фото {photo_path.name} повреждён/не распознан ({problem}) — фото пропущено"
                        )
                    else:
                        r.photo_bytes = data

        rows.append(r)

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

    all_rows: list[PersonRow] = []
    for page_name in page_names:
        rows, page_errors = load_page_rows(wb[page_name], page_name, photos_root)
        if page_errors:
            for e in page_errors:
                log.error("%s", e)
            all_rows.append(PersonRow(
                page=page_name, excel_row=1, person_number=None, fullname_raw="",
                fatal_errors=list(page_errors),
            ))
            continue
        log.info("Страница '%s': %d строк", page_name, len(rows))
        all_rows.extend(rows)

    return all_rows


# ======================================================================
#  Access — добавление ФИО+Должность+Фото, без изменения существующих
# ======================================================================

@dataclass
class AccessStats:
    inserted: int = 0
    already_existed: int = 0
    problems: list = field(default_factory=list)
    skipped_entirely: bool = False


def image_bytes_to_access_ole(image_bytes: bytes, max_dimension: Optional[int] = None) -> bytes:
    if not _PIL_AVAILABLE:
        raise RuntimeError("нужен Pillow: pip install Pillow")
    im = PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
    if max_dimension:
        w, h = im.size
        scale = min(1.0, max_dimension / max(w, h))
        if scale < 1.0:
            im = im.resize((max(1, round(w * scale)), max(1, round(h * scale))), PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="BMP")
    bmp_bytes = buf.getvalue()
    return (
        _ACCESS_OLE_FIXED_HEADER
        + struct.pack("<I", len(bmp_bytes))
        + bmp_bytes
        + _ACCESS_OLE_TRAILER_PREFIX
        + b"\x00\x00\x00\x00"
    )


def get_access_connection(access_db_path: str):
    if not _PYODBC_AVAILABLE:
        log.error("Не найден пакет pyodbc. Установите: pip install pyodbc")
        log.error(
            "Также нужен драйвер 'Microsoft Access Driver (*.mdb, *.accdb)' (Microsoft "
            "Access Database Engine Redistributable) — его разрядность (32/64-бит) "
            "должна совпадать с разрядностью Python."
        )
        sys.exit(1)
    conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" rf"DBQ={access_db_path};"
    try:
        return pyodbc.connect(conn_str, autocommit=False)
    except pyodbc.Error as exc:
        log.error("Не удалось подключиться к Access-базе %s: %s", access_db_path, exc)
        sys.exit(1)


def load_access_existing_names(cursor, table: str) -> set:
    cursor.execute(f"SELECT [ФИО] FROM [{table}]")
    return {str(row[0]).strip().lower() for row in cursor.fetchall() if row[0]}


def plan_access_inserts(valid_rows, existing_names: set):
    """-> (to_insert: list[PersonRow], problems: list[str])."""
    from collections import defaultdict

    groups: dict = defaultdict(list)
    for r in valid_rows:
        key = r.fullname_raw.strip().lower()
        if key in existing_names:
            continue
        groups[key].append(r)

    to_insert = []
    problems = []
    for rs in groups.values():
        if len(rs) > 1:
            rows_str = ", ".join(str(r.excel_row) for r in rs)
            problems.append(
                f"'{rs[0].fullname_raw}' встречается в файле {len(rs)} раза "
                f"(строки {rows_str}) и ещё не заведён(а) в Access — не добавлено, "
                f"т.к. неясно, чьё фото/должность использовать"
            )
            continue
        to_insert.append(rs[0])
    return to_insert, problems


def insert_access_row(cursor, table: str, fullname: str, position_text: Optional[str],
                       photo_blob: Optional[bytes]):
    cols = ["[ФИО]"]
    vals = [fullname]
    if position_text:
        cols.append("[Должность]")
        vals.append(position_text)
    if photo_blob is not None:
        cols.append("[Фото]")
        vals.append(photo_blob)
    placeholders = ", ".join("?" for _ in vals)
    cursor.execute(f"INSERT INTO [{table}] ({', '.join(cols)}) VALUES ({placeholders})", *vals)


def run_access_phase(args, valid_rows, overall_has_problems: bool) -> AccessStats:
    astats = AccessStats()
    if overall_has_problems and not args.allow_partial:
        astats.skipped_entirely = True
        log.info(
            "Access: пропущено целиком — найдены проблемы в Excel (см. выше). Ничего не добавлено."
        )
        return astats

    log.info("-" * 66)
    log.info("Access: %s (таблица [%s])", args.access_db, args.access_table)
    conn = get_access_connection(args.access_db)
    try:
        cur = conn.cursor()
        existing_names = load_access_existing_names(cur, args.access_table)
        to_insert, plan_problems = plan_access_inserts(valid_rows, existing_names)
        astats.problems.extend(plan_problems)
        for msg in plan_problems:
            log.warning("Access: %s", msg)
        astats.already_existed = sum(
            1 for r in valid_rows if r.fullname_raw.strip().lower() in existing_names
        )

        for r in to_insert:
            photo_blob = None
            if r.photo_bytes:
                try:
                    photo_blob = image_bytes_to_access_ole(r.photo_bytes, args.access_max_photo_dimension)
                except Exception as exc:
                    msg = f"{r.display_name}: не удалось подготовить фото — {exc}"
                    astats.problems.append(msg)
                    log.warning("Access: %s", msg)

            position_text = r.position_raw
            if position_text and args.access_position_case == "lower":
                position_text = position_text.lower()

            try:
                insert_access_row(cur, args.access_table, r.fullname_raw.strip(), position_text, photo_blob)
                astats.inserted += 1
                bits = []
                if position_text:
                    bits.append("должность")
                bits.append("с фото" if photo_blob else "без фото")
                log.info("  Access: %s — добавлен(а) (%s)", r.display_name, ", ".join(bits))
            except Exception as exc:
                msg = f"{r.display_name}: ошибка вставки в Access — {exc}"
                astats.problems.append(msg)
                log.error("Access: %s", msg)

        if args.dry_run:
            conn.rollback()
            log.info("Access: DRY RUN — изменения отменены (ROLLBACK).")
        elif astats.problems and not args.allow_partial:
            conn.rollback()
            log.error("Access: обнаружены проблемы — изменения отменены (ROLLBACK), ничего не добавлено.")
        else:
            conn.commit()
            log.info("Access: изменения сохранены (COMMIT).")
    finally:
        conn.close()
    return astats


# ======================================================================
#  CLI / main
# ======================================================================

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="access.py",
        description="Добавление ФИО+Должность+Фото из Excel в Access (карточки.accdb).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    common.add_common_cli_args(p)
    p.add_argument("access_db", help="путь к .accdb файлу (например, карточки.accdb)")
    p.add_argument("--access-table", default="Export", help="имя таблицы в Access (по умолчанию Export)")
    p.add_argument(
        "--access-max-photo-dimension", type=int, default=None,
        help="ужать фото до этого размера по большей стороне перед записью "
             "(у .accdb жёсткий лимит 2 ГБ на файл, фото хранятся как BMP без сжатия)",
    )
    p.add_argument(
        "--access-position-case", choices=["lower", "asis"], default="lower",
        help="в каком регистре писать Должность: lower (по умолчанию, как в "
             "существующих записях образца) или asis (как в Excel)",
    )
    return p


def print_summary(astats: AccessStats, skipped_rows: int, already_made: int, dry_run: bool, aborted: bool):
    log.info("=" * 66)
    if dry_run:
        log.info("ИТОГ (DRY RUN — в Access НИЧЕГО не сохранено)")
    elif aborted:
        log.info("ИТОГ (ОТМЕНЕНО — в Access НИЧЕГО не сохранено)")
    else:
        log.info("ИТОГ (сохранено в Access)")
    log.info(
        "Добавлено новых: %d, уже были в базе (не тронуты): %d, "
        "карточка уже изготовлена (пропущено): %d, "
        "пропущено из-за ошибок в Excel: %d, проблем при добавлении: %d",
        astats.inserted, astats.already_existed, already_made, skipped_rows, len(astats.problems),
    )
    if astats.problems:
        log.warning("Проблемы (%d):", len(astats.problems))
        for msg in astats.problems:
            log.warning("  %s", msg)


def main():
    args = build_arg_parser().parse_args()
    common.setup_logging(args.verbose, args.log_file)

    common.validate_only_person_requires_only_page(args)

    if not os.path.exists(args.excel_file):
        log.error("Excel-файл не найден: %s", args.excel_file)
        sys.exit(1)
    if not os.path.exists(args.access_db):
        log.error("Access-файл не найден: %s", args.access_db)
        sys.exit(1)

    try:
        rows = load_all_rows(args.excel_file, args.sheet_positions, args.photos_root, args.only_page)
    except ValueError as exc:
        log.error(str(exc))
        sys.exit(1)

    if args.only_person is not None:
        rows = [r for r in rows if r.person_number == args.only_person]
        if not rows:
            log.error("На странице '%s' человек с №%s не найден", args.only_page, args.only_person)
            sys.exit(1)

    for r in rows:
        for w in r.warnings:
            log.warning("%s: %s", r.display_name, w)
        for e in r.fatal_errors:
            log.error("%s: %s — строка пропущена", r.display_name, e)

    valid_rows = [r for r in rows if r.is_valid]
    skipped_rows = len(rows) - len(valid_rows)

    eligible_rows = [r for r in valid_rows if r.eligible]
    already_made = [r for r in valid_rows if not r.eligible]

    log.info(
        "Разобрано строк: %d, к добавлению (Карточка изготовлена = Нет): %d, "
        "уже изготовлена: %d, пропущено из-за ошибок: %d",
        len(rows), len(eligible_rows), len(already_made), skipped_rows,
    )
    if not args.allow_partial:
        log.info(
            "Режим «всё или ничего»: если хоть одна из подлежащих добавлению строк не "
            "пройдёт проверку, в Access не добавится ничего. --allow-partial отключает это."
        )
    if not eligible_rows:
        log.error("Нет ни одной строки для добавления (для всех 'Карточка изготовлена' != 'Нет', либо ошибки).")
        sys.exit(1)

    has_problems_upfront = bool(skipped_rows)
    astats = run_access_phase(args, eligible_rows, has_problems_upfront)
    aborted = astats.skipped_entirely or (
        bool(astats.problems) and not args.allow_partial and not args.dry_run
    )

    print_summary(astats, skipped_rows, len(already_made), args.dry_run, aborted)

    if skipped_rows or astats.problems:
        sys.exit(2)
    sys.exit(0)


if __name__ == "__main__":
    main()
