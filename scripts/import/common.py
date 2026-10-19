#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
common.py — общая часть rusguard.py и access.py.

Здесь — то, что у двух скриптов было буквально одинаковым: разбор общих
кусков Excel (страницы с людьми, поиск столбца по заголовку, значения
Да/Нет, фото на диске), общие флаги командной строки и настройка
логирования. Ничего специфичного для PostgreSQL или Access здесь нет —
это по-прежнему только в самих rusguard.py/access.py.

Не предназначен для самостоятельного запуска — только для импорта.
"""

from __future__ import annotations

import argparse
import io
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Не найден пакет openpyxl. Установите: pip install openpyxl")

try:
    from PIL import Image as PILImage
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False


log = logging.getLogger("common")

_WINDOWS_ABS_RE = re.compile(r"^[A-Za-z]:[\\/]|^\\\\")
_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}


# ======================================================================
#  Базовая строка "человек на странице"
# ======================================================================

@dataclass
class BasePersonRow:
    """Общая часть EmployeeRow (rusguard.py) и PersonRow (access.py).
    Каждый из скриптов наследуется от неё и добавляет свои поля."""

    page: str
    excel_row: int
    person_number: Optional[int]
    fullname_raw: str
    photo_bytes: Optional[bytes] = None
    eligible: bool = False  # решается по своему столбцу Да/Нет в каждом скрипте
    fatal_errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.fatal_errors

    @property
    def display_name(self) -> str:
        who = self.fullname_raw or f"строка {self.excel_row}"
        num = f"№{self.person_number} " if self.person_number is not None else ""
        return f"[{self.page}] {num}{who}"


# ======================================================================
#  Разбор Excel
# ======================================================================

def find_col(header_row, *names) -> Optional[int]:
    """Индекс столбца (1-based) по названию заголовка, без учёта регистра/пробелов."""
    wanted = {n.strip().lower() for n in names}
    for cell in header_row:
        if cell.value and str(cell.value).strip().lower() in wanted:
            return cell.column
    return None


def is_flag_no(raw_value, warnings_out: list, column_label: str) -> bool:
    """True если ячейка означает 'Нет' (без учёта регистра/пробелов).
    Пустая ячейка или что угодно, кроме 'Да'/'Нет', считается НЕ 'Нет'
    (строка не подлежит обработке); неожиданное непустое значение —
    предупреждение."""
    text = str(raw_value).strip().lower() if raw_value not in (None, "") else ""
    if text == "нет":
        return True
    if text and text != "да":
        warnings_out.append(
            f"в столбце '{column_label}' неожиданное значение {raw_value!r} "
            f"(ожидались 'Да'/'Нет') — считаю как 'Да'"
        )
    return False


def resolve_photo_folder(raw_path: str, photos_root: Optional[str]) -> Path:
    """Абсолютный Windows-путь (с буквой диска или UNC \\\\...) — как есть.
    Иначе — относительно --photos-root."""
    raw_path = raw_path.strip()
    if _WINDOWS_ABS_RE.match(raw_path):
        return Path(raw_path)
    if not photos_root:
        raise ValueError(
            f"путь к фото '{raw_path}' не похож на абсолютный путь Windows (нет буквы "
            f"диска) — укажите --photos-root, относительно которого его понимать"
        )
    return Path(photos_root) / raw_path.lstrip("/\\")


def index_photo_folder(folder: Path) -> dict:
    """{номер_человека: Path} по файлам вида '<номер>.<расширение>' в папке."""
    index: dict[int, Path] = {}
    if not folder.is_dir():
        log.warning("Папка с фото не найдена: %s", folder)
        return index
    dupes = set()
    for f in sorted(folder.iterdir()):
        if not f.is_file() or f.suffix.lower() not in _IMAGE_EXTENSIONS:
            continue
        try:
            n = int(f.stem)
        except ValueError:
            continue
        if n in index:
            dupes.add(n)
        else:
            index[n] = f
    for n in sorted(dupes):
        log.warning("В папке %s несколько файлов для номера %s — использован первый по алфавиту", folder, n)
    return index


def validate_photo_bytes(data: bytes) -> Optional[str]:
    """None если всё ок, иначе текст проблемы."""
    if not _PIL_AVAILABLE:
        return None
    try:
        with PILImage.open(io.BytesIO(data)) as im:
            im.verify()
        return None
    except Exception as exc:
        return str(exc)


def collect_page_wide_value(ws, col_idx: int, label: str, page_name: str) -> tuple:
    """Достаёт единственное значение из столбца col_idx (одно на всю
    страницу). -> (значение_или_None, список_ошибок)."""
    values = set()
    for row in ws.iter_rows(min_row=2):
        v = row[col_idx - 1].value
        if v not in (None, ""):
            values.add(str(v).strip())
    if not values:
        return None, []
    if len(values) > 1:
        return None, [
            f"на странице '{page_name}' указано несколько разных значений '{label}': "
            f"{', '.join(sorted(values))} — непонятно, какое использовать"
        ]
    return next(iter(values)), []


def open_workbook_and_get_pages(path: str, sheet_positions: str, only_page: Optional[str]):
    """Открывает книгу, проверяет наличие листа-справочника должностей,
    возвращает (wb, page_names) — список имён листов с людьми (без листа-
    справочника), при необходимости отфильтрованный до only_page."""
    wb = load_workbook(path, data_only=True, keep_vba=path.lower().endswith(".xlsm"))
    if sheet_positions not in wb.sheetnames:
        raise ValueError(
            f"В файле нет листа-справочника должностей '{sheet_positions}'. "
            f"Найденные листы: {', '.join(wb.sheetnames)}"
        )
    page_names = [s for s in wb.sheetnames if s != sheet_positions]
    if only_page:
        if only_page not in page_names:
            raise ValueError(
                f"Страница '{only_page}' не найдена. Доступные страницы: {', '.join(page_names)}"
            )
        page_names = [only_page]
    return wb, page_names


# ======================================================================
#  CLI / логирование
# ======================================================================

def add_common_cli_args(p: argparse.ArgumentParser):
    """Флаги, одинаковые для rusguard.py и access.py. Каждый скрипт
    добавляет свои специфичные аргументы отдельно (до или после этого
    вызова — порядок влияет только на вид --help, не на поведение)."""
    p.add_argument("excel_file", help="путь к .xlsx/.xlsm файлу")
    p.add_argument(
        "--sheet-positions", default="Должности",
        help="имя листа-справочника должностей (все остальные листы считаются страницами с людьми)",
    )
    p.add_argument(
        "--photos-root", default=None,
        help="базовая папка для путей к фото, не являющихся полным Windows-путём",
    )
    p.add_argument("--only-page", default=None, help="обработать только одну страницу (лист)")
    p.add_argument(
        "--only-person", type=int, default=None,
        help="обработать только человека с этим №; требует --only-page",
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="выполнить всё в транзакции и отменить (ROLLBACK) — изменения не сохранятся",
    )
    p.add_argument(
        "--allow-partial", action="store_true",
        help="обработать валидные строки, даже если другие не прошли проверку или есть "
             "конфликты. По умолчанию выключено: при любой проблеме хотя бы в одной "
             "строке не сохраняется вообще ничего.",
    )
    p.add_argument("--verbose", action="store_true", help="подробный вывод по каждой строке (DEBUG)")
    p.add_argument("--log-file", default=None, help="дополнительно писать лог в файл")


def validate_only_person_requires_only_page(args):
    """sys.exit с понятным сообщением, если --only-person задан без --only-page."""
    if args.only_person is not None and not args.only_page:
        sys.exit("--only-person требует также указать --only-page")


def setup_logging(verbose: bool, log_file: Optional[str]):
    level = logging.DEBUG if verbose else logging.INFO
    handlers = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(
        level=level, format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S", handlers=handlers, force=True,
    )
