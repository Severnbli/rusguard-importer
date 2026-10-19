# -*- coding: utf-8 -*-
"""Разбор номеров ключей (карт) в трёх форматах: десятичный, HEX, HEX с обратным порядком байт."""

from __future__ import annotations

import re

KEY_FORMATS = {
    "dec": "Десятичный",
    "hex": "HEX",
    "hex_rev": "HEX (обратный порядок байт)",
}

_SPLIT_RE = re.compile(r"[,;/\n]+")
_HEX_PREFIX_RE = re.compile(r"^0[xX]")
_HEX_JUNK_RE = re.compile(r"[\s:_-]")


def hex_to_keynumber(text: str, reverse_bytes: bool) -> int:
    """UID в hex -> keynumber (bigint без знака, как в SQL-инструкции). ValueError при проблеме."""
    clean = _HEX_JUNK_RE.sub("", _HEX_PREFIX_RE.sub("", text.strip())).upper()
    if not clean or not re.fullmatch(r"[0-9A-F]+", clean):
        raise ValueError(f"'{text}' не является шестнадцатеричным числом")
    if len(clean) > 16:
        raise ValueError(f"'{text}' длиннее 8 байт (16 hex-цифр)")
    if reverse_bytes:
        if len(clean) % 2:
            raise ValueError(f"'{text}': для разворота байт нужно чётное число hex-цифр")
        clean = "".join(reversed([clean[i:i + 2] for i in range(0, len(clean), 2)]))
    value = int(clean, 16)
    if value >= 2 ** 63:
        raise ValueError(f"'{text}' не помещается в bigint (значение >= 2^63)")
    if value == 0:
        raise ValueError(f"'{text}' равен нулю")
    return value


def parse_keys(value, key_format: str) -> tuple:
    """Содержимое ячейки -> (список номеров без повторов, список текстов ошибок)."""
    if value in (None, "") or isinstance(value, bool):
        return [], []
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    parts = [p.strip() for p in _SPLIT_RE.split(str(value)) if p.strip()]
    numbers, errors = [], []
    for part in parts:
        try:
            if key_format == "dec":
                digits = part.replace(" ", "")
                if not digits.isdigit() or int(digits) <= 0:
                    raise ValueError(f"'{part}' не является положительным числом")
                number = int(digits)
            else:
                number = hex_to_keynumber(part, reverse_bytes=(key_format == "hex_rev"))
        except ValueError as exc:
            errors.append(str(exc))
            continue
        if number not in numbers:
            numbers.append(number)
    return numbers, errors


def format_key(number: int, key_format: str) -> str:
    """Обратное преобразование номера ключа в текст выбранного формата."""
    if key_format == "dec":
        return str(number)
    digits = f"{number:X}"
    if len(digits) % 2:
        digits = "0" + digits
    if key_format == "hex_rev":
        digits = "".join(reversed([digits[i:i + 2] for i in range(0, len(digits), 2)]))
    return digits


def reformat_keys(text: str, old_format: str, new_format: str) -> str:
    """Переводит список ключей в тексте из одного формата в другой (при ошибке разбора — без изменений)."""
    numbers, errors = parse_keys(text, old_format)
    if errors or not numbers:
        return text
    return ", ".join(format_key(n, new_format) for n in numbers)
