# -*- coding: utf-8 -*-
"""Создание Access-базы из template.accdb и заполнение её карточками напрямую из данных БД."""

from __future__ import annotations

import io
import os
import shutil
import struct
from typing import Optional

try:
    import pyodbc
except ImportError:
    pyodbc = None

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None

from sheets import ACCESS_TEMPLATE

# Формат поля [Фото]: 78-байтный заголовок Paint.Picture + длина BMP + BMP + 12-байтный хвост.
_OLE_HEADER = bytes.fromhex(
    "151c33000200000011000e0014002500ffffffffd2eef7e5f7edfbe920f0e8f1f3edeeea"
    "005061696e742e50696374757265000105000002000000070000005042727573680000"
    "00000000000000"
)
_OLE_TRAILER = bytes.fromhex("0105000000000000") + b"\x00\x00\x00\x00"


def image_to_access_ole(image_bytes: bytes, max_dimension: Optional[int]) -> bytes:
    if PILImage is None:
        raise RuntimeError("нужен Pillow: pip install Pillow")
    im = PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
    if max_dimension:
        scale = min(1.0, max_dimension / max(im.size))
        if scale < 1.0:
            im = im.resize((max(1, round(im.width * scale)), max(1, round(im.height * scale))),
                           PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="BMP")
    bmp = buf.getvalue()
    return _OLE_HEADER + struct.pack("<I", len(bmp)) + bmp + _OLE_TRAILER


def export_cards(dest: str, cards: list, cfg: dict, log) -> None:
    """Копирует template.accdb в dest и добавляет карточки (словари fullname/position/photo).
    При любой ошибке частично созданный файл удаляется и исключение пробрасывается."""
    if pyodbc is None:
        raise RuntimeError(
            "не найден pyodbc (pip install pyodbc) и драйвер 'Microsoft Access Driver "
            "(*.mdb, *.accdb)' — Access-экспорт работает только в Windows"
        )
    if not ACCESS_TEMPLATE.exists():
        raise FileNotFoundError(f"Не найден шаблон Access: {ACCESS_TEMPLATE}")
    max_dim = int(cfg["access_max_photo_dimension"]) if str(cfg["access_max_photo_dimension"]).isdigit() else None
    table = cfg["access_table"]

    shutil.copyfile(ACCESS_TEMPLATE, dest)
    try:
        conn = pyodbc.connect(
            rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={dest};", autocommit=False,
        )
        try:
            cur = conn.cursor()
            for card in cards:
                cols, vals = ["[ФИО]"], [card["fullname"]]
                position = card["position"]
                if position:
                    cols.append("[Должность]")
                    vals.append(position.lower() if cfg["access_position_case"] == "lower" else position)
                if card["photo"]:
                    cols.append("[Фото]")
                    vals.append(image_to_access_ole(card["photo"], max_dim))
                cur.execute(
                    f"INSERT INTO [{table}] ({', '.join(cols)}) VALUES ({', '.join('?' * len(vals))})", *vals,
                )
                log(f"  Access: {card['fullname']} — добавлен(а)")
            conn.commit()
        finally:
            conn.close()
    except Exception:
        if os.path.exists(dest):
            os.remove(dest)
        raise
