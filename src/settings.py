# -*- coding: utf-8 -*-
"""Редко меняющиеся настройки (подключение к БД, параметры Access) — JSON в профиле пользователя."""

from __future__ import annotations

import json
import os
from pathlib import Path

DEFAULTS = {
    "host": "localhost",
    "port": "5432",
    "dbname": "",
    "user": "",
    "password": "",
    "access_table": "Import",
    "access_max_photo_dimension": "800",
    "access_position_case": "lower",
    "key_format": "dec",
    "photo_ratio": "3:4",
    "photo_max_side": "800",
    "ui_scale": "100",
}


def settings_path() -> Path:
    base = os.environ.get("APPDATA") if os.name == "nt" else None
    root = Path(base) if base else Path.home() / ".config"
    return root / "rusguard-importer" / "settings.json"


def load() -> dict:
    data = dict(DEFAULTS)
    try:
        data.update(json.loads(settings_path().read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    return data


def save(data: dict) -> None:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    if os.name != "nt":
        path.chmod(0o600)  # в файле лежит пароль
