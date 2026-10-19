# -*- coding: utf-8 -*-
"""Настройки (JSON в профиле пользователя) и пароль приложения.

Пароль БД хранится зашифрованным ключом, выведенным из пароля приложения; сам пароль приложения
не хранится — только проверочный токен. Забытый пароль восстановить нельзя: reset() стирает файл."""

from __future__ import annotations

import base64
import hashlib
import json
import os
from pathlib import Path
from typing import Optional

from cryptography.fernet import Fernet, InvalidToken

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
_INTERNAL = ("auth_salt", "auth_check", "password_enc")
_ITERATIONS = 400_000
MIN_PASSWORD_LENGTH = 4

_key: Optional[bytes] = None  # ключ шифрования, пока приложение разблокировано


def settings_path() -> Path:
    base = os.environ.get("APPDATA") if os.name == "nt" else None
    root = Path(base) if base else Path.home() / ".config"
    return root / "rusguard-importer" / "settings.json"


def _read_raw() -> dict:
    try:
        return json.loads(settings_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _write_raw(raw: dict) -> None:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    if os.name != "nt":
        path.chmod(0o600)


def _derive(password: str, salt: bytes) -> bytes:
    return base64.urlsafe_b64encode(hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _ITERATIONS))


def has_app_password() -> bool:
    return "auth_salt" in _read_raw()


def set_app_password(password: str) -> None:
    """Первичная установка пароля; ранее сохранённый открытый пароль БД шифруется."""
    global _key
    raw = _read_raw()
    salt = os.urandom(16)
    key = _derive(password, salt)
    fernet = Fernet(key)
    raw["auth_salt"] = base64.b64encode(salt).decode()
    raw["auth_check"] = fernet.encrypt(b"ok").decode()
    plain_db_password = raw.pop("password", "")
    raw["password_enc"] = fernet.encrypt(plain_db_password.encode("utf-8")).decode() if plain_db_password else ""
    _write_raw(raw)
    _key = key


def unlock(password: str) -> bool:
    global _key
    raw = _read_raw()
    try:
        key = _derive(password, base64.b64decode(raw["auth_salt"]))
        if Fernet(key).decrypt(raw["auth_check"].encode()) != b"ok":
            return False
    except (KeyError, ValueError, InvalidToken):
        return False
    _key = key
    return True


def reset() -> None:
    """Забыли пароль: удаляются все сохранённые данные."""
    global _key
    _key = None
    try:
        settings_path().unlink()
    except FileNotFoundError:
        pass


def load() -> dict:
    raw = _read_raw()
    data = dict(DEFAULTS)
    data.update({k: v for k, v in raw.items() if k in DEFAULTS and k != "password"})
    if raw.get("password_enc") and _key:
        try:
            data["password"] = Fernet(_key).decrypt(raw["password_enc"].encode()).decode("utf-8")
        except InvalidToken:
            pass
    return data


def save(data: dict) -> None:
    raw = {k: v for k, v in _read_raw().items() if k in _INTERNAL}
    raw.update({k: v for k, v in data.items() if k in DEFAULTS and k != "password"})
    password = data.get("password", "")
    if _key and password:
        raw["password_enc"] = Fernet(_key).encrypt(password.encode("utf-8")).decode()
    else:
        raw["password_enc"] = ""
    _write_raw(raw)
