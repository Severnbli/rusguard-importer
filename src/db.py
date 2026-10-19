# -*- coding: utf-8 -*-
"""Работа с PostgreSQL RusGuard."""

from __future__ import annotations

try:
    import psycopg2
except ImportError:
    raise SystemExit("Не найден пакет psycopg2. Установите: pip install psycopg2-binary")

EMPLOYEE_IMAGE_TYPE_CARD = 0
PHOTO_NUMBER_DEFAULT = 1


def connect(cfg: dict):
    try:
        port = int(cfg["port"])
    except (TypeError, ValueError):
        raise RuntimeError("порт БД должен быть числом")
    if not cfg.get("dbname") or not cfg.get("user"):
        raise RuntimeError("в настройках не заданы имя базы или пользователь")
    try:
        conn = psycopg2.connect(
            host=cfg["host"], port=port, dbname=cfg["dbname"],
            user=cfg["user"], password=cfg["password"], connect_timeout=10,
        )
    except psycopg2.OperationalError as exc:
        raise RuntimeError(f"не удалось подключиться к БД: {exc}") from exc
    conn.autocommit = False
    return conn


def fetch_positions(cur) -> list:
    cur.execute("SELECT _id::text, name FROM employeeposition ORDER BY name")
    return cur.fetchall()


def fetch_groups(cur) -> list:
    cur.execute("SELECT _id::text, name FROM employeegroup ORDER BY name")
    return cur.fetchall()


def fetch_lookups(cfg: dict) -> tuple:
    """-> (должности, группы) с сервера, каждая — список (id, название)."""
    conn = connect(cfg)
    try:
        with conn.cursor() as cur:
            return fetch_positions(cur), fetch_groups(cur)
    finally:
        conn.close()


def find_employees(cur, lastname, firstname, secondname) -> list:
    cur.execute(
        """
        SELECT _id FROM employee
        WHERE isremoved = false
          AND lastname IS NOT DISTINCT FROM %s
          AND firstname IS NOT DISTINCT FROM %s
          AND secondname IS NOT DISTINCT FROM %s
        """,
        (lastname or None, firstname or None, secondname),
    )
    return [r[0] for r in cur.fetchall()]


def create_employee(cur, row):
    cur.execute(
        """
        INSERT INTO employee
            (lastname, firstname, secondname, employeegroupid, employeepositionid,
             isaccesslevelsinherited)
        VALUES (%s, %s, %s, %s, %s, true)
        RETURNING _id
        """,
        (row.lastname or None, row.firstname or None, row.secondname, row.group_id, row.position_id),
    )
    return cur.fetchone()[0]


def ensure_key_assigned(cur, employee_id, keynumber: int, key_label: str) -> str:
    """-> created | reused | already | conflict."""
    cur.execute("SELECT 1 FROM acskeys WHERE keynumber = %s", (keynumber,))
    existed = cur.fetchone() is not None
    if not existed:
        cur.execute("INSERT INTO acskeys (keynumber, name) VALUES (%s, %s)", (keynumber, key_label))

    cur.execute(
        "SELECT employeeid FROM acskey2employeeassignment "
        "WHERE acskeyid = %s AND assignmentmodificationtype = 0",
        (keynumber,),
    )
    holder = cur.fetchone()
    if holder is not None:
        return "already" if str(holder[0]) == str(employee_id) else "conflict"

    cur.execute(
        "SELECT COALESCE(MAX(indexnumberofassignedkey), 0) + 1 FROM acskey2employeeassignment "
        "WHERE employeeid = %s AND assignmentmodificationtype = 0",
        (employee_id,),
    )
    next_index = cur.fetchone()[0]
    cur.execute(
        "INSERT INTO acskey2employeeassignment "
        "(employeeid, acskeyid, assignmentmodificationtype, indexnumberofassignedkey) "
        "VALUES (%s, %s, 0, %s)",
        (employee_id, keynumber, next_index),
    )
    return "reused" if existed else "created"


def upsert_photo(cur, employee_id, photo: bytes) -> None:
    cur.execute(
        "SELECT _id FROM employeephoto WHERE employeeid = %s AND photonumber = %s "
        "ORDER BY modificationdatetime DESC LIMIT 1",
        (employee_id, PHOTO_NUMBER_DEFAULT),
    )
    existing = cur.fetchone()
    if existing is not None:
        cur.execute(
            "UPDATE employeephoto SET photo = %s, employeeimagetype = %s WHERE _id = %s",
            (psycopg2.Binary(photo), EMPLOYEE_IMAGE_TYPE_CARD, existing[0]),
        )
        return
    cur.execute(
        "INSERT INTO employeephoto (employeeid, photonumber, photo, employeeimagetype) "
        "VALUES (%s, %s, %s, %s)",
        (employee_id, PHOTO_NUMBER_DEFAULT, psycopg2.Binary(photo), EMPLOYEE_IMAGE_TYPE_CARD),
    )


def search_employees(cfg: dict, text: str, limit: int = 100) -> list:
    """Поиск активных сотрудников по ФИО -> [(id, ФИО, должность)]."""
    pattern = "%" + text.strip().replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
    conn = connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT e._id::text, concat_ws(' ', e.lastname, e.firstname, e.secondname) AS fio, p.name
                FROM employee e LEFT JOIN employeeposition p ON p._id = e.employeepositionid
                WHERE e.isremoved = false
                  AND concat_ws(' ', e.lastname, e.firstname, e.secondname) ILIKE %s
                ORDER BY fio LIMIT %s
                """,
                (pattern, limit),
            )
            return cur.fetchall()
    finally:
        conn.close()


def fetch_card(cur, employee_id) -> dict:
    """ФИО, должность и последнее фото сотрудника — данные для карточки в Access."""
    cur.execute(
        """
        SELECT concat_ws(' ', e.lastname, e.firstname, e.secondname), p.name
        FROM employee e LEFT JOIN employeeposition p ON p._id = e.employeepositionid
        WHERE e._id = %s
        """,
        (employee_id,),
    )
    fullname, position = cur.fetchone()
    cur.execute(
        "SELECT photo FROM employeephoto WHERE employeeid = %s "
        "ORDER BY modificationdatetime DESC LIMIT 1",
        (employee_id,),
    )
    photo_row = cur.fetchone()
    return {
        "fullname": fullname,
        "position": position,
        "photo": bytes(photo_row[0]) if photo_row and photo_row[0] else None,
    }
