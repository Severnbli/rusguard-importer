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


def search_employees(cfg: dict, text: str, key_numbers=(), limit: int = 200) -> list:
    """Активные сотрудники, у которых все слова запроса входят в ФИО либо назначен один из key_numbers
    -> [(id, ФИО, должность, [ключи])]. Фильтрация по ФИО в Python: ILIKE не работает
    с недетерминированными правилами сортировки."""
    words = text.casefold().split()
    conn = connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT e._id::text, concat_ws(' ', e.lastname, e.firstname, e.secondname), p.name
                FROM employee e LEFT JOIN employeeposition p ON p._id = e.employeepositionid
                WHERE e.isremoved = false
                """
            )
            employees = cur.fetchall()
            holders = set()
            if key_numbers:
                cur.execute(
                    "SELECT employeeid::text FROM acskey2employeeassignment "
                    "WHERE assignmentmodificationtype = 0 AND acskeyid = ANY(%s)",
                    (list(key_numbers),),
                )
                holders = {r[0] for r in cur.fetchall()}
            found = [
                e for e in employees
                if e[0] in holders or (words and all(w in (e[1] or "").casefold() for w in words))
            ]
            found.sort(key=lambda e: (e[1] or "").casefold())
            found = found[:limit]
            keys_by_employee: dict = {}
            if found:
                cur.execute(
                    "SELECT employeeid::text, acskeyid FROM acskey2employeeassignment "
                    "WHERE assignmentmodificationtype = 0 AND employeeid = ANY(%s::uuid[])",
                    ([e[0] for e in found],),
                )
                for emp_id, key in cur.fetchall():
                    keys_by_employee.setdefault(emp_id, []).append(key)
            return [(i, fio, pos, keys_by_employee.get(i, [])) for i, fio, pos in found]
    finally:
        conn.close()


def fetch_profiles(cfg: dict, employee_ids: list) -> dict:
    """Текущие данные профилей {id: {position, group, photo, keys}} для режима редактирования."""
    result = {i: {"position": "", "group": "", "photo": None, "keys": []} for i in employee_ids}
    if not employee_ids:
        return result
    conn = connect(cfg)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT e._id::text, p.name, g.name
                FROM employee e
                LEFT JOIN employeeposition p ON p._id = e.employeepositionid
                LEFT JOIN employeegroup g ON g._id = e.employeegroupid
                WHERE e._id = ANY(%s::uuid[])
                """,
                (employee_ids,),
            )
            for emp_id, position, group in cur.fetchall():
                result[emp_id]["position"], result[emp_id]["group"] = position or "", group or ""
            cur.execute(
                "SELECT DISTINCT ON (employeeid) employeeid::text, photo FROM employeephoto "
                "WHERE employeeid = ANY(%s::uuid[]) ORDER BY employeeid, modificationdatetime DESC",
                (employee_ids,),
            )
            for emp_id, photo in cur.fetchall():
                if photo:
                    result[emp_id]["photo"] = bytes(photo)
            cur.execute(
                "SELECT employeeid::text, acskeyid FROM acskey2employeeassignment "
                "WHERE assignmentmodificationtype = 0 AND employeeid = ANY(%s::uuid[])",
                (employee_ids,),
            )
            for emp_id, key in cur.fetchall():
                result[emp_id]["keys"].append(key)
    finally:
        conn.close()
    return result


def update_employee(cur, employee_id, position_id, group_id) -> None:
    """Меняет только переданные (не None) должность/группу."""
    parts, values = [], []
    if position_id:
        parts.append("employeepositionid = %s")
        values.append(position_id)
    if group_id:
        parts.append("employeegroupid = %s")
        values.append(group_id)
    if parts:
        cur.execute(f"UPDATE employee SET {', '.join(parts)} WHERE _id = %s", (*values, employee_id))


def unassign_key(cur, employee_id, keynumber: int) -> bool:
    """Снимает активное назначение ключа сотруднику (сам ключ остаётся в acskeys)."""
    cur.execute(
        "DELETE FROM acskey2employeeassignment "
        "WHERE employeeid = %s AND acskeyid = %s AND assignmentmodificationtype = 0",
        (employee_id, keynumber),
    )
    return cur.rowcount > 0


def fetch_cards(cfg: dict, employee_ids: list) -> dict:
    """{id: карточка для Access} для сотрудников из БД."""
    conn = connect(cfg)
    try:
        with conn.cursor() as cur:
            return {i: fetch_card(cur, i) for i in employee_ids}
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
