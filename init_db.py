#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Создаёт файл SQLite 'app.db', включает foreign_keys и разворачивает схему из schema.sql,
а затем заливает начальные справочники.
"""

import sqlite3
from pathlib import Path
from werkzeug.security import generate_password_hash
DB_FILE = Path("app.db")
SCHEMA_FILE = Path("schema.sql")

DEPARTMENTS = [
    "Производство",
    "ИТ",
    "Отдел продаж",
    "Бухгалтерия",
]

ACTIVITY_TYPES = [
    "Задача",
    "Поручение",
    "Служебная записка",
    "Проект",
]

def run_schema(conn: sqlite3.Connection, schema_path: Path):
    with schema_path.open("r", encoding="utf-8") as f:
        sql = f.read()
    conn.executescript(sql)

def seed_reference(conn: sqlite3.Connection):
    cur = conn.cursor()

    # Departments
    cur.executemany(
        "INSERT OR IGNORE INTO departments(name) VALUES (?)",
        [(d,) for d in DEPARTMENTS]
    )

    # Activity Types
    cur.executemany(
        "INSERT OR IGNORE INTO activity_types(name) VALUES (?)",
        [(t,) for t in ACTIVITY_TYPES]
    )

    conn.commit()

def main():
    # Создаём файл БД (или подключаемся к существующему)
    initializing = not DB_FILE.exists()
    conn = sqlite3.connect(DB_FILE)
    try:
        # Обязательно включаем внешние ключи
        conn.execute("PRAGMA foreign_keys = ON;")

        # Разворачиваем схему только при первом создании или если явно хотим пересоздать
        if initializing:
            print("▶ Разворачиваем схему из schema.sql ...")
            run_schema(conn, SCHEMA_FILE)
            print("✔ Схема создана.")

            print("▶ Сидим справочники ...")
            seed_reference(conn)
            print("✔ Справочники загружены.")
        else:
            print("ℹ База уже существует. Ничего не меняю.")

    finally:
        seed_reference(conn)
        seed_admin(conn)
        conn.close()


def seed_admin(conn):
    pw_hash = generate_password_hash("admin")  # <-- правильный формат
    conn.execute("""
        INSERT OR IGNORE INTO users(full_name, username, password_hash, role, is_approved)
        VALUES ('Администратор', 'admin', ?, 'admin', 1)
    """, (pw_hash,))
    conn.commit()
if __name__ == "__main__":
    main()
