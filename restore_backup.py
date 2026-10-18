"""استرجاع نسخة احتياطية (JSON) إلى قاعدة البيانات الحالية.

الاستخدام:
    python restore_backup.py backup.json

يضيف الطلبات غير الموجودة فقط (حسب رقم التسجيل)، ويستبدل الإعدادات والأيام المخصصة.
"""
from __future__ import annotations

import json
import sys

import app as booking


def restore(path: str) -> None:
    data = json.load(open(path, encoding="utf-8"))
    conn = booking.connect()
    columns = {
        r["name"] if "name" in r else r[1]
        for r in (
            conn.execute(
                "SELECT column_name AS name FROM information_schema.columns"
                " WHERE table_schema = ? AND table_name = 'registrations'",
                (booking.PG_SCHEMA,),
            )
            if booking.USE_POSTGRES
            else conn.execute("PRAGMA table_info(registrations)")
        )
    }
    conn.execute("BEGIN IMMEDIATE")
    added = skipped = 0
    for row in data.get("registrations", []):
        if conn.execute("SELECT 1 FROM registrations WHERE code = ?", (row["code"],)).fetchone():
            skipped += 1
            continue
        record = {k: v for k, v in row.items() if k in columns and k != "id"}
        conn.execute(
            f"INSERT INTO registrations ({', '.join(record)}) VALUES ({', '.join('?' * len(record))})",
            list(record.values()),
        )
        added += 1
    for row in data.get("settings", []):
        if row["key"] == "secret_key":
            continue
        conn.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?)"
            " ON CONFLICT (key) DO UPDATE SET value = excluded.value",
            (row["key"], row["value"]),
        )
    for row in data.get("day_overrides", []):
        conn.execute(
            "INSERT INTO day_overrides (day, capacity, closed) VALUES (?, ?, ?)"
            " ON CONFLICT (day) DO UPDATE SET capacity = excluded.capacity, closed = excluded.closed",
            (row["day"], row["capacity"], row["closed"]),
        )
    conn.execute("COMMIT")
    total = conn.execute("SELECT COUNT(*) FROM registrations").fetchone()[0]
    print(f"added={added} skipped={skipped} total={total}")
    conn.close()


if __name__ == "__main__":
    restore(sys.argv[1])
