#!/usr/bin/env python3
"""
GeometryHome WFMS - SQLite to PostgreSQL Migration Script
"""
import sqlite3
import argparse
import psycopg2


def migrate(sqlite_path: str, pg_url: str):
    print("\n" + "=" * 55)
    print("  GeometryHome WFMS — Database Migration")
    print(f"  Source : {sqlite_path}")
    print(f"  Target : {pg_url[:40]}...")
    print("=" * 55 + "\n")

    # ── Connect ─────────────────────────────
    sq = sqlite3.connect(sqlite_path)
    sq.row_factory = sqlite3.Row

    pg = psycopg2.connect(pg_url)
    pg.autocommit = False
    cur = pg.cursor()

    # ── Create tables ───────────────────────
    print("▶ Creating tables (if not exist)...")

    stmts = [
        """CREATE TABLE IF NOT EXISTS users(
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'viewer',
            full_name TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS locations(
            id SERIAL PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            description TEXT,
            loc_type TEXT DEFAULT 'accommodation',
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS food_preferences(
            id SERIAL PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            category TEXT,
            is_active INTEGER DEFAULT 1
        )""",

        """CREATE TABLE IF NOT EXISTS employees(
            id SERIAL PRIMARY KEY,
            emp_id TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            department TEXT,
            accommodation_id INTEGER,
            shift_type TEXT DEFAULT 'normal',
            food_pref_id INTEGER,
            no_food_sunday INTEGER DEFAULT 0,
            remarks TEXT,
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS attendance(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER,
            att_date TEXT NOT NULL,
            status TEXT DEFAULT 'absent',
            reason TEXT,
            marked_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(employee_id, att_date)
        )""",

        """CREATE TABLE IF NOT EXISTS suspensions(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER,
            start_date TEXT NOT NULL,
            end_date TEXT,
            reason TEXT,
            is_active INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS fasting_records(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER,
            start_date TEXT NOT NULL,
            end_date TEXT,
            reason TEXT,
            is_active INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS vacation_records(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            reason TEXT,
            is_active INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS meal_exceptions(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER,
            meal_type TEXT NOT NULL,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            reason TEXT,
            is_active INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS temp_meal_overrides(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER,
            override_shift_type TEXT,
            override_accommodation_id INTEGER,
            orig_shift_type TEXT,
            orig_accommodation_id INTEGER,
            orig_acc_name TEXT,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            reason TEXT,
            is_active INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS report_history(
            id SERIAL PRIMARY KEY,
            report_date TEXT NOT NULL UNIQUE,
            report_data TEXT NOT NULL,
            generated_by TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS meal_prices(
            food_pref TEXT NOT NULL,
            meal_type TEXT NOT NULL,
            price REAL DEFAULT 0,
            PRIMARY KEY(food_pref, meal_type)
        )""",

        """CREATE TABLE IF NOT EXISTS holidays(
            id SERIAL PRIMARY KEY,
            date TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""",

        """CREATE TABLE IF NOT EXISTS settings(
            key TEXT PRIMARY KEY,
            value TEXT
        )""",
    ]

    for s in stmts:
        cur.execute(s)

    pg.commit()
    print("  ✅ Tables ready\n")

    # ── Migration order ─────────────────────
    order = [
        'settings',
        'locations',
        'food_preferences',
        'users',
        'employees',
        'attendance',
        'suspensions',
        'fasting_records',
        'vacation_records',
        'meal_exceptions',
        'temp_meal_overrides',
        'report_history',
        'meal_prices',
        'holidays',
    ]

    total = 0

    for table in order:
        rows = sq.execute(f"SELECT * FROM {table}").fetchall()

        if not rows:
            print(f"  ⬜ {table}: empty")
            continue

        cols = list(rows[0].keys())

        placeholders = ",".join(["%s"] * len(cols))
        col_str = ",".join(cols)

        sql = f"INSERT INTO {table} ({col_str}) VALUES ({placeholders}) ON CONFLICT DO NOTHING"

        count = 0
        for r in rows:
            try:
                cur.execute(sql, [r[c] for c in cols])
                count += 1
            except Exception:
                pg.rollback()

        pg.commit()

        print(f"  ✅ {table}: {count}/{len(rows)} rows migrated")
        total += count

    sq.close()
    cur.close()
    pg.close()

    print("\n" + "=" * 55)
    print(f" Migration complete: {total} rows")
    print("=" * 55 + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--sqlite", required=True)
    parser.add_argument("--pg", required=True)
    args = parser.parse_args()

    migrate(args.sqlite, args.pg)