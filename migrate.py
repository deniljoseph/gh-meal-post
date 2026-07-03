#!/usr/bin/env python3
"""
GeometryHome WFMS - SQLite to PostgreSQL Migration Script
Usage:
  python migrate.py --sqlite workforce.db --pg "postgresql://user:pass@host:5432/dbname"
"""
import sqlite3, sys, argparse, json
from datetime import datetime

def migrate(sqlite_path: str, pg_url: str):
    print(f"\n{'='*55}")
    print(f"  GeometryHome WFMS — Database Migration")
    print(f"  Source : {sqlite_path}")
    print(f"  Target : {pg_url[:40]}...")
    print(f"{'='*55}\n")

    # ── Connect ────────────────────────────────────────────────
    sq = sqlite3.connect(sqlite_path)
    sq.row_factory = sqlite3.Row

    import psycopg2
    pg = psycopg2.connect(pg_url)
    pg.autocommit = False
    cur = pg.cursor()

    # ── Create all tables ──────────────────────────────────────
    print("▶ Creating tables (if not exist)...")
    cur.executescript = lambda s: [cur.execute(stmt.strip()) for stmt in s.split(';') if stmt.strip()]
    stmts = [
        """CREATE TABLE IF NOT EXISTS users(id SERIAL PRIMARY KEY,username TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,role TEXT NOT NULL DEFAULT 'viewer',full_name TEXT,is_active INTEGER DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS locations(id SERIAL PRIMARY KEY,name TEXT UNIQUE NOT NULL,description TEXT,loc_type TEXT DEFAULT 'accommodation',is_active INTEGER DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS food_preferences(id SERIAL PRIMARY KEY,name TEXT UNIQUE NOT NULL,category TEXT,is_active INTEGER DEFAULT 1)""",
        """CREATE TABLE IF NOT EXISTS employees(id SERIAL PRIMARY KEY,emp_id TEXT UNIQUE NOT NULL,full_name TEXT NOT NULL,department TEXT,accommodation_id INTEGER,shift_type TEXT DEFAULT 'normal',food_pref_id INTEGER,no_food_sunday INTEGER DEFAULT 0,remarks TEXT,status TEXT DEFAULT 'active',created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS attendance(id SERIAL PRIMARY KEY,employee_id INTEGER,att_date TEXT NOT NULL,status TEXT DEFAULT 'absent',reason TEXT,marked_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP,UNIQUE(employee_id,att_date))""",
        """CREATE TABLE IF NOT EXISTS suspensions(id SERIAL PRIMARY KEY,employee_id INTEGER,start_date TEXT NOT NULL,end_date TEXT,reason TEXT,is_active INTEGER DEFAULT 1,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS fasting_records(id SERIAL PRIMARY KEY,employee_id INTEGER,start_date TEXT NOT NULL,end_date TEXT,reason TEXT,is_active INTEGER DEFAULT 1,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS vacation_records(id SERIAL PRIMARY KEY,employee_id INTEGER,start_date TEXT NOT NULL,end_date TEXT NOT NULL,reason TEXT,is_active INTEGER DEFAULT 1,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS meal_exceptions(id SERIAL PRIMARY KEY,employee_id INTEGER,meal_type TEXT NOT NULL,start_date TEXT NOT NULL,end_date TEXT NOT NULL,reason TEXT,is_active INTEGER DEFAULT 1,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS temp_meal_overrides(id SERIAL PRIMARY KEY,employee_id INTEGER,override_shift_type TEXT,override_accommodation_id INTEGER,orig_shift_type TEXT,orig_accommodation_id INTEGER,orig_acc_name TEXT,start_date TEXT NOT NULL,end_date TEXT NOT NULL,reason TEXT,is_active INTEGER DEFAULT 1,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS report_history(id SERIAL PRIMARY KEY,report_date TEXT NOT NULL UNIQUE,report_data TEXT NOT NULL,generated_by TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS meal_prices(food_pref TEXT NOT NULL,meal_type TEXT NOT NULL,price REAL DEFAULT 0,PRIMARY KEY(food_pref,meal_type))""",
        """CREATE TABLE IF NOT EXISTS holidays(id SERIAL PRIMARY KEY,date TEXT NOT NULL UNIQUE,name TEXT NOT NULL,created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT)""",
    ]
    for s in stmts:
        cur.execute(s)
    pg.commit()
    print("  ✅ Tables ready\n")

    # ── Migration order (respects FK dependencies) ─────────────
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

    total_migrated = 0

    for table in order:
        rows = sq.execute(f"SELECT * FROM {table}").fetchall()
        if not rows:
            print(f"  ⬜ {table}: empty, skipped")
            continue

        cols = list(rows[0].keys())
        # Skip 'id' for serial tables (PG auto-assigns), keep for settings/meal_prices
        has_serial = table not in ('settings', 'meal_prices')
        insert_cols = cols  # include id so we preserve FKs

        placeholders = ','.join(['%s'] * len(insert_cols))
        col_str = ','.join(insert_cols)

        # Build upsert to avoid duplicate errors on re-run
        if table == 'settings':
            sql = f"INSERT INTO {table}({col_str}) VALUES({placeholders}) ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value"
        elif table == 'meal_prices':
            sql = f"INSERT INTO {table}({col_str}) VALUES({placeholders}) ON CONFLICT(food_pref,meal_type) DO UPDATE SET price=EXCLUDED.price"
        elif table in ('report_history',):
            sql = f"INSERT INTO {table}({col_str}) VALUES({placeholders}) ON CONFLICT(report_date) DO UPDATE SET report_data=EXCLUDED.report_data,generated_by=EXCLUDED.generated_by"
        elif table == 'attendance':
            sql = f"INSERT INTO {table}({col_str}) VALUES({placeholders}) ON CONFLICT(employee_id,att_date) DO UPDATE SET status=EXCLUDED.status"
        else:
            sql = f"INSERT INTO {table}({col_str}) VALUES({placeholders}) ON CONFLICT(id) DO UPDATE SET id=EXCLUDED.id"

        count = 0
        errors = []
        for row in rows:
            vals = [row[c] for c in insert_cols]
            try:
                cur.execute(sql, vals)
                count += 1
            except Exception as e:
                errors.append(str(e)[:80])
                pg.rollback()

        pg.commit()

        # Reset sequence so new inserts don't collide with migrated IDs
        if has_serial and rows:
            max_id = max(row['id'] for row in rows if 'id' in row.keys())
            try:
                cur.execute(f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), %s)", (max_id,))
                pg.commit()
            except Exception as e:
                pg.rollback()

        status = '✅' if not errors else '⚠️'
        print(f"  {status} {table}: {count}/{len(rows)} rows migrated" + (f" | {len(errors)} errors" if errors else ''))
        if errors:
            for e in errors[:3]:
                print(f"      ↳ {e}")
        total_migrated += count

    sq.close()
    cur.close()
    pg.close()

    print(f"\n{'='*55}")
    print(f"  Migration complete! {total_migrated} rows migrated.")
    print(f"  Your app on Railway is ready to use.")
    print(f"{'='*55}\n")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Migrate SQLite → PostgreSQL for WFMS')
    parser.add_argument('--sqlite', required=True, help='Path to SQLite .db file')
    parser.add_argument('--pg', required=True, help='PostgreSQL connection URL')
    args = parser.parse_args()
    migrate(args.sqlite, args.pg)
