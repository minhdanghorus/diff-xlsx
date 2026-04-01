#!/usr/bin/env python3
"""
Database helper for the Diff Tool.

Usage:
    Standalone — test a connection profile:
        python sql_src/db.py
        python -m sql_src.db

    From diff_xlsx.py:
        from sql_src.db import list_profiles, list_sql_files, read_sql_source
"""

import os
import sys
import glob
import psycopg2

# Support both direct execution (python sql_src/db.py) and package import
if __name__ == "__main__" and __package__ is None:
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sql_src.config import CONNECTIONS

SQL_DIR = os.path.dirname(os.path.abspath(__file__))


# ─── Helpers ─────────────────────────────────────────────────────────────────

def list_profiles():
    """Return the list of profile names from config."""
    return list(CONNECTIONS.keys())


def list_sql_files():
    """Return a sorted list of .sql file paths found in sql_src/."""
    return sorted(glob.glob(os.path.join(SQL_DIR, "*.sql")))


def get_connection(profile_name):
    """Open and return a psycopg2 connection for the given profile."""
    if profile_name not in CONNECTIONS:
        raise ValueError(f"Unknown profile: '{profile_name}'")
    params = CONNECTIONS[profile_name]
    return psycopg2.connect(
        host=params["host"],
        port=params.get("port", 5432),
        database=params["database"],
        user=params["user"],
        password=params["password"],
    )


def test_connection(profile_name):
    """Test connectivity for a profile. Prints success or error."""
    print(f"Testing connection '{profile_name}' ...")
    try:
        conn = get_connection(profile_name)
        cur = conn.cursor()
        cur.execute("SELECT version();")
        version = cur.fetchone()[0]
        cur.close()
        conn.close()
        print(f"  OK — {version}")
        return True
    except Exception as e:
        print(f"  FAILED — {e}")
        return False


def execute_query(profile_name, sql):
    """
    Run a SELECT query and return (headers, rows) in the same format
    as read_sheet() in diff_xlsx.py.

    headers: list of column-name strings
    rows:    list of lists (one per data row)
    """
    conn = get_connection(profile_name)
    try:
        cur = conn.cursor()
        cur.execute(sql)
        headers = [desc[0] for desc in cur.description]
        rows = [list(row) for row in cur.fetchall()]
        cur.close()
        return headers, rows
    finally:
        conn.close()


# ─── Interactive prompts (used by diff_xlsx.py) ─────────────────────────────

def ask_profile():
    """Prompt the user to pick a connection profile. Returns the profile name."""
    profiles = list_profiles()
    if not profiles:
        raise RuntimeError("No connection profiles defined in sql_src/config.py")

    print("\nAvailable connection profiles:")
    for i, name in enumerate(profiles, 1):
        cfg = CONNECTIONS[name]
        print(f"  {i}. {name}  ({cfg['host']}:{cfg.get('port', 5432)}/{cfg['database']})")

    while True:
        try:
            choice = int(input("Select profile number: ").strip())
            if 1 <= choice <= len(profiles):
                return profiles[choice - 1]
            print(f"Please enter a number between 1 and {len(profiles)}.")
        except ValueError:
            print("Please enter a valid number.")


def ask_sql_query():
    """
    Prompt the user to pick a .sql file from sql_src/ or enter a query manually.
    Returns the SQL string.
    """
    sql_files = list_sql_files()

    options = []
    if sql_files:
        print("\nAvailable .sql files:")
        for i, f in enumerate(sql_files, 1):
            options.append(f)
            print(f"  {i}. {os.path.basename(f)}")
        print(f"  {len(sql_files) + 1}. Enter query manually")
    else:
        print("\n  No .sql files found in sql_src/.")

    if sql_files:
        while True:
            try:
                choice = int(input("Select option: ").strip())
                if 1 <= choice <= len(sql_files):
                    path = sql_files[choice - 1]
                    with open(path, encoding="utf-8") as fh:
                        sql = fh.read().strip()
                    print(f"  Loaded: {os.path.basename(path)}")
                    return sql
                elif choice == len(sql_files) + 1:
                    break  # fall through to manual entry
                else:
                    print(f"Please enter a number between 1 and {len(sql_files) + 1}.")
            except ValueError:
                print("Please enter a valid number.")

    # Manual entry
    print("Enter your SQL query (end with a semicolon on its own line or an empty line):")
    lines = []
    while True:
        line = input()
        if line.strip() == "" or line.strip() == ";":
            break
        lines.append(line)
    sql = "\n".join(lines).strip().rstrip(";")
    if not sql:
        raise ValueError("Empty query.")
    return sql


def read_sql_source():
    """
    Full interactive flow for one SQL source:
      1. Pick profile
      2. Pick or enter query
      3. Execute and return (headers, rows)

    Returns (source_name, headers, rows) where source_name is
    "<profile>: <query_file_or_manual>" for display purposes.
    """
    profile = ask_profile()
    sql = ask_sql_query()

    print(f"  Executing query on '{profile}' ...")
    headers, rows = execute_query(profile, sql)
    print(f"  Got {len(rows)} rows, {len(headers)} columns.")

    source_name = f"SQL:{profile}"
    return source_name, headers, rows


# ─── Standalone entry point ─────────────────────────────────────────────────

def _standalone():
    """Interactive mode: test one or all connection profiles."""
    profiles = list_profiles()
    if not profiles:
        print("No connection profiles defined in sql_src/config.py")
        return

    print("=== SQL Connection Tester ===\n")
    print("Profiles:")
    for i, name in enumerate(profiles, 1):
        cfg = CONNECTIONS[name]
        print(f"  {i}. {name}  ({cfg['host']}:{cfg.get('port', 5432)}/{cfg['database']})")
    print(f"  {len(profiles) + 1}. Test ALL")

    while True:
        try:
            choice = int(input("\nSelect profile to test: ").strip())
            if 1 <= choice <= len(profiles):
                test_connection(profiles[choice - 1])
                break
            elif choice == len(profiles) + 1:
                print()
                for p in profiles:
                    test_connection(p)
                break
            else:
                print(f"Please enter a number between 1 and {len(profiles) + 1}.")
        except ValueError:
            print("Please enter a valid number.")


if __name__ == "__main__":
    _standalone()
