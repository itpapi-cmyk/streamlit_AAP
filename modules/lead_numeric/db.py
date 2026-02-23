import sqlite3
from pathlib import Path

DB_PATH = Path("data/audit.db")

def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(DB_PATH)
