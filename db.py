import os
import sqlite3
from pathlib import Path
from typing import Any

try:
    from psycopg import connect as pg_connect
    from psycopg.rows import dict_row
except Exception:
    pg_connect = None
    dict_row = None

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR_ENV = os.getenv("ALMOCO_DATA_DIR")
DB_DIR = Path(DATA_DIR_ENV) if DATA_DIR_ENV else BASE_DIR / "data"
DB_PATH = DB_DIR / "almoco.db"
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[len("postgres://"):]
USE_POSTGRES = DATABASE_URL.startswith("postgresql://")

class DBConnection:
    def __init__(self, raw: Any, is_postgres: bool):
        self.raw = raw
        self.is_postgres = is_postgres

    def execute(self, query: str, params: tuple[Any, ...] = ()):  # noqa: ANN001
        if self.is_postgres:
            query = query.replace("?", "%s")
        return self.raw.execute(query, params)

    def commit(self) -> None:
        self.raw.commit()

    def rollback(self) -> None:
        self.raw.rollback()

    def close(self) -> None:
        self.raw.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type is not None:
            try:
                self.rollback()
            except Exception:
                pass
        self.close()
        return False

def get_conn() -> DBConnection:
    if USE_POSTGRES:
        print(f"[ALMOCO] Usando Postgres: {DATABASE_URL}")
        if pg_connect is None or dict_row is None:
            raise RuntimeError("psycopg não instalado. Adicione 'psycopg[binary]' no requirements.")
        conn = pg_connect(DATABASE_URL, row_factory=dict_row)
        return DBConnection(conn, is_postgres=True)

    print(f"[ALMOCO] Usando SQLite: {DB_PATH}")
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return DBConnection(conn, is_postgres=False)
