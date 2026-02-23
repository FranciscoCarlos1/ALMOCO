import sqlite3
import os
from psycopg import connect

SQLITE_DB = "data/almoco.db"
POSTGRES_URL = os.getenv("python app.py DATABASE_URL", "postgresql://neondb_owner:npg_8pec2mvaJqfM@ep-royal-hill-aigd1y0z-pooler.c-4.us-east-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require:5432/almoco_db")

def migrate():
    sqlite_conn = sqlite3.connect(SQLITE_DB)
    sqlite_conn.row_factory = sqlite3.Row

    pg_conn = connect(POSTGRES_URL)
    pg_cursor = pg_conn.cursor()

    sqlite_cursor = sqlite_conn.cursor()
    sqlite_cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")

    tables = [row["name"] for row in sqlite_cursor.fetchall()]

    for table in tables:
        if table == "sqlite_sequence":
            continue

        print(f"Migrando tabela {table}")

        rows = sqlite_conn.execute(f"SELECT * FROM {table}").fetchall()

        for row in rows:
            columns = row.keys()
            values = tuple(row)

            placeholders = ", ".join(["%s"] * len(values))
            columns_str = ", ".join(columns)

            query = f"INSERT INTO {table} ({columns_str}) VALUES ({placeholders})"
            pg_cursor.execute(query, values)

    pg_conn.commit()
    pg_cursor.close()
    pg_conn.close()
    sqlite_conn.close()

if __name__ == "__main__":
    migrate()