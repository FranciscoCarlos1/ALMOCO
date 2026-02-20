from __future__ import annotations

import os
import sqlite3
from pathlib import Path

from psycopg import connect


BASE_DIR = Path(__file__).resolve().parents[1]
SQLITE_PATH = BASE_DIR / "data" / "almoco.db"


def ensure_postgres_schema(conn) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS alunos (
                matricula TEXT PRIMARY KEY,
                nome TEXT NOT NULL,
                turma TEXT NOT NULL,
                atualizado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS respostas (
                id BIGSERIAL PRIMARY KEY,
                nome TEXT NOT NULL,
                matricula TEXT NOT NULL,
                turma TEXT NOT NULL,
                data_almoco DATE NOT NULL,
                intencao TEXT NOT NULL CHECK (intencao IN ('SIM', 'NAO')),
                criado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(matricula, data_almoco)
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS quadro_importado (
                turma TEXT NOT NULL,
                data_almoco DATE NOT NULL,
                sim INTEGER NOT NULL DEFAULT 0,
                atualizado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (turma, data_almoco)
            )
            """
        )
    conn.commit()


def sqlite_table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    ).fetchone()
    return row is not None


def main() -> None:
    database_url = os.getenv("DATABASE_URL", "").strip()
    if not database_url:
        raise SystemExit("Defina DATABASE_URL para o PostgreSQL antes de executar.")

    if not SQLITE_PATH.exists():
        raise SystemExit(f"SQLite não encontrado em: {SQLITE_PATH}")

    sqlite_conn = sqlite3.connect(SQLITE_PATH)
    sqlite_conn.row_factory = sqlite3.Row

    pg_conn = connect(database_url)
    ensure_postgres_schema(pg_conn)

    alunos_rows = sqlite_conn.execute("SELECT matricula, nome, turma FROM alunos").fetchall() if sqlite_table_exists(sqlite_conn, "alunos") else []
    respostas_rows = (
        sqlite_conn.execute("SELECT nome, matricula, turma, data_almoco, intencao FROM respostas").fetchall()
        if sqlite_table_exists(sqlite_conn, "respostas")
        else []
    )
    quadro_rows = (
        sqlite_conn.execute("SELECT turma, data_almoco, sim FROM quadro_importado").fetchall()
        if sqlite_table_exists(sqlite_conn, "quadro_importado")
        else []
    )

    with pg_conn.cursor() as cur:
        for row in alunos_rows:
            cur.execute(
                """
                INSERT INTO alunos (matricula, nome, turma)
                VALUES (%s, %s, %s)
                ON CONFLICT(matricula)
                DO UPDATE SET nome = excluded.nome, turma = excluded.turma, atualizado_em = CURRENT_TIMESTAMP
                """,
                (row["matricula"], row["nome"], row["turma"]),
            )

        for row in respostas_rows:
            cur.execute(
                """
                INSERT INTO respostas (nome, matricula, turma, data_almoco, intencao)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT(matricula, data_almoco)
                DO UPDATE SET nome = excluded.nome, turma = excluded.turma, intencao = excluded.intencao, criado_em = CURRENT_TIMESTAMP
                """,
                (row["nome"], row["matricula"], row["turma"], row["data_almoco"], row["intencao"]),
            )

        for row in quadro_rows:
            cur.execute(
                """
                INSERT INTO quadro_importado (turma, data_almoco, sim)
                VALUES (%s, %s, %s)
                ON CONFLICT(turma, data_almoco)
                DO UPDATE SET sim = excluded.sim, atualizado_em = CURRENT_TIMESTAMP
                """,
                (row["turma"], row["data_almoco"], int(row["sim"] or 0)),
            )

    pg_conn.commit()

    with pg_conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) FROM alunos")
        alunos_pg = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM respostas")
        respostas_pg = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM quadro_importado")
        quadro_pg = cur.fetchone()[0]

    print(f"SQLite -> alunos: {len(alunos_rows)}, respostas: {len(respostas_rows)}, quadro_importado: {len(quadro_rows)}")
    print(f"Postgres -> alunos: {alunos_pg}, respostas: {respostas_pg}, quadro_importado: {quadro_pg}")
    print("MIGRACAO_OK")

    sqlite_conn.close()
    pg_conn.close()


if __name__ == "__main__":
    main()
