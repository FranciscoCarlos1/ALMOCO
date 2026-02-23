from db import get_conn

MIGRATIONS = {
    1: """
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """,
    2: """
        ALTER TABLE usuarios ADD COLUMN ativo BOOLEAN DEFAULT TRUE;
    """
}

def run_migrations():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS schema_version (
                version INTEGER PRIMARY KEY
            );
        """)
        result = conn.raw.execute("SELECT MAX(version) as v FROM schema_version")
        row = result.fetchone()
        current_version = row["v"] if row and row["v"] else 0

        for version in sorted(MIGRATIONS.keys()):
            if version > current_version:
                print(f"Rodando migration {version}")
                conn.execute(MIGRATIONS[version])
                conn.execute(
                    "INSERT INTO schema_version (version) VALUES (?)"
                    if not conn.is_postgres else
                    "INSERT INTO schema_version (version) VALUES (%s)",
                    (version,)
                )
                conn.commit()