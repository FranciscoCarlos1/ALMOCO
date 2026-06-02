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
    """,
    3: """
        CREATE TABLE IF NOT EXISTS cardapios (
            data_almoco TEXT PRIMARY KEY,
            descricao TEXT NOT NULL,
            atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """,
    4: """
        ALTER TABLE cardapios ADD COLUMN imagem_path TEXT;
    """,
    5: """
        ALTER TABLE cardapios ADD COLUMN imagem_blob BYTEA;
    """,
    6: """
        ALTER TABLE cardapios ADD COLUMN imagem_mime TEXT;
    """,
    7: """
        CREATE TABLE IF NOT EXISTS alunos (
            matricula TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            turma TEXT NOT NULL,
            atualizado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
    """,
    8: """
        CREATE TABLE IF NOT EXISTS respostas (
            id BIGSERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            matricula TEXT NOT NULL,
            turma TEXT NOT NULL,
            data_almoco DATE NOT NULL,
            intencao TEXT NOT NULL CHECK (intencao IN ('SIM', 'NAO')),
            criado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(matricula, data_almoco)
        );
    """,
    9: """
        CREATE TABLE IF NOT EXISTS quadro_importado (
            turma TEXT NOT NULL,
            data_almoco DATE NOT NULL,
            sim INTEGER NOT NULL DEFAULT 0,
            atualizado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (turma, data_almoco)
        );
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