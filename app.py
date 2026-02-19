from __future__ import annotations

import csv
import os
import sqlite3
from datetime import date, datetime, timedelta
from io import StringIO
from pathlib import Path

from flask import Flask, Response, abort, jsonify, redirect, render_template, request, url_for

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR_ENV = os.getenv("ALMOCO_DATA_DIR")
DB_DIR = Path(DATA_DIR_ENV) if DATA_DIR_ENV else BASE_DIR / "data"
DB_PATH = DB_DIR / "almoco.db"
ADMIN_TOKEN = os.getenv("ALMOCO_ADMIN_TOKEN", "ifc-sbs")

TURMAS = [
    "TIN I",
    "TIN II",
    "TIN III",
    "TAI I",
    "TAI II",
    "TAI III",
    "TST I",
    "TST II",
    "TST III",
]

INTENCOES = ["SIM", "NAO"]
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex"]

app = Flask(__name__)


def get_conn() -> sqlite3.Connection:
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS alunos (
                matricula TEXT PRIMARY KEY,
                nome TEXT NOT NULL,
                turma TEXT NOT NULL,
                atualizado_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS respostas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                matricula TEXT NOT NULL,
                turma TEXT NOT NULL,
                data_almoco DATE NOT NULL,
                intencao TEXT NOT NULL CHECK (intencao IN ('SIM', 'NAO')),
                criado_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(matricula, data_almoco)
            )
            """
        )
        conn.commit()


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def week_start(given_date: date) -> date:
    return given_date - timedelta(days=given_date.weekday())


def normalize_header(header: str) -> str:
    return (
        header.strip()
        .lower()
        .replace("á", "a")
        .replace("à", "a")
        .replace("â", "a")
        .replace("ã", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
    )


def get_csv_value(row: dict[str, str], candidates: list[str]) -> str:
    for key in candidates:
        if key in row and row[key] is not None:
            return str(row[key]).strip()
    return ""


def is_admin_allowed() -> bool:
    token = request.args.get("token", "")
    return token == ADMIN_TOKEN


def is_admin_allowed_form() -> bool:
    token = request.form.get("token", "")
    return token == ADMIN_TOKEN


@app.get("/")
def index() -> str:
    sucesso = request.args.get("sucesso") == "1"
    erro = request.args.get("erro")
    hoje = date.today().isoformat()
    return render_template(
        "index.html",
        turmas=TURMAS,
        intencoes=INTENCOES,
        sucesso=sucesso,
        erro=erro,
        hoje=hoje,
    )


@app.get("/aluno")
def buscar_aluno():
    matricula = request.args.get("matricula", "").strip()
    if not matricula:
        return jsonify({"ok": False, "erro": "Matrícula não informada."}), 400

    with get_conn() as conn:
        aluno = conn.execute(
            """
            SELECT nome, matricula, turma
            FROM alunos
            WHERE matricula = ?
            """,
            (matricula,),
        ).fetchone()

    if not aluno:
        return jsonify({"ok": False, "erro": "Matrícula não encontrada."}), 404

    return jsonify(
        {
            "ok": True,
            "nome": aluno["nome"],
            "matricula": aluno["matricula"],
            "turma": aluno["turma"],
        }
    )


@app.post("/enviar")
def enviar():
    nome = request.form.get("nome", "").strip()
    matricula = request.form.get("matricula", "").strip()
    turma = request.form.get("turma", "").strip()
    data_referencia = request.form.get("data_almoco", "").strip()
    dias_raw = request.form.getlist("dias")
    dias_marcados: list[str] = []
    for raw in dias_raw:
        normalizado = raw.replace(";", ",").replace(" ", ",")
        partes = [item.strip().lower() for item in normalizado.split(",") if item.strip()]
        dias_marcados.extend(partes)
    dias_marcados = list(dict.fromkeys(dias_marcados))

    if not nome:
        return redirect(url_for("index", erro="Informe seu nome."))
    if turma not in TURMAS:
        return redirect(url_for("index", erro="Selecione uma turma válida."))

    if not matricula:
        matricula = f"AUTO::{turma}::{nome}".upper()

    if not dias_marcados:
        return redirect(url_for("index", erro="Marque pelo menos um dia da semana."))
    if any(item not in DIAS_SEMANA for item in dias_marcados):
        return redirect(url_for("index", erro="Seleção de dias inválida."))

    try:
        data_ref = parse_iso_date(data_referencia)
    except ValueError:
        return redirect(url_for("index", erro="Informe uma data válida."))

    segunda = week_start(data_ref)
    datas_semana = {
        "seg": segunda,
        "ter": segunda + timedelta(days=1),
        "qua": segunda + timedelta(days=2),
        "qui": segunda + timedelta(days=3),
        "sex": segunda + timedelta(days=4),
    }

    with get_conn() as conn:
        for dia, data_almoco in datas_semana.items():
            intencao = "SIM" if dia in dias_marcados else "NAO"
            conn.execute(
                """
                INSERT INTO respostas (nome, matricula, turma, data_almoco, intencao)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(matricula, data_almoco)
                DO UPDATE SET
                    nome = excluded.nome,
                    turma = excluded.turma,
                    intencao = excluded.intencao,
                    criado_em = CURRENT_TIMESTAMP
                """,
                (nome, matricula, turma, data_almoco.isoformat(), intencao),
            )
        conn.commit()

    return redirect(url_for("index", sucesso=1))


@app.get("/admin")
def admin() -> str:
    if not is_admin_allowed():
        abort(403, "Acesso negado. Informe um token válido na URL.")

    data_filtro = request.args.get("data") or date.today().isoformat()
    try:
        data_base = parse_iso_date(data_filtro)
    except ValueError:
        data_base = date.today()
        data_filtro = data_base.isoformat()

    segunda = week_start(data_base)
    sexta = segunda + timedelta(days=4)

    with get_conn() as conn:
        resumo_rows = conn.execute(
            """
            SELECT turma,
                   SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END) AS sim,
                   SUM(CASE WHEN intencao = 'NAO' THEN 1 ELSE 0 END) AS nao,
                 SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END) AS total
            FROM respostas
            WHERE data_almoco = ?
            GROUP BY turma
            ORDER BY turma
            """,
            (data_filtro,),
        ).fetchall()

        turma_semana_rows = conn.execute(
            """
            SELECT turma,
                   data_almoco,
                   SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END) AS sim
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            GROUP BY turma, data_almoco
            ORDER BY turma, data_almoco
            """,
            (segunda.isoformat(), sexta.isoformat()),
        ).fetchall()

        semana_rows = conn.execute(
            """
            SELECT data_almoco,
                   SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END) AS sim
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            GROUP BY data_almoco
            ORDER BY data_almoco
            """,
            (segunda.isoformat(), sexta.isoformat()),
        ).fetchall()

        respostas = conn.execute(
            """
            SELECT nome, matricula, turma, intencao, criado_em
            FROM respostas
            WHERE data_almoco = ?
            ORDER BY turma, nome
            """,
            (data_filtro,),
        ).fetchall()

    resumo = {turma: {"sim": 0, "nao": 0, "total": 0} for turma in TURMAS}
    for row in resumo_rows:
        resumo[row["turma"]] = {
            "sim": row["sim"] or 0,
            "nao": row["nao"] or 0,
            "total": row["total"] or 0,
        }

    total_sim = sum(item["sim"] for item in resumo.values())
    total_nao = sum(item["nao"] for item in resumo.values())
    total_geral = total_sim

    semana_sim = {"seg": 0, "ter": 0, "qua": 0, "qui": 0, "sex": 0}
    week_map = {
        segunda.isoformat(): "seg",
        (segunda + timedelta(days=1)).isoformat(): "ter",
        (segunda + timedelta(days=2)).isoformat(): "qua",
        (segunda + timedelta(days=3)).isoformat(): "qui",
        (segunda + timedelta(days=4)).isoformat(): "sex",
    }

    turma_semana = {turma: {"seg": 0, "ter": 0, "qua": 0, "qui": 0, "sex": 0, "total": 0} for turma in TURMAS}
    for row in turma_semana_rows:
        turma = row["turma"]
        dia = week_map.get(row["data_almoco"])
        if turma not in turma_semana or not dia:
            continue
        valor = row["sim"] or 0
        turma_semana[turma][dia] = valor
        turma_semana[turma]["total"] += valor

    for row in semana_rows:
        key = week_map.get(row["data_almoco"])
        if key:
            semana_sim[key] = row["sim"] or 0

    total_semana_geral = sum(semana_sim.values())

    return render_template(
        "admin.html",
        resumo=resumo,
        data_filtro=data_filtro,
        total_sim=total_sim,
        total_nao=total_nao,
        total_geral=total_geral,
        respostas=respostas,
        token=request.args.get("token", ""),
        importado=request.args.get("importado") == "1",
        import_error=request.args.get("import_error"),
        semana_sim=semana_sim,
        turma_semana=turma_semana,
        total_semana_geral=total_semana_geral,
        semana_inicio=segunda.isoformat(),
        semana_fim=sexta.isoformat(),
    )


@app.post("/admin/importar_alunos")
def importar_alunos():
    if not is_admin_allowed_form():
        abort(403, "Acesso negado. Informe um token válido.")

    file = request.files.get("arquivo_csv")
    token = request.form.get("token", "")
    data_filtro = request.form.get("data", "")

    if not file or not file.filename:
        return redirect(url_for("admin", token=token, data=data_filtro, import_error="Selecione um arquivo CSV."))

    try:
        payload = file.stream.read().decode("utf-8-sig")
        sample = payload[:2048]
        dialect = csv.Sniffer().sniff(sample, delimiters=",;")
        reader = csv.DictReader(StringIO(payload), dialect=dialect)
    except Exception:
        return redirect(url_for("admin", token=token, data=data_filtro, import_error="Não foi possível ler o CSV."))

    if not reader.fieldnames:
        return redirect(url_for("admin", token=token, data=data_filtro, import_error="CSV sem cabeçalho."))

    normalized = [normalize_header(item) for item in reader.fieldnames]
    header_map = {normalized[i]: reader.fieldnames[i] for i in range(len(normalized))}

    nome_col = next((h for h in ["nome", "aluno", "nome completo"] if h in header_map), None)
    matricula_col = next((h for h in ["matricula", "matricula aluno", "ra"] if h in header_map), None)
    turma_col = next((h for h in ["turma", "serie", "classe"] if h in header_map), None)

    if not nome_col or not matricula_col or not turma_col:
        return redirect(
            url_for(
                "admin",
                token=token,
                data=data_filtro,
                import_error="CSV precisa das colunas: nome, matricula e turma.",
            )
        )

    importados = 0
    with get_conn() as conn:
        for row in reader:
            nome = get_csv_value(row, [header_map[nome_col]])
            matricula = get_csv_value(row, [header_map[matricula_col]])
            turma = get_csv_value(row, [header_map[turma_col]])

            if not nome or not matricula or turma not in TURMAS:
                continue

            conn.execute(
                """
                INSERT INTO alunos (matricula, nome, turma)
                VALUES (?, ?, ?)
                ON CONFLICT(matricula)
                DO UPDATE SET
                    nome = excluded.nome,
                    turma = excluded.turma,
                    atualizado_em = CURRENT_TIMESTAMP
                """,
                (matricula, nome, turma),
            )
            importados += 1
        conn.commit()

    if importados == 0:
        return redirect(
            url_for(
                "admin",
                token=token,
                data=data_filtro,
                import_error="Nenhum aluno válido importado (verifique turma e colunas).",
            )
        )

    return redirect(url_for("admin", token=token, data=data_filtro, importado=1))


@app.get("/admin/planilha")
def planilha_semana() -> str:
    if not is_admin_allowed():
        abort(403, "Acesso negado. Informe um token válido na URL.")

    token = request.args.get("token", "")
    turma = request.args.get("turma") or TURMAS[0]
    semana_ref = request.args.get("semana") or date.today().isoformat()

    try:
        monday = week_start(parse_iso_date(semana_ref))
    except ValueError:
        monday = week_start(date.today())

    week_dates = [monday + timedelta(days=i) for i in range(5)]
    week_map = {week_dates[i].isoformat(): DIAS_SEMANA[i] for i in range(5)}

    with get_conn() as conn:
        alunos = conn.execute(
            """
            SELECT nome, matricula, turma
            FROM alunos
            WHERE turma = ?
            ORDER BY nome
            """,
            (turma,),
        ).fetchall()

        respostas = conn.execute(
            """
            SELECT matricula, data_almoco, intencao
            FROM respostas
            WHERE turma = ?
              AND data_almoco BETWEEN ? AND ?
            """,
            (turma, week_dates[0].isoformat(), week_dates[-1].isoformat()),
        ).fetchall()

    marks: dict[str, dict[str, str]] = {}
    for row in respostas:
        data_almoco = row["data_almoco"]
        if data_almoco not in week_map:
            continue
        dia = week_map[data_almoco]
        marks.setdefault(row["matricula"], {})
        marks[row["matricula"]][dia] = "X" if row["intencao"] == "SIM" else "-"

    linhas = []
    for idx, aluno in enumerate(alunos, start=1):
        aluno_marks = marks.get(aluno["matricula"], {})
        linhas.append(
            {
                "ordem": idx,
                "nome": aluno["nome"],
                "seg": aluno_marks.get("seg", ""),
                "ter": aluno_marks.get("ter", ""),
                "qua": aluno_marks.get("qua", ""),
                "qui": aluno_marks.get("qui", ""),
                "sex": aluno_marks.get("sex", ""),
            }
        )

    return render_template(
        "planilha.html",
        token=token,
        turma=turma,
        turmas=TURMAS,
        semana_ref=semana_ref,
        monday=week_dates[0].isoformat(),
        friday=week_dates[-1].isoformat(),
        linhas=linhas,
    )


@app.get("/export_semana.csv")
def export_semana_csv() -> Response:
    if not is_admin_allowed():
        abort(403, "Acesso negado. Informe um token válido na URL.")

    turma = request.args.get("turma") or TURMAS[0]
    semana_ref = request.args.get("semana") or date.today().isoformat()

    try:
        monday = week_start(parse_iso_date(semana_ref))
    except ValueError:
        monday = week_start(date.today())

    week_dates = [monday + timedelta(days=i) for i in range(5)]
    week_map = {week_dates[i].isoformat(): DIAS_SEMANA[i] for i in range(5)}

    with get_conn() as conn:
        alunos = conn.execute(
            """
            SELECT nome, matricula
            FROM alunos
            WHERE turma = ?
            ORDER BY nome
            """,
            (turma,),
        ).fetchall()

        respostas = conn.execute(
            """
            SELECT matricula, data_almoco, intencao
            FROM respostas
            WHERE turma = ?
              AND data_almoco BETWEEN ? AND ?
            """,
            (turma, week_dates[0].isoformat(), week_dates[-1].isoformat()),
        ).fetchall()

    marks: dict[str, dict[str, str]] = {}
    for row in respostas:
        data_almoco = row["data_almoco"]
        if data_almoco not in week_map:
            continue
        dia = week_map[data_almoco]
        marks.setdefault(row["matricula"], {})
        marks[row["matricula"]][dia] = "X" if row["intencao"] == "SIM" else "-"

    output = StringIO()
    writer = csv.writer(output, delimiter=';')
    writer.writerow(["N", "Nome", "Seg", "Ter", "Qua", "Qui", "Sex"])

    for idx, aluno in enumerate(alunos, start=1):
        aluno_marks = marks.get(aluno["matricula"], {})
        writer.writerow(
            [
                idx,
                aluno["nome"],
                aluno_marks.get("seg", ""),
                aluno_marks.get("ter", ""),
                aluno_marks.get("qua", ""),
                aluno_marks.get("qui", ""),
                aluno_marks.get("sex", ""),
            ]
        )

    csv_data = output.getvalue()
    output.close()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={
            "Content-Disposition": f"attachment; filename=planilha_{turma}_{week_dates[0].isoformat()}.csv"
        },
    )


@app.get("/export.csv")
def export_csv() -> Response:
    if not is_admin_allowed():
        abort(403, "Acesso negado. Informe um token válido na URL.")

    data_filtro = request.args.get("data") or date.today().isoformat()
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT nome, matricula, turma, data_almoco, intencao, criado_em
            FROM respostas
            WHERE data_almoco = ?
            ORDER BY turma, nome
            """,
            (data_filtro,),
        ).fetchall()

    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["nome", "matricula", "turma", "data_almoco", "intencao", "criado_em"])
    for row in rows:
        writer.writerow([row["nome"], row["matricula"], row["turma"], row["data_almoco"], row["intencao"], row["criado_em"]])

    csv_data = output.getvalue()
    output.close()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=almoco_{data_filtro}.csv"},
    )


init_db()

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "0") == "1",
    )
