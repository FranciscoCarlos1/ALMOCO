from __future__ import annotations

import csv
import os
import re
import sqlite3
from datetime import date, datetime, timedelta
from io import BytesIO, StringIO
from pathlib import Path

from openpyxl import Workbook, load_workbook
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
    "SERVIDORES",
]

TURMAS_LABEL = {
    "TAI I": "TÉCNICO EM AUTOMAÇÃO INDUSTRIAL – 1",
    "TAI II": "TÉCNICO EM AUTOMAÇÃO INDUSTRIAL – 2",
    "TAI III": "TÉCNICO EM AUTOMAÇÃO INDUSTRIAL – 3",
    "TIN I": "TÉCNICO EM INFORMÁTICA – 1",
    "TIN II": "TÉCNICO EM INFORMÁTICA – 2",
    "TIN III": "TÉCNICO EM INFORMÁTICA – 3",
    "TST I": "TÉCNICO EM SEGURANÇA DO TRABALHO – 1",
    "TST II": "TÉCNICO EM SEGURANÇA DO TRABALHO – 2",
    "TST III": "TÉCNICO EM SEGURANÇA DO TRABALHO – 3",
    "SERVIDORES": "SERVIDORES",
}

TURMAS_ORDEM_QUADRO = ["TAI I", "TAI II", "TAI III", "TIN I", "TIN II", "TIN III", "TST I", "TST II", "TST III", "SERVIDORES"]

INTENCOES = ["SIM", "NAO"]
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex"]

app = Flask(__name__)


def write_backup_xlsx() -> None:
    backup_dir = DB_DIR / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir / f"almoco_backup_{date.today().isoformat()}.xlsx"

    with get_conn() as conn:
        respostas_rows = conn.execute(
            """
            SELECT id, nome, matricula, turma, data_almoco, intencao, criado_em
            FROM respostas
            ORDER BY data_almoco, turma, nome
            """
        ).fetchall()

        alunos_rows = conn.execute(
            """
            SELECT matricula, nome, turma, atualizado_em
            FROM alunos
            ORDER BY turma, nome
            """
        ).fetchall()

        quadro_rows = conn.execute(
            """
            SELECT turma, data_almoco, sim, atualizado_em
            FROM quadro_importado
            ORDER BY data_almoco, turma
            """
        ).fetchall()

    workbook = Workbook()

    ws_respostas = workbook.active
    ws_respostas.title = "respostas"
    ws_respostas.append(["id", "nome", "matricula", "turma", "data_almoco", "intencao", "criado_em"])
    for row in respostas_rows:
        ws_respostas.append([row["id"], row["nome"], row["matricula"], row["turma"], row["data_almoco"], row["intencao"], row["criado_em"]])

    ws_alunos = workbook.create_sheet("alunos")
    ws_alunos.append(["matricula", "nome", "turma", "atualizado_em"])
    for row in alunos_rows:
        ws_alunos.append([row["matricula"], row["nome"], row["turma"], row["atualizado_em"]])

    ws_quadro = workbook.create_sheet("quadro_importado")
    ws_quadro.append(["turma", "data_almoco", "sim", "atualizado_em"])
    for row in quadro_rows:
        ws_quadro.append([row["turma"], row["data_almoco"], row["sim"], row["atualizado_em"]])

    workbook.save(backup_path)
    prune_old_backups(30)


def prune_old_backups(max_backups: int) -> None:
    if max_backups <= 0:
        return

    backup_dir = DB_DIR / "backups"
    if not backup_dir.exists():
        return

    files = sorted(
        backup_dir.glob("almoco_backup_*.xlsx"),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )

    for old_file in files[max_backups:]:
        try:
            old_file.unlink(missing_ok=True)
        except OSError:
            pass


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
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS quadro_importado (
                turma TEXT NOT NULL,
                data_almoco DATE NOT NULL,
                sim INTEGER NOT NULL DEFAULT 0,
                atualizado_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (turma, data_almoco)
            )
            """
        )
        conn.commit()

    write_backup_xlsx()


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def week_start(given_date: date) -> date:
    return given_date - timedelta(days=given_date.weekday())


def month_bounds(given_date: date) -> tuple[date, date]:
    inicio = given_date.replace(day=1)
    if given_date.month == 12:
        proximo = date(given_date.year + 1, 1, 1)
    else:
        proximo = date(given_date.year, given_date.month + 1, 1)
    fim = proximo - timedelta(days=1)
    return inicio, fim


def year_bounds(given_date: date) -> tuple[date, date]:
    return date(given_date.year, 1, 1), date(given_date.year, 12, 31)


def period_bounds(given_date: date, periodo: str) -> tuple[date, date, str]:
    if periodo == "mes":
        inicio, fim = month_bounds(given_date)
        return inicio, fim, "Mês"
    if periodo == "ano":
        inicio, fim = year_bounds(given_date)
        return inicio, fim, "Ano"
    segunda = week_start(given_date)
    sexta = segunda + timedelta(days=4)
    return segunda, sexta, "Semana"


def build_quadro_semana(conn: sqlite3.Connection, segunda: date, sexta: date) -> tuple[dict[str, int], list[dict[str, int | str]], int]:
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

    semana_sim: dict[str, int] = {"seg": 0, "ter": 0, "qua": 0, "qui": 0, "sex": 0}
    week_map = {
        segunda.isoformat(): "seg",
        (segunda + timedelta(days=1)).isoformat(): "ter",
        (segunda + timedelta(days=2)).isoformat(): "qua",
        (segunda + timedelta(days=3)).isoformat(): "qui",
        (segunda + timedelta(days=4)).isoformat(): "sex",
    }

    turma_semana: dict[str, dict[str, int]] = {
        turma: {"seg": 0, "ter": 0, "qua": 0, "qui": 0, "sex": 0, "total": 0} for turma in TURMAS
    }
    for row in turma_semana_rows:
        turma = row["turma"]
        dia = week_map.get(row["data_almoco"])
        if turma not in turma_semana or not dia:
            continue
        valor = row["sim"] or 0
        turma_semana[turma][dia] = valor

    quadro_importado_rows = conn.execute(
        """
        SELECT turma, data_almoco, sim
        FROM quadro_importado
        WHERE data_almoco BETWEEN ? AND ?
        """,
        (segunda.isoformat(), sexta.isoformat()),
    ).fetchall()

    for row in quadro_importado_rows:
        turma = row["turma"]
        dia = week_map.get(row["data_almoco"])
        if turma not in turma_semana or not dia:
            continue
        importado = max(0, int(row["sim"] or 0))
        turma_semana[turma][dia] = max(turma_semana[turma][dia], importado)

    for turma in TURMAS:
        item = turma_semana[turma]
        item["total"] = item["seg"] + item["ter"] + item["qua"] + item["qui"] + item["sex"]
        semana_sim["seg"] += item["seg"]
        semana_sim["ter"] += item["ter"]
        semana_sim["qua"] += item["qua"]
        semana_sim["qui"] += item["qui"]
        semana_sim["sex"] += item["sex"]

    quadro_rows: list[dict[str, int | str]] = []
    for idx, turma in enumerate(TURMAS_ORDEM_QUADRO, start=1):
        item = turma_semana.get(turma, {"seg": 0, "ter": 0, "qua": 0, "qui": 0, "sex": 0, "total": 0})
        quadro_rows.append(
            {
                "ordem": idx,
                "turma_nome": TURMAS_LABEL.get(turma, turma),
                "seg": item["seg"],
                "ter": item["ter"],
                "qua": item["qua"],
                "qui": item["qui"],
                "sex": item["sex"],
                "total": item["total"],
            }
        )

    total_semana_geral = sum(semana_sim.values())
    return semana_sim, quadro_rows, total_semana_geral


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


def as_clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_positive_int(value: object) -> int:
    text = as_clean_text(value).replace(",", ".")
    if not text:
        return 0
    try:
        return max(0, int(round(float(text))))
    except ValueError:
        return 0


def map_turma_value(value: str) -> str | None:
    normalized = normalize_header(value)
    if not normalized:
        return None

    for turma in TURMAS:
        if normalize_header(turma) == normalized:
            return turma

    for turma, turma_label in TURMAS_LABEL.items():
        if normalize_header(turma_label) == normalized:
            return turma

    if "servidor" in normalized:
        return "SERVIDORES"

    serie_match = re.search(r"([123])(?=[^0-9]|$)", normalized)
    if not serie_match:
        return None

    serie = serie_match.group(1)
    if "informatica" in normalized:
        return {"1": "TIN I", "2": "TIN II", "3": "TIN III"}.get(serie)
    if "automacao" in normalized:
        return {"1": "TAI I", "2": "TAI II", "3": "TAI III"}.get(serie)
    if "seguranca" in normalized and "trabalho" in normalized:
        return {"1": "TST I", "2": "TST II", "3": "TST III"}.get(serie)

    return None


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

    matricula = f"AUTO::{turma}::{nome}".upper()

    if not dias_marcados:
        return redirect(url_for("index", erro="Marque pelo menos um dia da semana."))
    if any(item not in DIAS_SEMANA for item in dias_marcados):
        return redirect(url_for("index", erro="Seleção de dias inválida."))

    if data_referencia:
        try:
            data_ref = parse_iso_date(data_referencia)
        except ValueError:
            return redirect(url_for("index", erro="Informe uma data válida."))
    else:
        data_ref = date.today()

    segunda = week_start(data_ref)
    datas_semana = {
        "seg": segunda,
        "ter": segunda + timedelta(days=1),
        "qua": segunda + timedelta(days=2),
        "qui": segunda + timedelta(days=3),
        "sex": segunda + timedelta(days=4),
    }

    with get_conn() as conn:
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

    write_backup_xlsx()

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
    periodo = request.args.get("periodo", "semana").strip().lower()
    if periodo not in {"semana", "mes", "ano"}:
        periodo = "semana"

    periodo_inicio, periodo_fim, periodo_label = period_bounds(data_base, periodo)

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

        respostas_semana_rows = conn.execute(
            """
            SELECT nome, matricula, turma, data_almoco, intencao
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            ORDER BY turma, nome, data_almoco
            """,
            (segunda.isoformat(), sexta.isoformat()),
        ).fetchall()

        relatorio_periodo_rows = conn.execute(
            """
            SELECT data_almoco,
                   SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END) AS sim,
                   SUM(CASE WHEN intencao = 'NAO' THEN 1 ELSE 0 END) AS nao
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            GROUP BY data_almoco
            ORDER BY data_almoco
            """,
            (periodo_inicio.isoformat(), periodo_fim.isoformat()),
        ).fetchall()

        total_semana_periodo = conn.execute(
            """
            SELECT COALESCE(SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END), 0) AS total
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            """,
            (segunda.isoformat(), sexta.isoformat()),
        ).fetchone()["total"]

        mes_inicio, mes_fim = month_bounds(data_base)
        total_mes_periodo = conn.execute(
            """
            SELECT COALESCE(SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END), 0) AS total
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            """,
            (mes_inicio.isoformat(), mes_fim.isoformat()),
        ).fetchone()["total"]

        ano_inicio, ano_fim = year_bounds(data_base)
        total_ano_periodo = conn.execute(
            """
            SELECT COALESCE(SUM(CASE WHEN intencao = 'SIM' THEN 1 ELSE 0 END), 0) AS total
            FROM respostas
            WHERE data_almoco BETWEEN ? AND ?
            """,
            (ano_inicio.isoformat(), ano_fim.isoformat()),
        ).fetchone()["total"]

        semana_sim, quadro_rows, total_semana_geral = build_quadro_semana(conn, segunda, sexta)

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

    respostas_por_pessoa: dict[str, dict[str, str | dict[str, bool]]] = {}
    week_map_respostas = {
        segunda.isoformat(): "seg",
        (segunda + timedelta(days=1)).isoformat(): "ter",
        (segunda + timedelta(days=2)).isoformat(): "qua",
        (segunda + timedelta(days=3)).isoformat(): "qui",
        (segunda + timedelta(days=4)).isoformat(): "sex",
    }

    for row in respostas_semana_rows:
        matricula = row["matricula"]
        if matricula not in respostas_por_pessoa:
            respostas_por_pessoa[matricula] = {
                "nome": row["nome"],
                "turma": row["turma"],
                "dias": {"seg": False, "ter": False, "qua": False, "qui": False, "sex": False},
            }

        dia = week_map_respostas.get(row["data_almoco"])
        if dia and row["intencao"] == "SIM":
            respostas_por_pessoa[matricula]["dias"][dia] = True

    dias_label = {"seg": "Seg", "ter": "Ter", "qua": "Qua", "qui": "Qui", "sex": "Sex"}
    respostas = []
    for item in sorted(respostas_por_pessoa.values(), key=lambda x: (x["turma"], x["nome"])):
        dias = item["dias"]
        checks = [f"{dias_label[dia]} ✅" for dia in DIAS_SEMANA if dias[dia]]
        respostas.append(
            {
                "nome": item["nome"],
                "turma": item["turma"],
                "intencao": " | ".join(checks) if checks else "Sem check na semana",
            }
        )

    relatorio_periodo = []
    total_periodo_sim = 0
    total_periodo_nao = 0
    for row in relatorio_periodo_rows:
        sim = row["sim"] or 0
        nao = row["nao"] or 0
        relatorio_periodo.append(
            {
                "data": row["data_almoco"],
                "sim": sim,
                "nao": nao,
                "total": sim + nao,
            }
        )
        total_periodo_sim += sim
        total_periodo_nao += nao

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
        importado_quadro=request.args.get("importado_quadro") == "1",
        import_quadro_error=request.args.get("import_quadro_error"),
        backup_restaurado=request.args.get("backup_restaurado") == "1",
        backup_restore_error=request.args.get("backup_restore_error"),
        backup_restore_file=request.args.get("backup_restore_file", ""),
        semana_sim=semana_sim,
        total_semana_geral=total_semana_geral,
        quadro_rows=quadro_rows,
        semana_inicio=segunda.isoformat(),
        semana_fim=sexta.isoformat(),
        periodo=periodo,
        periodo_label=periodo_label,
        periodo_inicio=periodo_inicio.isoformat(),
        periodo_fim=periodo_fim.isoformat(),
        relatorio_periodo=relatorio_periodo,
        total_periodo_sim=total_periodo_sim,
        total_periodo_nao=total_periodo_nao,
        total_semana_periodo=total_semana_periodo,
        total_mes_periodo=total_mes_periodo,
        total_ano_periodo=total_ano_periodo,
    )


@app.post("/admin/importar_alunos")
def importar_alunos():
    if not is_admin_allowed_form():
        abort(403, "Acesso negado. Informe um token válido.")

    file = request.files.get("arquivo_csv")
    token = request.form.get("token", "")
    data_filtro = request.form.get("data", "")

    if not file or not file.filename:
        return redirect(url_for("admin", token=token, data=data_filtro, import_error="Selecione um arquivo CSV ou XLSX."))

    filename = file.filename.lower()
    rows_for_import: list[dict[str, str]] = []

    try:
        if filename.endswith(".xlsx"):
            workbook = load_workbook(file.stream, data_only=True)
            sheet = workbook.active
            raw_rows = list(sheet.iter_rows(values_only=True))

            if not raw_rows:
                return redirect(url_for("admin", token=token, data=data_filtro, import_error="Planilha XLSX vazia."))

            header_original = [as_clean_text(col) for col in raw_rows[0]]
            if not any(header_original):
                return redirect(url_for("admin", token=token, data=data_filtro, import_error="XLSX sem cabeçalho."))

            header_normalized = [normalize_header(col) for col in header_original]
            for values in raw_rows[1:]:
                item = {header_normalized[i]: as_clean_text(values[i]) for i in range(min(len(header_normalized), len(values)))}
                rows_for_import.append(item)
        else:
            payload = file.stream.read().decode("utf-8-sig")
            sample = payload[:2048]
            dialect = csv.Sniffer().sniff(sample, delimiters=",;")
            reader = csv.DictReader(StringIO(payload), dialect=dialect)

            if not reader.fieldnames:
                return redirect(url_for("admin", token=token, data=data_filtro, import_error="CSV sem cabeçalho."))

            header_normalized = [normalize_header(item) for item in reader.fieldnames]
            for row in reader:
                item = {header_normalized[i]: as_clean_text(row.get(reader.fieldnames[i])) for i in range(len(header_normalized))}
                rows_for_import.append(item)
    except Exception:
        return redirect(url_for("admin", token=token, data=data_filtro, import_error="Não foi possível ler o arquivo (use CSV ou XLSX válido)."))

    if not rows_for_import:
        return redirect(url_for("admin", token=token, data=data_filtro, import_error="Arquivo sem dados para importar."))

    first_row_keys = set(rows_for_import[0].keys())
    nome_key = next((h for h in ["nome", "aluno", "nome completo"] if h in first_row_keys), None)
    matricula_key = next((h for h in ["matricula", "matricula aluno", "ra"] if h in first_row_keys), None)
    turma_key = next((h for h in ["turma", "serie", "classe"] if h in first_row_keys), None)

    if not nome_key or not matricula_key or not turma_key:
        return redirect(
            url_for(
                "admin",
                token=token,
                data=data_filtro,
                import_error="Arquivo precisa das colunas: nome, matricula e turma.",
            )
        )

    importados = 0
    with get_conn() as conn:
        for row in rows_for_import:
            nome = as_clean_text(row.get(nome_key))
            matricula = as_clean_text(row.get(matricula_key))
            turma = as_clean_text(row.get(turma_key))

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

    write_backup_xlsx()

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


@app.post("/admin/importar_quadro")
def importar_quadro_semanal():
    if not is_admin_allowed_form():
        abort(403, "Acesso negado. Informe um token válido.")

    file = request.files.get("arquivo_quadro")
    token = request.form.get("token", "")
    data_filtro = request.form.get("data", "")

    if not file or not file.filename:
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="Selecione um arquivo XLSX."))

    if not file.filename.lower().endswith(".xlsx"):
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="Use um arquivo .xlsx para o quadro semanal."))

    try:
        data_base = parse_iso_date(data_filtro) if data_filtro else date.today()
    except ValueError:
        data_base = date.today()

    segunda = week_start(data_base)
    semana_datas = {
        "seg": segunda,
        "ter": segunda + timedelta(days=1),
        "qua": segunda + timedelta(days=2),
        "qui": segunda + timedelta(days=3),
        "sex": segunda + timedelta(days=4),
    }

    try:
        workbook = load_workbook(file.stream, data_only=True)
        sheet = workbook.active
        rows = list(sheet.iter_rows(values_only=True))
    except Exception:
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="Não foi possível ler o XLSX do quadro semanal."))

    if not rows:
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="XLSX do quadro semanal está vazio."))

    required_cols = ["turma", "seg", "ter", "qua", "qui", "sex"]
    header_row_idx = None
    header: list[str] = []
    for idx, row in enumerate(rows[:15]):
        candidate = [normalize_header(as_clean_text(col)) for col in row]
        if all(col in candidate for col in required_cols):
            header_row_idx = idx
            header = candidate
            break

    if header_row_idx is None:
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="Cabeçalho não encontrado. Use colunas: turma, seg, ter, qua, qui, sex."))

    col_index: dict[str, int] = {}
    for col in required_cols:
        col_index[col] = header.index(col)

    quadro_import: dict[str, dict[str, int]] = {}
    for values in rows[header_row_idx + 1:]:
        turma_raw = as_clean_text(values[col_index["turma"]] if col_index["turma"] < len(values) else "")
        if turma_raw.isdigit() and (col_index["turma"] + 1) < len(values):
            turma_raw = as_clean_text(values[col_index["turma"] + 1])
        if not turma_raw:
            continue
        if normalize_header(turma_raw) == "total":
            continue

        turma = map_turma_value(turma_raw)
        if not turma:
            continue

        linha = {
            "seg": parse_positive_int(values[col_index["seg"]] if col_index["seg"] < len(values) else 0),
            "ter": parse_positive_int(values[col_index["ter"]] if col_index["ter"] < len(values) else 0),
            "qua": parse_positive_int(values[col_index["qua"]] if col_index["qua"] < len(values) else 0),
            "qui": parse_positive_int(values[col_index["qui"]] if col_index["qui"] < len(values) else 0),
            "sex": parse_positive_int(values[col_index["sex"]] if col_index["sex"] < len(values) else 0),
        }

        if turma not in quadro_import:
            quadro_import[turma] = linha
        else:
            quadro_import[turma]["seg"] += linha["seg"]
            quadro_import[turma]["ter"] += linha["ter"]
            quadro_import[turma]["qua"] += linha["qua"]
            quadro_import[turma]["qui"] += linha["qui"]
            quadro_import[turma]["sex"] += linha["sex"]

    if not quadro_import:
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="Nenhuma turma válida encontrada no XLSX do quadro."))

    total_importado = sum(
        dias["seg"] + dias["ter"] + dias["qua"] + dias["qui"] + dias["sex"] for dias in quadro_import.values()
    )
    if total_importado == 0:
        return redirect(url_for("admin", token=token, data=data_filtro, import_quadro_error="Arquivo importado resultou em total zero. Operação cancelada para evitar apagar o quadro."))

    sexta = segunda + timedelta(days=4)
    with get_conn() as conn:
        for turma, dias in quadro_import.items():
            for dia, valor in dias.items():
                conn.execute(
                    """
                    INSERT INTO quadro_importado (turma, data_almoco, sim)
                    VALUES (?, ?, ?)
                    ON CONFLICT(turma, data_almoco)
                    DO UPDATE SET
                        sim = excluded.sim,
                        atualizado_em = CURRENT_TIMESTAMP
                    """,
                    (turma, semana_datas[dia].isoformat(), valor),
                )
        conn.commit()

    write_backup_xlsx()

    return redirect(url_for("admin", token=token, data=segunda.isoformat(), importado_quadro=1))


@app.post("/admin/restaurar_backup")
def restaurar_backup_quadro():
    if not is_admin_allowed_form():
        abort(403, "Acesso negado. Informe um token válido.")

    token = request.form.get("token", "")
    data_filtro = request.form.get("data", "")

    try:
        data_base = parse_iso_date(data_filtro) if data_filtro else date.today()
    except ValueError:
        data_base = date.today()

    segunda = week_start(data_base)
    sexta = segunda + timedelta(days=4)

    backup_dir = DB_DIR / "backups"
    backup_files = sorted(
        backup_dir.glob("almoco_backup_*.xlsx"),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )
    if not backup_files:
        return redirect(
            url_for(
                "admin",
                token=token,
                data=segunda.isoformat(),
                backup_restore_error="Nenhum backup XLSX encontrado.",
            )
        )

    selected_backup = None
    linhas_restauracao: list[tuple[str, str, int]] = []

    for backup_file in backup_files:
        try:
            workbook = load_workbook(backup_file, data_only=True)
        except Exception:
            continue

        if "quadro_importado" not in workbook.sheetnames:
            continue

        sheet = workbook["quadro_importado"]
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue

        header = [normalize_header(as_clean_text(col)) for col in rows[0]]
        required = ["turma", "data_almoco", "sim"]
        if not all(col in header for col in required):
            continue

        idx_turma = header.index("turma")
        idx_data = header.index("data_almoco")
        idx_sim = header.index("sim")

        candidate_rows: list[tuple[str, str, int]] = []
        for values in rows[1:]:
            turma_raw = as_clean_text(values[idx_turma] if idx_turma < len(values) else "")
            data_raw = values[idx_data] if idx_data < len(values) else ""
            sim_raw = values[idx_sim] if idx_sim < len(values) else 0

            turma = map_turma_value(turma_raw)
            if not turma:
                continue

            if isinstance(data_raw, datetime):
                data_row = data_raw.date()
            elif isinstance(data_raw, date):
                data_row = data_raw
            else:
                data_text = as_clean_text(data_raw)
                try:
                    data_row = parse_iso_date(data_text)
                except ValueError:
                    continue

            if data_row < segunda or data_row > sexta:
                continue

            candidate_rows.append((turma, data_row.isoformat(), parse_positive_int(sim_raw)))

        if candidate_rows:
            selected_backup = backup_file
            linhas_restauracao = candidate_rows
            break

    if not linhas_restauracao or selected_backup is None:
        return redirect(
            url_for(
                "admin",
                token=token,
                data=segunda.isoformat(),
                backup_restore_error="Nenhum backup possui linhas da semana selecionada.",
            )
        )

    with get_conn() as conn:
        for turma, data_almoco, sim in linhas_restauracao:
            conn.execute(
                """
                INSERT INTO quadro_importado (turma, data_almoco, sim)
                VALUES (?, ?, ?)
                ON CONFLICT(turma, data_almoco)
                DO UPDATE SET
                    sim = excluded.sim,
                    atualizado_em = CURRENT_TIMESTAMP
                """,
                (turma, data_almoco, sim),
            )
        conn.commit()

    write_backup_xlsx()

    return redirect(
        url_for(
            "admin",
            token=token,
            data=segunda.isoformat(),
            backup_restaurado=1,
            backup_restore_file=selected_backup.name,
        )
    )


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


@app.get("/export_quadro.csv")
def export_quadro_csv() -> Response:
    if not is_admin_allowed():
        abort(403, "Acesso negado. Informe um token válido na URL.")

    data_filtro = request.args.get("data") or date.today().isoformat()
    try:
        data_base = parse_iso_date(data_filtro)
    except ValueError:
        data_base = date.today()

    segunda = week_start(data_base)
    sexta = segunda + timedelta(days=4)

    with get_conn() as conn:
        semana_sim, quadro_rows, total_semana_geral = build_quadro_semana(conn, segunda, sexta)

    output = StringIO()
    writer = csv.writer(output, delimiter=';')
    writer.writerow(["#", "Turma", "Seg", "Ter", "Qua", "Qui", "Sex", "Total"])

    for row in quadro_rows:
        writer.writerow(
            [
                row["ordem"],
                row["turma_nome"],
                row["seg"],
                row["ter"],
                row["qua"],
                row["qui"],
                row["sex"],
                row["total"],
            ]
        )

    writer.writerow(["", "Total", semana_sim["seg"], semana_sim["ter"], semana_sim["qua"], semana_sim["qui"], semana_sim["sex"], total_semana_geral])

    csv_data = output.getvalue()
    output.close()

    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=quadro_semanal_{segunda.isoformat()}_{sexta.isoformat()}.csv"},
    )


@app.get("/export_quadro.xlsx")
def export_quadro_xlsx() -> Response:
    if not is_admin_allowed():
        abort(403, "Acesso negado. Informe um token válido na URL.")

    data_filtro = request.args.get("data") or date.today().isoformat()
    try:
        data_base = parse_iso_date(data_filtro)
    except ValueError:
        data_base = date.today()

    segunda = week_start(data_base)
    sexta = segunda + timedelta(days=4)

    with get_conn() as conn:
        semana_sim, quadro_rows, total_semana_geral = build_quadro_semana(conn, segunda, sexta)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "quadro_semanal"
    sheet.append(["#", "Turma", "Seg", "Ter", "Qua", "Qui", "Sex", "Total"])

    for row in quadro_rows:
        sheet.append(
            [
                row["ordem"],
                row["turma_nome"],
                row["seg"],
                row["ter"],
                row["qua"],
                row["qui"],
                row["sex"],
                row["total"],
            ]
        )

    sheet.append(["", "Total", semana_sim["seg"], semana_sim["ter"], semana_sim["qua"], semana_sim["qui"], semana_sim["sex"], total_semana_geral])

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    return Response(
        buffer.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=quadro_semanal_{segunda.isoformat()}_{sexta.isoformat()}.xlsx"},
    )


init_db()

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "0") == "1",
    )
