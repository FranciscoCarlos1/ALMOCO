from io import BytesIO

from flask import Blueprint, render_template, request, redirect, url_for, jsonify, abort, send_file
from datetime import date, datetime, timedelta
from db import get_conn

bp_main = Blueprint('main', __name__)

TURMAS = [
    "TIN I", "TIN II", "TIN III",
    "TAI I", "TAI II", "TAI III",
    "TST I", "TST II", "TST III", "SERVIDORES"
]
INTENCOES = ["SIM", "NAO"]
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex"]


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def week_start(given_date: date) -> date:
    return given_date - timedelta(days=given_date.weekday())

@bp_main.route("/")
def index():
    sucesso = request.args.get("sucesso") == "1"
    erro = request.args.get("erro")
    hoje = date.today().isoformat()
    with get_conn() as conn:
        cardapio = conn.execute(
            """
            SELECT descricao, imagem_blob
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (hoje,),
        ).fetchone()
    return render_template(
        "index.html",
        turmas=TURMAS,
        intencoes=INTENCOES,
        sucesso=sucesso,
        erro=erro,
        hoje=hoje,
        cardapio_hoje=cardapio["descricao"] if cardapio else None,
        cardapio_imagem=hoje if cardapio and cardapio["imagem_blob"] else None,
    )


@bp_main.route("/cardapio/imagens/<path:nome_arquivo>")
def cardapio_imagem(nome_arquivo: str):
    with get_conn() as conn:
        cardapio = conn.execute(
            """
            SELECT imagem_blob, imagem_mime
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (nome_arquivo,),
        ).fetchone()

    if not cardapio or not cardapio["imagem_blob"]:
        abort(404)

    imagem_blob = cardapio["imagem_blob"]
    if isinstance(imagem_blob, memoryview):
        imagem_blob = imagem_blob.tobytes()

    return send_file(
        BytesIO(imagem_blob),
        mimetype=cardapio["imagem_mime"] or "application/octet-stream",
    )

@bp_main.route("/aluno")
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
    return jsonify({
        "ok": True,
        "nome": aluno["nome"],
        "matricula": aluno["matricula"],
        "turma": aluno["turma"],
    })


@bp_main.post("/enviar")
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
        return redirect(url_for("main.index", erro="Informe seu nome."))
    if turma not in TURMAS:
        return redirect(url_for("main.index", erro="Selecione uma turma válida."))
    if not dias_marcados:
        return redirect(url_for("main.index", erro="Marque pelo menos um dia da semana."))
    if any(item not in DIAS_SEMANA for item in dias_marcados):
        return redirect(url_for("main.index", erro="Seleção de dias inválida."))

    if data_referencia:
        try:
            data_ref = parse_iso_date(data_referencia)
        except ValueError:
            return redirect(url_for("main.index", erro="Informe uma data válida."))
    else:
        data_ref = date.today()

    matricula = f"AUTO::{turma}::{nome}".upper()
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

    return redirect(url_for("main.index", sucesso=1))


@bp_main.get("/enviar")
def enviar_redirect():
    return redirect(url_for("main.index"))
